"""Run the local workpaper pipeline or delegate to the safe profiling CLI."""

import argparse
from contextlib import closing, redirect_stderr, redirect_stdout
from dataclasses import replace
from importlib import import_module
import json
import os
from pathlib import Path
import re
import sys
from tempfile import TemporaryDirectory
from urllib.parse import quote
from uuid import uuid4
import warnings

from openpyxl import Workbook

from adapters.base_adapter import BaseAdapter, source_totals
from core.accounts import Ledger
from core.blueprint import load_blueprint, load_blueprints, matches_code
from core.cas_checker import opening_suggestions, scan_counterparties, scan_routed, scan_tb, write_cas_sheet
from core.cas_routing import confirmed_account_names
from core.excel_writer import ExcelWriter
from core.fill_workpaper import select_coverage, workpaper_writer
from core.fillers import make_filler
from core.masker import main as probe_main, safe_path
from core.master_tb import MASTER_FILE, STAGES, build_master, check_master
from core.recalc import Calculation, formulas, read_caches, recalculate, scan, write_caches
from core.reconcile import (
    BATCH_COMPARISON_HEADERS, cached_money, check_batch_coverage, check_conservation, check_excel, check_opening, check_subject,
    combine_gates, console_lines, gates_passed, money, opening_differences, run_gates, write_report,
)
from core.schema import AuditError, GL_SHEET, TB_SHEET, account_code, validate

ROOT = Path(__file__).resolve().parent


class Parser(argparse.ArgumentParser):
    def error(self, message):
        self.exit(2, "FAILED\n命令参数无效。请使用 --help 查看用法。\n")


def positive_integer(value):
    number = int(value)
    if number < 1:
        raise argparse.ArgumentTypeError("expected a positive integer")
    return number


def parse_args(argv):
    parser = Parser(description="在本地生成实质性底稿及总试算平衡表")
    parser.add_argument("--adapter", required=True, help="adapters 下已审核的模块名")
    parser.add_argument("--tb", required=True, type=Path)
    parser.add_argument("--gl", required=True, type=Path)
    parser.add_argument("--prior-tb", type=Path, help="可选的上年审定期末余额表")
    parser.add_argument("--subjects", help="指定科目编码、用逗号分隔的多个编码，或 ALL")
    parser.add_argument("--ar-account", help="应收账款的一级科目编码，默认取 Adapter")
    parser.add_argument("--engine", choices=("openpyxl", "xlwings"), default="openpyxl",
                        help="底稿写入引擎，默认 openpyxl")
    parser.add_argument("--recalc", choices=("auto", "libreoffice", "xlwings"), default="auto")
    parser.add_argument("--timeout", type=positive_integer, default=120, help="每次 Excel 原生写入或重算的超时秒数")
    parser.add_argument("--out", help="data/output 下的 ASCII XLSX 文件名")
    parser.add_argument("--local-amounts", action="store_true", help="仅供本地数据持有人输出真实输入的金额证据")
    parser.add_argument("--cas-routed", action="store_true", help="启用按科目路由的 CAS 筛查")
    parser.add_argument("--complete-vouchers", action="store_true", help="确认凭证已核对完整")
    parser.add_argument("--period-closed", action="store_true", help="确认相关损益结转已完成")
    parser.add_argument("--currency-closed", action="store_true",
                        help="确认外币折算及汇兑结转已完成，且 TB 已统一为记账本位币")
    parser.add_argument("--account-names", type=Path, metavar="FILE", help="已核实科目编码到名称的 JSON 映射")
    return parser.parse_args(argv)


def store(name):
    path = ROOT / "data" / name
    if path.resolve() != path:
        raise AuditError("出于数据安全，项目 data/ 下的数据目录不允许使用符号链接或重定向。")
    if not path.is_dir():
        raise AuditError("项目数据目录缺失。")
    return path


def load_adapter(name):
    if not re.fullmatch(r"[a-z][a-z0-9_]*", name):
        raise AuditError("适配器名称无效。")
    if not (ROOT / "adapters" / f"{name}.py").is_file():
        raise AuditError("找不到指定适配器。")
    adapter_type = getattr(import_module(f"adapters.{name}"), "Adapter", None)
    if not isinstance(adapter_type, type) or not issubclass(adapter_type, BaseAdapter):
        raise AuditError("适配器须导出 BaseAdapter 的 Adapter 子类。")
    return adapter_type()


def load_account_names(path):
    try:
        with path.open(encoding="utf-8-sig") as file:
            names = json.load(file)
    except (OSError, ValueError, RecursionError):
        raise AuditError("无法读取 CAS 科目名称映射。请检查文件是否为有效的 UTF-8 JSON。") from None
    if not isinstance(names, dict):
        raise AuditError("CAS 科目名称映射须为科目编码到已核实名称的 JSON 对象。")
    return confirmed_account_names(names)


def native_calculation(path, actual):
    caches = read_caches(path, actual)
    return Calculation("xlwings", caches, scan(actual, caches))


def calculation_for(args, stage, actual, *, native_writer=False):
    if native_writer and args.recalc in ("auto", "xlwings"):
        return native_calculation(stage, actual)
    return recalculate(stage, actual, args.recalc, args.timeout)


def cas_scan(args, data, account_codes, ar_account, account_names, adapter, exclude_reclass=()):
    if args.cas_routed:
        result = scan_routed(data, account_codes, account_names=account_names,
                             complete_vouchers=args.complete_vouchers, period_closed=args.period_closed,
                             currency_exchange_closed=args.currency_closed, account_pairs=adapter.cas_allowance_pairs)
        return result.rules, result.findings, result.deferred_rule_ids
    rules, findings = scan_tb(data.tb, ar_account, account_pairs=adapter.cas_allowance_pairs,
                              exclude_reclass=exclude_reclass)
    return rules, findings, ()


def generate(args):
    account_names = load_account_names(args.account_names) if args.account_names is not None else None
    tb, tb_kind = safe_path(args.tb)
    gl, gl_kind = safe_path(args.gl)
    if tb_kind != gl_kind or tb.parent != gl.parent or tb.samefile(gl):
        raise AuditError("TB 与 GL 须为同一数据目录中的两个独立文件。")
    prior = None
    if args.prior_tb is not None:
        prior, prior_kind = safe_path(args.prior_tb)
        if prior_kind != tb_kind or any(prior.samefile(path) for path in (tb, gl)):
            raise AuditError("上年审定数须为同类数据目录中的独立文件。")
    with open(os.devnull, "w", encoding="utf-8") as quiet, redirect_stdout(quiet), redirect_stderr(quiet):
        adapter = load_adapter(args.adapter)
        source = source_totals(tb, adapter.tb_columns, adapter.tb_source, adapter.blank_tb_amounts_are_zero)
        data = adapter.transform(tb, gl)
        prior_rows = adapter.convert_prior_tb(prior) if prior is not None else None
        if prior is not None and not isinstance(prior_rows, tuple):
            raise AuditError("convert_prior_tb() 须返回 Standard_Prior_TB 记录元组。")
        data = replace(data, prior_tb=prior_rows)
    validate(data)
    try:
        ar_account = account_code(args.ar_account if args.ar_account is not None else adapter.ar_account)
    except ValueError:
        raise AuditError("应收一级科目编码无效。") from None
    if args.out and not re.fullmatch(r"[A-Za-z0-9][A-Za-z0-9_-]{0,80}\.xlsx", args.out):
        raise AuditError("输出名称须为不含路径的 ASCII XLSX 文件名。")
    if args.out and args.out.casefold() == MASTER_FILE.casefold():
        raise AuditError("输出名称与 Master TB 保留文件名冲突。")
    show_amounts = tb_kind == "mock" or args.local_amounts
    if args.subjects is not None:
        return generate_batch(args, data, source, ar_account, account_names, adapter, show_amounts)
    return generate_legacy(args, data, source, ar_account, account_names, adapter, show_amounts)


def generate_legacy(args, data, source, ar_account, account_names, adapter, show_amounts):
    blueprint = load_blueprint()
    coverage = select_coverage(data, ar_account)
    name = args.out or f"ar_workpaper_{uuid4().hex[:16]}.xlsx"
    differences = opening_differences(Ledger(data)) if data.prior_tb is not None else ()
    opening = check_opening(data, differences)
    output = store("output")
    approved = output / name
    review = approved.with_suffix(".review.xlsx")
    claim = output / f".claim-{approved.stem.lower()}"
    try:
        claim.mkdir()
    except FileExistsError:
        raise AuditError("同名底稿正在生成。") from None
    try:
        if any(os.path.lexists(path) for path in (approved, review)):
            raise AuditError("同名底稿或待复核底稿已存在。")
        with TemporaryDirectory(prefix=".run-", dir=output) as temp:
            with workpaper_writer(args.engine, blueprint, Path(temp), args.timeout) as writer:
                layout, actual = writer.populate(data, coverage, ar_account)
                calculation = calculation_for(args, writer.stage, actual, native_writer=args.engine == "xlwings")
                gates, controls, balances = run_gates(data, source, coverage, layout, actual, calculation)
                gates = (*gates, opening)
                book = writer.book
                write_report(book, gates, controls, balances, calculation)
                codes = tuple(dict.fromkeys(row.account_code for row in data.tb))
                rules, findings, deferred = cas_scan(args, data, codes, ar_account, account_names, adapter)
                roll_rules, roll_findings = opening_suggestions(differences)
                write_cas_sheet(book, {**rules, **roll_rules}, (*findings, *roll_findings), deferred_rule_ids=deferred)
                passed = gates_passed(gates)
                final = writer.finish(calculation)
            if passed and args.engine == "xlwings":
                saved_calculation = native_calculation(final, actual)
                saved_gates, _, saved_balances = run_gates(data, source, coverage, layout, actual, saved_calculation)
                if saved_balances != balances or not gates_passed(saved_gates):
                    raise AuditError("最终 Excel 工作簿的公式缓存与勾稽报告不一致。")
            destination = approved if passed else review
            if os.name == "nt":
                final.rename(destination)
            else:
                os.link(final, destination)
    finally:
        claim.rmdir()
    for line in console_lines(gates, controls, balances, show_amounts=show_amounts):
        print(line)
    print(f"data/output/{destination.name}")
    return 0 if passed else 1


def select_subjects(value, blueprints, ledger):
    value = value.strip().upper()
    if value == "ALL":
        requested = tuple(blueprints)
    else:
        try:
            requested = tuple(dict.fromkeys(account_code(entry) for entry in value.split(",")))
        except ValueError:
            raise AuditError("--subjects 须为科目编码、逗号分隔编码或 ALL。") from None
        if any(code not in blueprints for code in requested):
            raise AuditError("--subjects 包含蓝图未配置的科目。")
    selected, owners, skipped = {}, {}, []
    for code in requested:
        blueprint = blueprints[code]
        accounts = tuple(root for root in ledger.current.roots if matches_code(root, blueprint.account_codes))
        if not accounts:
            if value != "ALL":
                raise AuditError("指定科目在本期 TB 中没有匹配科目。")
            skipped.append(code)
            continue
        if owners.keys() & set(accounts):
            raise AuditError("本次科目蓝图范围重叠。请检查 account_codes。")
        selected[code] = blueprint, accounts
        owners.update(dict.fromkeys(accounts, code))
    if not selected:
        raise AuditError("本期 TB 没有与本次底稿蓝图匹配的科目。")
    return selected, owners, skipped


def scoped_findings(data, ledger, codes, findings):
    for finding in findings:
        row = re.match(rf"^({TB_SHEET}|{GL_SHEET})!([0-9]+)", finding.position)
        if row:
            records = data.tb if row[1] == TB_SHEET else data.gl
            if ledger.owner(records[int(row[2]) - 2].account_code) not in codes:
                continue
        yield finding


def publish_batch(output, staged, passed, scratch):
    names = [(source, Path(name) if passed else Path(name).with_suffix(".review.xlsx")) for source, name in staged]
    if any(os.path.lexists(output / name) for _, name in names):
        folder = scratch / "publish"
        folder.mkdir()
        for source, name in names:
            source.rename(folder / name)
        destination = output / f"run_{uuid4().hex[:16]}"
        folder.rename(destination)
        return [destination / name for _, name in names]

    published = []
    try:
        for source, name in names:
            destination = output / name
            if os.name == "nt":
                source.rename(destination)
            else:
                os.link(source, destination)
            published.append(destination)
    except BaseException:
        for destination in published:
            destination.unlink()
        raise
    return published


def generate_batch(args, data, source, ar_account, account_names, adapter, show_amounts):
    ledger = Ledger(data)
    selected, owners, skipped = select_subjects(args.subjects, load_blueprints(), ledger)
    if args.out and len(selected) != 1:
        raise AuditError("--out 只可为单份科目底稿指定文件名。")
    output = store("output")
    claim = output / ".claim-00_master_tb"
    conservation, controls = check_conservation(data, *source)
    coverage = check_batch_coverage(data, ledger, owners)
    differences = opening_differences(ledger)
    opening = check_opening(data, differences)
    roll_rules, roll_findings = opening_suggestions(differences)
    counterparty_rows = frozenset(f"{TB_SHEET}!{line}" for blueprint, codes in selected.values()
                                  if blueprint.archetype == "counterparty_tb"
                                  for code in codes for line, _ in ledger.members[code])
    route_codes = tuple(dict.fromkeys((*owners, *(row.account_code for row in data.tb
                                                if ledger.owner(row.account_code) in owners))))
    try:
        claim.mkdir()
    except FileExistsError:
        raise AuditError("另一个批量底稿任务正在生成。") from None
    try:
        rules, findings, deferred = cas_scan(args, data, route_codes, ar_account, account_names, adapter, counterparty_rows)
        results, staged, group_findings = [], [], []
        with TemporaryDirectory(prefix=".run-", dir=output) as directory:
            scratch = Path(directory)
            for index, (code, (blueprint, accounts)) in enumerate(selected.items(), 1):
                subject_scratch = scratch / f"subject-{index}"
                subject_scratch.mkdir()
                with ExcelWriter(args.engine, blueprint, subject_scratch, args.timeout) as writer:
                    layout = make_filler(writer, ledger, accounts).populate()
                    actual = writer.apply()
                    calculation = calculation_for(args, writer.stage, actual, native_writer=args.engine == "xlwings")
                    cross, balances = check_subject(layout, calculation)
                    health = check_excel(layout, actual, calculation)
                    local_differences = tuple(item for item in differences if ledger.owner(item.code) in accounts)
                    local_opening = check_opening(data, local_differences)
                    gates = conservation, coverage, cross, health, local_opening
                    extra_rules, extra_findings = scan_counterparties(layout.counterparty_groups, ar_account,
                                                                      quantify=not args.cas_routed)
                    group_findings.extend(extra_findings)
                    local_roll_rules, local_roll = opening_suggestions(local_differences)
                    local_findings = tuple(scoped_findings(data, ledger, accounts, findings))
                    write_report(writer.book, gates, controls, balances, calculation, comparison_headers=BATCH_COMPARISON_HEADERS)
                    write_cas_sheet(writer.book, {**rules, **extra_rules, **local_roll_rules},
                                    (*local_findings, *extra_findings, *local_roll), deferred_rule_ids=deferred)
                    final = writer.finish(calculation)
                    if args.engine == "xlwings":
                        saved = native_calculation(final, actual)
                        if check_subject(layout, saved) != (cross, balances) or check_excel(layout, actual, saved) != health:
                            raise AuditError("最终 Excel 工作簿的公式缓存与勾稽报告不一致。")
                staged.append((final, args.out or f"{quote(code, safe='')}_workpaper.xlsx"))
                results.append((code, blueprint.archetype, cross, health, local_opening, balances))

            with closing(Workbook()) as book:
                active_findings = tuple(scoped_findings(data, ledger, owners, findings))
                master_layout = build_master(book, ledger, owners, (*active_findings, *group_findings, *roll_findings),
                                             account_names)
                actual = formulas(book)
                master_stage = scratch / MASTER_FILE
                book.save(master_stage)
                master_calculation = calculation_for(args, master_stage, actual)
                master = check_master(master_layout, master_calculation)
                master_health = check_excel(master_layout, actual, master_calculation)
                cross = combine_gates([*(item[2] for item in results), master])
                health = combine_gates([*(item[3] for item in results), master_health])
                gates = conservation, coverage, cross, health, opening
                write_report(book, gates, controls, None, master_calculation)
                write_cas_sheet(book, roll_rules, roll_findings)
                book.save(master_stage)
            cached_master = scratch / "master-cached.xlsx"
            write_caches(master_stage, cached_master, master_calculation.caches)
            staged.append((cached_master, MASTER_FILE))
            passed = gates_passed(gates)
            destinations = publish_batch(output, staged, passed, scratch)
    finally:
        claim.rmdir()

    lines = console_lines(gates, controls, None, show_amounts=show_amounts)
    for line in lines[:-1]:
        print(line)
    if show_amounts:
        for position, issue, difference, _ in (*cross.findings, *health.findings):
            print(f"  {position} | {issue}" + (f" | diff={money(difference)}" if difference is not None else ""))
        for finding in roll_findings:
            print(f"  期初滚调建议 | 借={finding.debit} | 贷={finding.credit} | 金额={money(finding.draft_amount)} | 待复核")
    for code in skipped:
        print(f"SUBJECT {code} SKIPPED | 本期 TB 无匹配科目")
    for code, archetype, cross, health, local_opening, balances in results:
        line = f"SUBJECT {code} {archetype} | gate3={cross.status} gate4={health.status} gate5={local_opening.status}"
        if show_amounts:
            summary, detail, raw = balances
            diff = detail - raw if detail is not None else None
            line += f" | summary={money(summary)} detail={money(detail)} raw={money(raw)} diff={money(diff)}"
        print(line)
    print(f"MASTER TB {master.status} | 公式健康度={master_health.status} | backend={master_calculation.backend}")
    if show_amounts:
        for stage, check in zip(STAGES, master_layout.differences, strict=True):
            cache = master_calculation.caches.get(check.position)
            diff = "SKIPPED" if check.expected == "" and cache and cache.value == "" else money(cached_money(cache))
            print(f"  {stage}借贷差额={diff}")
    print(lines[-1])
    for path in destinations:
        print(f"data/output/{path.relative_to(output).as_posix()}")
    return 0 if passed else 1


def main(argv=None):
    argv = sys.argv[1:] if argv is None else argv
    if argv and argv[0] == "profile":
        return probe_main(argv[1:])
    args = parse_args(argv)
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        try:
            return generate(args)
        except AuditError as exc:
            print(str(exc), file=sys.stderr)
        except Exception:
            print("运行异常。请在本地检查输入格式、适配器、模板和输出权限。", file=sys.stderr)
    print("FAILED")
    return 2


if __name__ == "__main__":
    raise SystemExit(main())
