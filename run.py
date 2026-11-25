"""Run the local workpaper pipeline or delegate to the safe profiling CLI."""

import argparse
from contextlib import redirect_stderr, redirect_stdout
from importlib import import_module
import json
import os
from pathlib import Path
import re
import sys
from tempfile import TemporaryDirectory
from uuid import uuid4
import warnings

from adapters.base_adapter import BaseAdapter, source_totals
from core.blueprint import load_blueprint
from core.cas_checker import scan_routed, scan_tb, write_cas_sheet
from core.cas_routing import confirmed_account_names
from core.fill_workpaper import select_coverage, workpaper_writer
from core.masker import main as probe_main, safe_path
from core.recalc import Calculation, read_caches, recalculate, scan
from core.reconcile import console_lines, run_gates, write_report
from core.schema import AuditError, account_code, validate

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
    parser = Parser(description="在本地生成应收账款审计底稿")
    parser.add_argument("--adapter", required=True, help="adapters 下已审核的模块名")
    parser.add_argument("--tb", required=True, type=Path)
    parser.add_argument("--gl", required=True, type=Path)
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


def generate(args):
    blueprint = load_blueprint()
    account_names = load_account_names(args.account_names) if args.account_names is not None else None
    tb, tb_kind = safe_path(args.tb)
    gl, gl_kind = safe_path(args.gl)
    if tb_kind != gl_kind or tb.parent != gl.parent or tb.samefile(gl):
        raise AuditError("TB 与 GL 须为同一数据目录中的两个独立文件。")
    with open(os.devnull, "w", encoding="utf-8") as quiet, redirect_stdout(quiet), redirect_stderr(quiet):
        adapter = load_adapter(args.adapter)
        source = source_totals(tb, adapter.tb_columns, adapter.tb_source, adapter.blank_tb_amounts_are_zero)
        data = adapter.transform(tb, gl)
    validate(data)
    try:
        ar_account = account_code(args.ar_account if args.ar_account is not None else adapter.ar_account)
    except ValueError:
        raise AuditError("应收一级科目编码无效。") from None
    coverage = select_coverage(data, ar_account)
    name = args.out or f"ar_workpaper_{uuid4().hex[:16]}.xlsx"
    if not re.fullmatch(r"[A-Za-z0-9][A-Za-z0-9_-]{0,80}\.xlsx", name):
        raise AuditError("输出名称须为不含路径的 ASCII XLSX 文件名。")
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
                if args.engine == "xlwings" and args.recalc in ("auto", "xlwings"):
                    calculation = native_calculation(writer.stage, actual)
                else:
                    calculation = recalculate(writer.stage, actual, args.recalc, args.timeout)
                gates, controls, balances = run_gates(data, source, coverage, layout, actual, calculation)
                book = writer.book
                write_report(book, gates, controls, balances, calculation)
                if args.cas_routed:
                    account_codes = tuple(dict.fromkeys(row.account_code for row in data.tb))
                    result = scan_routed(data, account_codes, account_names=account_names,
                                         complete_vouchers=args.complete_vouchers, period_closed=args.period_closed,
                                         currency_exchange_closed=args.currency_closed,
                                         account_pairs=adapter.cas_allowance_pairs)
                    write_cas_sheet(book, result.rules, result.findings, deferred_rule_ids=result.deferred_rule_ids)
                else:
                    rules, findings = scan_tb(data.tb, ar_account, account_pairs=adapter.cas_allowance_pairs)
                    write_cas_sheet(book, rules, findings)
                passed = all(gate.status == "PASS" for gate in gates)
                final = writer.finish(calculation)
            if passed and args.engine == "xlwings":
                saved_calculation = native_calculation(final, actual)
                saved_gates, _, saved_balances = run_gates(data, source, coverage, layout, actual, saved_calculation)
                if saved_balances != balances or any(gate.status != "PASS" for gate in saved_gates):
                    raise AuditError("最终 Excel 工作簿的公式缓存与勾稽报告不一致。")
            destination = approved if passed else review
            if os.name == "nt":
                final.rename(destination)
            else:
                os.link(final, destination)
    finally:
        claim.rmdir()
    for line in console_lines(gates, controls, balances, show_amounts=tb_kind == "mock" or args.local_amounts):
        print(line)
    print(f"data/output/{destination.name}")
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
