"""Run the local workpaper pipeline or delegate to the safe profiling CLI."""

import argparse
from contextlib import redirect_stderr, redirect_stdout
from importlib import import_module
import os
from pathlib import Path
import re
import sys
from tempfile import TemporaryDirectory
from uuid import uuid4
import warnings

from adapters.base_adapter import BaseAdapter, source_totals
from core.cas_checker import scan_tb, write_cas_sheet
from core.fill_workpaper import build, select_coverage
from core.masker import main as probe_main, safe_path
from core.recalc import formulas, recalculate, write_caches
from core.reconcile import console_lines, run_gates, write_report
from core.schema import AuditError, account_code, validate

ROOT = Path(__file__).resolve().parent
TEMPLATE = ROOT / "templates" / "ar_template.xlsx"


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
    parser.add_argument("--recalc", choices=("auto", "libreoffice", "xlwings"), default="auto")
    parser.add_argument("--timeout", type=positive_integer, default=120, help="每个重算后端的超时秒数")
    parser.add_argument("--out", help="data/output 下的 ASCII XLSX 文件名")
    parser.add_argument("--local-amounts", action="store_true", help="仅供本地数据持有人输出真实输入的金额证据")
    return parser.parse_args(argv)


def store(name):
    path = ROOT / "data" / name
    if not path.is_dir() or path.resolve() != path:
        raise AuditError("项目数据目录缺失或已重定向。")
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


def generate(args):
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
            scratch = Path(temp)
            book, layout = build(TEMPLATE, data, coverage, ar_account)
            try:
                actual = formulas(book)
                stage = scratch / "workpaper.xlsx"
                book.save(stage)
                calculation = recalculate(stage, actual, args.recalc, args.timeout)
                gates, controls, balances = run_gates(data, source, coverage, layout, actual, calculation)
                write_report(book, gates, controls, balances, calculation)
                rules, findings = scan_tb(data.tb, ar_account, account_pairs=adapter.cas_allowance_pairs)
                write_cas_sheet(book, rules, findings)
                passed = all(gate.status == "PASS" for gate in gates)
                book.save(stage)
            finally:
                book.close()
            final = stage
            if calculation.caches:
                final = scratch / "cached.xlsx"
                write_caches(stage, final, calculation.caches)
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
