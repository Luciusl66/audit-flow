"""Run native Excel writes in a local, hidden worker."""

from importlib.util import find_spec
from itertools import islice
import json
import os
from pathlib import Path
import subprocess
import sys

from core.schema import AuditError

ROOT = Path(__file__).resolve().parents[1]
EXCEL_UNAVAILABLE = 3
BATCH_ROWS = 2_000


def require_excel_writer():
    if os.name != "nt" or find_spec("xlwings") is None:
        raise AuditError("xlwings 写入需要 Windows 本机 Microsoft Excel 和 xlwings。请使用 --engine openpyxl。")


def run_excel_writer(operation, source, payload, destination, timeout, sheet_order=()):
    stage = {"populate": "填充模板", "attach_reports": "插入报告"}[operation]
    command = [sys.executable, "-m", "core.excel_writer", operation,
               str(source), str(payload), str(destination), *sheet_order]
    scratch = str(destination.parent)
    env = {**os.environ, "TMP": scratch, "TEMP": scratch, "TMPDIR": scratch}
    try:
        result = subprocess.run(command, cwd=ROOT, env=env, timeout=timeout, stdin=subprocess.DEVNULL,
                                stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    except subprocess.TimeoutExpired:
        raise AuditError(f"本机 Excel {stage}超时。请检查 Excel 状态，或选用 --engine openpyxl。") from None
    except OSError:
        raise AuditError("无法启动本机 Excel 写入进程。请检查本机环境。") from None
    if result.returncode == EXCEL_UNAVAILABLE:
        raise AuditError("无法启动本机 Microsoft Excel。请检查 Excel 和 xlwings，或选用 --engine openpyxl。")
    if result.returncode or not destination.is_file():
        raise AuditError(f"本机 Excel {stage}失败。请检查模板和 Excel 状态。")


def batches(stream, first, count):
    for offset in range(0, count, BATCH_ROWS):
        size = min(BATCH_ROWS, count - offset)
        rows = [json.loads(line) for line in islice(stream, size)]
        if len(rows) != size:
            raise ValueError("Incomplete write data.")
        yield first + offset, rows


def excel_literal(value):
    # An explicit text prefix keeps formula-like values and leading zeros literal.
    if not isinstance(value, str):
        return value
    if len(value) > 16_383 and len(value.encode("utf-16-le")) >= 2 * 32_767:
        raise AuditError("文本超出 Excel 原生写入可完整保留的长度。请使用 --engine openpyxl。")
    return "'" + value if value else None


def table(book, spec, stream):
    sheet = book.sheets[spec["sheet"]]
    first, count = spec["first"], spec["count"]
    last = first + count - 1
    width = len(spec["formats"])
    if count:
        for column, fmt in enumerate(spec["formats"], 1):
            sheet.range((first, column), (last, column)).number_format = fmt
        for row, values in batches(stream, first, count):
            body = sheet.range((row, 1), (row + len(values) - 1, width))
            body.api.Value2 = tuple(tuple(record) for record in values)
        sheet.activate()
        if sheet.api.AutoFilterMode:
            sheet.api.AutoFilterMode = False
        sheet.range((first - 1, 1), (last, width)).api.AutoFilter(Field=1)


def detail(book, spec, stream):
    sheet = book.sheets[spec["sheet"]]
    start, total, count = spec["start"], spec["total"], spec["count"]
    if total > start + 1:
        height = sheet.range((start, 1)).row_height
        inserted = sheet.range(f"{start + 1}:{total - 1}")
        inserted.insert("down", "format_from_left_or_above")
        inserted.row_height = height
    if not count:
        sheet.range((start, 1), (start, 6)).clear_contents()
    for row, values in batches(stream, start, count):
        end = row + len(values) - 1
        labels = sheet.range((row, 1), (end, 2))
        labels.number_format = "@"
        labels.api.Value2 = tuple(tuple(record[:2]) for record in values)
        sheet.range((row, 3), (end, 6)).formula = tuple(tuple(record[2:]) for record in values)
    sheet.range((total, 1)).api.Value2 = "合计"
    sheet.range((total, 3), (total, 6)).formula = (tuple(spec["formulas"]),)


def populate(book, payload):
    with payload.open(encoding="utf-8") as stream:
        manifest = json.loads(next(stream))
        name, address, value = manifest["scope"]
        scope = book.sheets[name].range(address)
        scope.number_format = "@"
        scope.api.Value2 = value
        for spec in manifest["tables"]:
            table(book, spec, stream)
        detail(book, manifest["detail"], stream)
        book.sheets[name].activate()


def open_book(app, path):
    return app.books.open(str(path), update_links=False, read_only=True, ignore_read_only_recommended=True,
                          notify=False, add_to_mru=False)


def attach_reports(app, book, payload, sheet_order):
    reports = open_book(app, payload)
    try:
        by_name = {sheet.name: sheet for sheet in reports.sheets}
        for index, name in enumerate(sheet_order):
            if name in by_name:
                if index < len(book.sheets):
                    by_name[name].copy(before=book.sheets[index])
                else:
                    by_name[name].copy(after=book.sheets[-1])
        book.sheets[0].activate()
    finally:
        reports.close()


def worker(operation, source, payload, destination, sheet_order):
    try:
        import xlwings as xw

        app = xw.App(visible=False, add_book=False)
    except Exception:
        return EXCEL_UNAVAILABLE
    with app:
        app.display_alerts = False
        app.enable_events = False
        app.screen_updating = False
        app.api.AutomationSecurity = 3
        book = open_book(app, source)
        try:
            app.calculation = "manual"
            if operation == "populate":
                populate(book, payload)
            elif operation == "attach_reports":
                attach_reports(app, book, payload, sheet_order)
            else:
                raise ValueError("Unknown Excel write operation.")
            app.calculation = "automatic"
            app.api.CalculateFullRebuild()
            book.save(str(destination))
        finally:
            book.close()
    return 0


if __name__ == "__main__" and len(sys.argv) >= 5:
    operation, source, payload, destination, *sheet_order = sys.argv[1:]
    raise SystemExit(worker(operation, Path(source), Path(payload), Path(destination), sheet_order))
