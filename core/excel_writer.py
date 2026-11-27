"""Apply the same checked template writes with openpyxl or native Excel."""

from copy import copy
from dataclasses import dataclass
from datetime import date
from decimal import Decimal
from importlib.util import find_spec
from itertools import groupby, islice
import json
import os
from pathlib import Path
from shutil import copyfile
import subprocess
import sys
from zipfile import ZipFile

from defusedxml.ElementTree import fromstring
from openpyxl import Workbook, load_workbook
from openpyxl.formula import Tokenizer
from openpyxl.formula.tokenizer import TokenizerError
from openpyxl.utils import coordinate_to_tuple, get_column_letter, quote_sheetname
from openpyxl.utils.datetime import to_excel
from openpyxl.workbook.properties import CalcProperties
from openpyxl.worksheet.cell_range import CellRange

from core.blueprint import SheetBlueprint, formula_signature
from core.recalc import formulas as workbook_formulas, write_caches
from core.schema import AuditError

ROOT = Path(__file__).resolve().parents[1]
EXCEL_UNAVAILABLE = 3
BATCH_ROWS = 2_000


def require_excel_writer():
    if os.name != "nt" or find_spec("xlwings") is None:
        raise AuditError("xlwings 写入需要 Windows 本机 Microsoft Excel 和 xlwings。请使用 --engine openpyxl。")


def run_excel_writer(operation, source, payload, destination, timeout, sheet_order=()):
    stage = {"populate": "填充模板", "apply_plan": "填充模板", "attach_reports": "插入报告"}[operation]
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


def literal(cell, value):
    cell.value = value
    if isinstance(value, str):
        cell.data_type = "s"


def check_template(path):
    blocked = ("externallinks/", "connections.xml", "querytables/", "vbaproject", "activex/", "embeddings/")
    with ZipFile(path) as archive:
        for name in archive.namelist():
            if any(part in name.lower() for part in blocked):
                raise AuditError("模板包含外部数据连接或活动内容。")
            if name.endswith(".rels"):
                relationships = fromstring(archive.read(name))
                if any(rel.get("TargetMode", "").lower() == "external" for rel in relationships):
                    raise AuditError("模板包含外部关系。")


def reject_merged_writes(sheet, area):
    for merged in sheet.merged_cells.ranges:
        if area.isdisjoint(merged):
            continue
        if area.min_row == area.max_row == merged.min_row and area.min_col == area.max_col == merged.min_col:
            continue
        raise AuditError(f"工作表 {sheet.title} 的蓝图写入区域包含合并单元格的非锚点。")


@dataclass(frozen=True)
class Section:
    spec: SheetBlueprint
    first: int
    total: int
    column: int
    width: int

    def cell(self, offset, row, *, fixed=False):
        letter = get_column_letter(self.column + offset)
        return f"${letter}${row}" if fixed else f"{letter}{row}"

    def position(self, offset, row):
        return self.spec.name, self.cell(offset, row)


class ExcelWriter:
    """Plan writes against reviewed anchors before touching either workbook engine."""

    def __init__(self, engine, blueprint, scratch, timeout):
        self.engine = engine
        self.blueprint = blueprint
        self.scratch = scratch
        self.timeout = timeout
        self.stage = scratch / "workpaper.xlsx"
        self.source = scratch / "template.xlsx"
        self.operations = []
        self.moved = {}
        self.shifts = {}
        self.generated = {}
        self.book = None

    def __enter__(self):
        if self.engine == "xlwings":
            require_excel_writer()
        copyfile(self.blueprint.template, self.source)
        check_template(self.source)
        self.template = load_workbook(self.source, keep_links=False)
        try:
            self.blueprint.validate(self.template, source="template")
            self.frozen = self.blueprint.frozen_formulas(self.template)
            self.existing = workbook_formulas(self.template)
        except BaseException:
            self.template.close()
            raise
        return self

    def __exit__(self, *_):
        self.template.close()
        if self.book is not None and self.book is not self.template:
            self.book.close()

    def reserve(self, role, count):
        spec = self.blueprint.role(role)
        stem = f"WP_{role.upper()}"
        try:
            first, column = coordinate_to_tuple(spec.anchors[f"{stem}_FIRST"])
            initial_total, total_column = coordinate_to_tuple(spec.anchors[f"{stem}_TOTAL"])
        except KeyError:
            raise AuditError(f"底稿蓝图缺少 {role} 的明细起始或合计锚点。") from None
        total = first + max(count, 1)
        if (spec.data.min_row != first or spec.data.min_col != column or column != total_column
                or initial_total != first + 1 or total > spec.data.max_row):
            raise AuditError(f"工作表 {spec.name} 的明细锚点或行数超出蓝图布局。")
        if total > initial_total:
            sheet = self.template[spec.name]
            if self.engine == "openpyxl" and self.requires_native(sheet, first, initial_total):
                raise AuditError("模板扩行涉及合并、条件格式或已有引用。请使用 --engine xlwings。")
            added = total - initial_total
            self.shifts[spec.name] = (initial_total, added)
            for name, address in spec.anchors.items():
                row, col = coordinate_to_tuple(address)
                if row >= initial_total:
                    self.moved[name] = f"{get_column_letter(col)}{row + added}"
            self.operations.append(({"type": "insert", "sheet": spec.name, "row": initial_total, "added": added,
                                     "first_col": column, "last_col": spec.data.max_col}, ()))
        return Section(spec, first, total, column, spec.data.max_col - column + 1)

    def requires_native(self, sheet, first, initial_total):
        if (any(area.max_row >= initial_total for area in sheet.merged_cells.ranges)
                or any(area.max_row >= first for rule in sheet.conditional_formatting for area in rule.sqref.ranges)
                or any(area.max_row >= first for rule in sheet.data_validations.dataValidation
                       for area in rule.sqref.ranges)
                or sheet.tables):
            return True
        spec = self.blueprint.sheet(sheet.title)
        for_sheet = {sheet.title, quote_sheetname(sheet.title)}
        for (name, address), formula in self.existing.items():
            if name == sheet.title:
                if spec.detail_formulas is not None and address in spec.detail_formulas:
                    continue
                if coordinate_to_tuple(address)[0] >= initial_total or not isinstance(formula, str):
                    return True
                try:
                    references = (token.value for token in Tokenizer(formula).items if token.subtype == "RANGE")
                    if any("!" not in reference or reference.rsplit("!", 1)[0] in for_sheet
                           for reference in references):
                        return True
                except TokenizerError:
                    return True
            elif isinstance(formula, str) and any(f"{title}!" in formula for title in for_sheet):
                return True
        return False

    def original_row(self, sheet, row):
        at, added = self.shifts.get(sheet, (row + 1, 0))
        if row < at:
            return row
        return None if row < at + added else row - added

    def grid(self, spec, row, column, values, formula_offsets=()):
        values = tuple(values)
        if not values:
            return
        width = len(values[0])
        area = CellRange(min_row=row, max_row=row + len(values) - 1, min_col=column, max_col=column + width - 1)
        if not area.issubset(spec.data):
            raise AuditError(f"工作表 {spec.name} 的写入超出蓝图范围。")
        formula_offsets = frozenset(formula_offsets)
        for index, record in enumerate(values, row):
            if len(record) != width:
                raise AuditError("底稿写入数据列数与蓝图不一致。")
            original = self.original_row(spec.name, index)
            for offset, value in enumerate(record):
                col = column + offset
                address = f"{get_column_letter(col)}{index}"
                old_address = f"{get_column_letter(col)}{original}" if original is not None else None
                dynamic = spec.detail_formulas
                allowed = dynamic is not None and dynamic.min_col <= col <= dynamic.max_col and (
                    original is None or old_address in dynamic)
                if ((spec.name, old_address) in self.frozen
                        or (spec.name, old_address) in self.existing and not allowed):
                    raise AuditError(f"工作表 {spec.name} 的写入与模板保护公式重叠。")
                if original is not None:
                    for merged in self.template[spec.name].merged_cells.ranges:
                        if old_address in merged and (original, col) != (merged.min_row, merged.min_col):
                            raise AuditError(f"工作表 {spec.name} 的写入与合并单元格非锚点重叠。")
                if offset in formula_offsets and value is not None:
                    if not allowed or not isinstance(value, str) or not value.startswith("="):
                        raise AuditError(f"工作表 {spec.name} 的动态公式超出蓝图公式区或为静态值。")
                    self.generated[spec.name, address] = value
        self.operations.append(({"type": "grid", "sheet": spec.name, "first": row, "column": column,
                                 "count": len(values), "width": width, "formulas": sorted(formula_offsets)}, values))

    def table(self, name, fields, records, count, formats):
        spec = self.blueprint.sheet(name)
        end = get_column_letter(len(fields))
        if (spec.data.min_row != 2 or spec.data.min_col != 1 or spec.data.max_col != len(fields)
                or spec.headers != ((CellRange(f"A1:{end}1"), tuple(fields)),)):
            raise AuditError(f"工作表 {name} 的标准字段位置与底稿蓝图不兼容。")
        if count + 1 > spec.data.max_row:
            raise AuditError("标准数据超出底稿蓝图行数上限。")
        sheet = self.template[name]
        if any(cell.value is not None for row in sheet.iter_rows(min_row=2) for cell in row):
            raise AuditError("底层模板表须只有标准字段表头和空白数据区。")
        reject_merged_writes(sheet, spec.data)
        self.operations.append(({"type": "table", "sheet": name, "first": 2, "count": count,
                                 "formats": [formats.get(field, "@") for field in fields]}, records))

    def validate_written(self, book, *, source=None):
        self.blueprint.validate(book, source=source, moved_anchors=self.moved)
        for (name, original), expected in self.frozen.items():
            row, col = coordinate_to_tuple(original)
            at, added = self.shifts.get(name, (row + 1, 0))
            cell = book[name].cell(row + (added if row >= at else 0), col)
            if cell.data_type != "f" or formula_signature(cell.value) != expected:
                raise AuditError("写入改变了底稿蓝图保护的公式。")
        actual = workbook_formulas(book)
        if not self.generated.keys() <= actual.keys():
            raise AuditError("写入后的工作簿缺少预期动态公式。")
        return actual

    def apply(self):
        if self.engine == "openpyxl":
            self.book = self.template
            for spec, rows in self.operations:
                apply_openpyxl(self.book, spec, rows)
            for name, address in self.moved.items():
                sheet, _ = next(self.book.defined_names[name].destinations)
                row, col = coordinate_to_tuple(address)
                self.book.defined_names[name].attr_text = f"{quote_sheetname(sheet)}!${get_column_letter(col)}${row}"
            self.book.calculation = CalcProperties(calcMode="auto", fullCalcOnLoad=True, forceFullCalc=True)
            self.actual = self.validate_written(self.book, source="template")
            self.book.save(self.stage)
        else:
            payload = self.scratch / "write-data.jsonl"
            with payload.open("w", encoding="utf-8") as stream:
                for spec, rows in self.operations:
                    stream.write(json.dumps(spec, ensure_ascii=False) + "\n")
                    for row in rows:
                        native = native_values(row, spec.get("formulas", ()), self.template.epoch)
                        stream.write(json.dumps(native, ensure_ascii=False) + "\n")
            run_excel_writer("apply_plan", self.source, payload, self.stage, self.timeout)
            self.actual = self.inspect(self.stage, source="template")
            self.book = Workbook()
            self.book.remove(self.book.active)
        return self.actual

    def inspect(self, path, *, source=None):
        check_template(path)
        book = load_workbook(path, read_only=True, keep_links=False)
        try:
            return self.validate_written(book, source=source)
        finally:
            book.close()

    def finish(self, calculation):
        if self.engine == "openpyxl":
            if self.validate_written(self.book) != self.actual:
                raise AuditError("整理报告工作表后，底稿公式发生变化。")
            self.book.save(self.stage)
            if not calculation.caches:
                return self.stage
            final = self.scratch / "cached.xlsx"
            write_caches(self.stage, final, calculation.caches)
        else:
            self.blueprint.validate(self.book, source="report")
            reports = self.scratch / "reports.xlsx"
            self.book.save(reports)
            final = self.scratch / "excel.xlsx"
            run_excel_writer("attach_reports", self.stage, reports, final, self.timeout,
                             tuple(spec.name for spec in self.blueprint.sheets))
            if self.inspect(final) != self.actual:
                raise AuditError("整理报告工作表后，Excel 原生底稿的公式发生变化。")
        return final


def native_values(values, formula_offsets, epoch):
    return tuple(
        float(value) if isinstance(value, Decimal) else to_excel(value, epoch) if isinstance(value, date)
        else value if offset in formula_offsets else excel_literal(value)
        for offset, value in enumerate(values)
    )


def apply_openpyxl(book, spec, rows):
    sheet = book[spec["sheet"]]
    operation = spec["type"]
    if operation == "insert":
        first, count = spec["row"], spec["added"]
        for row in sorted((row for row in sheet.row_dimensions if row >= first), reverse=True):
            dimension = sheet.row_dimensions.pop(row)
            dimension.index = row + count
            sheet.row_dimensions[row + count] = dimension
        sheet.insert_rows(first, count)
        model = first - 1
        for row in range(first, first + count):
            sheet.row_dimensions[row].height = sheet.row_dimensions[model].height
            for col in range(spec["first_col"], spec["last_col"] + 1):
                sheet.cell(row, col)._style = copy(sheet.cell(model, col)._style)
        return
    formats = spec.get("formats")
    formula_columns = set(spec.get("formulas", ()))
    first_col = spec.get("column", 1)
    for row, record in enumerate(rows, spec["first"]):
        for offset, value in enumerate(record):
            cell = sheet.cell(row, first_col + offset)
            if offset in formula_columns:
                cell.value = value
            else:
                literal(cell, value)
            if formats:
                cell.number_format = formats[offset]
    if formats:
        sheet.auto_filter.ref = f"A1:{get_column_letter(len(formats))}{spec['count'] + 1}"


def apply_native_plan(book, payload):
    with payload.open(encoding="utf-8") as stream:
        while (line := stream.readline()):
            spec = json.loads(line)
            operation = spec["type"]
            if operation == "table":
                table(book, spec, stream)
                continue
            sheet = book.sheets[spec["sheet"]]
            if operation == "insert":
                first, last = spec["row"], spec["row"] + spec["added"] - 1
                height = sheet.range((first - 1, spec["first_col"])).row_height
                inserted = sheet.range(f"{first}:{last}")
                inserted.insert("down", "format_from_left_or_above")
                inserted.row_height = height
                continue
            if operation != "grid":
                raise ValueError("Unknown write operation.")
            column, width = spec["column"], spec["width"]
            formula_columns = set(spec["formulas"])
            for first, rows in batches(stream, spec["first"], spec["count"]):
                last = first + len(rows) - 1
                for is_formula, group in groupby(range(width), key=lambda col: col in formula_columns):
                    columns = tuple(group)
                    start, stop = columns[0], columns[-1] + 1
                    target = sheet.range((first, column + start), (last, column + stop - 1))
                    values = tuple(tuple(row[start:stop]) for row in rows)
                    if is_formula:
                        target.formula = values
                    else:
                        target.api.Value2 = values


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
            elif operation == "apply_plan":
                apply_native_plan(book, payload)
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
