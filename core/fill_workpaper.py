"""Populate source tabs and expand the AR detail without editing summary formulas."""

from abc import ABC, abstractmethod
from contextlib import closing
from copy import copy
from dataclasses import astuple, dataclass
from datetime import date
from decimal import Decimal
import json
from shutil import copyfile
from openpyxl import Workbook, load_workbook
from openpyxl.utils import coordinate_to_tuple, get_column_letter, quote_sheetname
from openpyxl.utils.datetime import to_excel
from openpyxl.workbook.properties import CalcProperties
from openpyxl.worksheet.cell_range import CellRange

from core.accounts import Ledger, leaf_rows
from core.blueprint import formula_signature, load_blueprint
from core.excel_writer import (
    check_template, excel_literal, insertion_requires_native, literal, reject_merged_writes, require_excel_writer,
    run_excel_writer,
)
from core.recalc import formulas, write_caches
from core.schema import AuditError, GL_FIELDS, GL_SHEET, PRIOR_TB_SHEET, TB_SHEET
from core.workpaper_data import ACCOUNT_KEY_PREFIX, AJE, AMOUNT_FIELDS, AMOUNT_FORMAT, RAW_PRIOR_FIELDS, RAW_TB_FIELDS, TABLE_FORMATS, prior_records

DETAIL = "应收账款明细表"
SUMMARY = "应收账款审定表"
DETAIL_FIRST = "AR_DETAIL_FIRST"
DETAIL_TOTAL = "AR_DETAIL_TOTAL"
SUMMARY_FIRST = "AR_SUMMARY_FIRST"
DETAIL_HEADERS = ("科目编码", "客商 / 辅助核算", "期初净额", "本期借方", "本期贷方", "期末余额")
SUMMARY_LABELS = ("期初净额", "本期借方", "本期贷方", "未审数", "审计调整", "审定数")


@dataclass(frozen=True)
class Coverage:
    root_rows: tuple[int, ...]
    detail_rows: tuple[int, ...]
    uncovered_codes: tuple[str, ...]
    tb_codes: frozenset[str]
    grain_issues: tuple = ()


@dataclass(frozen=True)
class Layout:
    detail_start: int
    detail_total: int
    summary_start: int
    details: int

    def required_formulas(self):
        return {
            *((SUMMARY, f"B{row}") for row in range(self.summary_start, self.summary_start + len(SUMMARY_LABELS))),
            *((DETAIL, f"{column}{row}") for row in (*range(self.detail_start, self.detail_start + self.details),
                                                        self.detail_total) for column in "CDEF"),
        }


def select_coverage(data, target, *, ledger=None):
    roots = {row.account_code for row in data.tb if row.level == 1 and row.aux_name is None}
    selected = []
    uncovered = set()
    root_rows = []
    for index, row in enumerate(data.tb, 2):
        code = row.account_code
        if row.level == 1:
            owner = code
        else:
            owner = next((code[:size] for size in range(len(code) - 1, 0, -1) if code[:size] in roots), None)
        if owner != target or target not in roots:
            uncovered.add(code)
            continue
        selected.append((index, row))
        if row.level == 1 and row.aux_name is None:
            root_rows.append(index)
    if ledger is None:
        ledger = Ledger(data)
    return Coverage(tuple(root_rows), tuple(index for index, _ in leaf_rows(selected)), tuple(sorted(uncovered)),
                    frozenset(row.account_code for row in data.tb), tuple(ledger.grain_issues({target})))


def tables(blueprint, data):
    result = [(blueprint.sheet(TB_SHEET), RAW_TB_FIELDS, data.tb),
              (blueprint.sheet(GL_SHEET), GL_FIELDS, data.gl)]
    if data.prior_tb is not None or any(spec.name == PRIOR_TB_SHEET for spec in blueprint.sheets):
        result.append((blueprint.sheet(PRIOR_TB_SHEET), RAW_PRIOR_FIELDS, data.prior_tb or ()))
    return result


def table_values(name, records, data):
    if name == PRIOR_TB_SHEET:
        if records:
            yield from prior_records(data, Ledger(data))
        return
    for record in records:
        values = astuple(record)
        if name == TB_SHEET:
            # The hidden text key keeps codes such as 001 and 1 distinct in SUMIFS.
            values += (f"{ACCOUNT_KEY_PREFIX}{record.account_code}",)
        yield values


def prepare_layout(book, blueprint, data, coverage):
    blueprint.validate(book, source="template")
    summary, detail = blueprint.sheet(SUMMARY), blueprint.sheet(DETAIL)
    try:
        start, detail_column = coordinate_to_tuple(detail.anchors[DETAIL_FIRST])
        initial_total, total_column = coordinate_to_tuple(detail.anchors[DETAIL_TOTAL])
        summary_start, summary_column = coordinate_to_tuple(summary.anchors[SUMMARY_FIRST])
    except KeyError:
        raise AuditError("底稿蓝图缺少应收审定表或明细表锚点。") from None
    if (detail_column != 1 or total_column != 1 or initial_total != start + 1 or summary_column != 2
            or detail.data.min_row != start or detail.data.min_col != 1 or detail.data.max_col != len(DETAIL_HEADERS)
            or detail.detail_formulas != CellRange(f"C{start}:F{initial_total}") or summary.data.coord != "B3"):
        raise AuditError("底稿蓝图与应收工作表的写入布局不兼容。")
    if book[DETAIL][f"A{start}"].value != "{{DATA_START}}" or book[DETAIL][f"A{initial_total}"].value != "{{TOTAL}}":
        raise AuditError("模板须保留明细起始和合计锚点。")
    if book[SUMMARY][summary.data.coord].data_type == "f":
        raise AuditError("审定表一级科目输入格须保留为文本。")
    reject_merged_writes(book[SUMMARY], summary.data)
    for address in (f"A{start}:F{start}", f"A{initial_total}", f"C{initial_total}:F{initial_total}"):
        reject_merged_writes(book[DETAIL], CellRange(address))
    detail_total = start + max(1, len(coverage.detail_rows))
    if detail_total > detail.data.max_row:
        raise AuditError("应收明细超出底稿蓝图行数上限。")
    for spec, names, records in tables(blueprint, data):
        if (spec.data.min_row != 2 or spec.data.min_col != 1 or spec.data.max_col != len(names)
                or spec.headers != ((CellRange(f"A1:{get_column_letter(len(names))}1"), names),)):
            raise AuditError("底层模板表的标准字段位置与底稿蓝图不兼容。")
        sheet = book[spec.name]
        if any(cell.value is not None for row in sheet.iter_rows(min_row=2) for cell in row):
            raise AuditError("底层模板表须只有标准字段表头和空白数据区。")
        if len(records) + spec.data.min_row - 1 > spec.data.max_row:
            raise AuditError("标准数据超出底稿蓝图行数上限。")
        reject_merged_writes(sheet, spec.data)
    frozen = blueprint.frozen_formulas(book)
    if not {(SUMMARY, f"B{row}") for row in range(summary_start, summary_start + len(SUMMARY_LABELS))} <= frozen.keys():
        raise AuditError("底稿蓝图须保护全部审定表公式。")
    return Layout(start, detail_total, summary_start, len(coverage.detail_rows)), frozen


def detail_values(data, coverage, layout):
    for row, source in enumerate(coverage.detail_rows, layout.detail_start):
        record = data.tb[source - 2]
        yield (record.account_code, record.aux_name or record.account_name,
               f"=ROUND(Raw_TB!E{source}-Raw_TB!F{source},2)", f"=Raw_TB!G{source}", f"=Raw_TB!H{source}",
               f"=ROUND(C{row}+D{row}-E{row},2)")


def total_formulas(layout):
    return tuple(f"=ROUND(SUM({col}{layout.detail_start}:{col}{layout.detail_total - 1}),2)" for col in "CDEF")


def openpyxl_table(sheet, spec, names, records, data):
    for index, values in enumerate(table_values(spec.name, records, data), spec.data.min_row):
        for column, (field, value) in enumerate(zip(names, values, strict=True), start=1):
            cell = sheet.cell(index, column)
            literal(cell, value)
            if field in AMOUNT_FIELDS:
                cell.number_format = AMOUNT_FORMAT
            elif isinstance(value, str):
                cell.number_format = "@"
    sheet.auto_filter.ref = f"A1:{get_column_letter(spec.data.max_col)}{len(records) + spec.data.min_row - 1}"


def build(template, data, coverage, ar_account, *, blueprint=None):
    """Build a workpaper with openpyxl, leaving the report sheets to the caller."""
    if blueprint is None:
        blueprint = load_blueprint()
    check_template(template)
    book = load_workbook(template, keep_links=False)
    try:
        layout, frozen = prepare_layout(book, blueprint, data, coverage)
        initial_total = layout.detail_start + 1
        if layout.detail_total > initial_total and insertion_requires_native(
                book[DETAIL], layout.detail_start, initial_total, blueprint.sheet(DETAIL).detail_formulas,
                formulas(book), (DETAIL_TOTAL,)):
            raise AuditError("此模板的明细扩行涉及合并、条件格式或已有引用。请使用 --engine xlwings。")
        for spec, names, records in tables(blueprint, data):
            openpyxl_table(book[spec.name], spec, names, records, data)
        fill_openpyxl_detail(book, blueprint, data, coverage, ar_account, layout)
        validate_written_book(book, blueprint, layout, frozen, source="template")
    except BaseException:
        book.close()
        raise
    return book, layout


def fill_openpyxl_detail(book, blueprint, data, coverage, ar_account, layout):
    start, total = layout.detail_start, layout.detail_total
    detail = book[DETAIL]
    literal(book[SUMMARY][blueprint.sheet(SUMMARY).data.coord], ar_account)
    added = total - start - 1
    if added:
        initial_total = start + 1
        total_height = detail.row_dimensions[initial_total].height
        detail.insert_rows(initial_total, added)
        detail.row_dimensions[total].height = total_height
    book.defined_names[DETAIL_TOTAL].attr_text = f"{quote_sheetname(detail.title)}!$A${total}"
    model_height = detail.row_dimensions[start].height
    for index, values in enumerate(detail_values(data, coverage, layout), start=start):
        detail.row_dimensions[index].height = model_height
        for column, value in enumerate(values, 1):
            cell = detail.cell(index, column)
            cell._style = copy(detail.cell(start, column)._style)
            if column <= 2:
                literal(cell, value)
            else:
                cell.value = value
    if not layout.details:
        detail[f"A{start}"] = None
        for letter in "CDEF":
            detail[f"{letter}{start}"] = None
    detail[f"A{total}"] = "合计"
    for letter, formula in zip("CDEF", total_formulas(layout), strict=True):
        detail[f"{letter}{total}"] = formula
    book.calculation = CalcProperties(calcMode="auto", fullCalcOnLoad=True, forceFullCalc=True)


def validate_written_book(book, blueprint, layout, frozen, *, source=None):
    blueprint.validate(book, source=source, moved_anchors={DETAIL_TOTAL: f"A{layout.detail_total}"})
    if blueprint.frozen_formulas(book) != frozen:
        raise AuditError("写入改变了底稿蓝图保护的公式。")


def native_payload(path, blueprint, data, coverage, ar_account, layout, epoch):
    def number(value):
        if isinstance(value, Decimal):
            return float(value)
        if isinstance(value, date):
            return to_excel(value, epoch)
        raise TypeError

    table_specs = tables(blueprint, data)
    manifest = {
        "scope": (SUMMARY, blueprint.sheet(SUMMARY).data.coord, excel_literal(ar_account)),
        "tables": [
            {"sheet": spec.name, "first": spec.data.min_row, "count": len(records),
             "formats": [TABLE_FORMATS.get(name, "@") for name in names]}
            for spec, names, records in table_specs
        ],
        "detail": {"sheet": DETAIL, "start": layout.detail_start, "total": layout.detail_total,
                   "count": layout.details, "formulas": total_formulas(layout)},
    }
    with path.open("w", encoding="utf-8") as stream:
        stream.write(json.dumps(manifest, ensure_ascii=False) + "\n")
        for spec, _, records in table_specs:
            for row in table_values(spec.name, records, data):
                values = tuple(excel_literal(value) for value in row)
                stream.write(json.dumps(values, ensure_ascii=False, default=number) + "\n")
        for row in detail_values(data, coverage, layout):
            values = (excel_literal(row[0]), excel_literal(row[1]), *row[2:])
            stream.write(json.dumps(values, ensure_ascii=False) + "\n")


class WorkbookWriter(ABC):
    """Populate the staged workbook and expose a book for the existing report renderers."""

    def __init__(self, blueprint, scratch, timeout):
        self.blueprint = blueprint
        self.scratch = scratch
        self.timeout = timeout
        self.stage = scratch / "workpaper.xlsx"
        self.book = None

    @abstractmethod
    def populate(self, data, coverage, ar_account):
        """Write the staged template and return its layout and formulas."""

    @abstractmethod
    def finish(self, calculation):
        """Finish the report sheets and return the validated output path."""

    def close(self):
        if self.book is not None:
            self.book.close()


class OpenpyxlWriter(WorkbookWriter):
    def populate(self, data, coverage, ar_account):
        self.book, self.layout = build(self.blueprint.template, data, coverage, ar_account, blueprint=self.blueprint)
        self.frozen = self.blueprint.frozen_formulas(self.book)
        actual = formulas(self.book)
        self.book.save(self.stage)
        return self.layout, actual

    def finish(self, calculation):
        validate_written_book(self.book, self.blueprint, self.layout, self.frozen)
        self.book.save(self.stage)
        if not calculation.caches:
            return self.stage
        final = self.scratch / "cached.xlsx"
        write_caches(self.stage, final, calculation.caches)
        return final


class XlwingsWriter(WorkbookWriter):
    def inspect(self, path, *, source=None):
        check_template(path)
        with closing(load_workbook(path, read_only=True, keep_links=False)) as book:
            validate_written_book(book, self.blueprint, self.layout, self.frozen, source=source)
            return formulas(book)

    def populate(self, data, coverage, ar_account):
        require_excel_writer()
        template = self.scratch / "template.xlsx"
        copyfile(self.blueprint.template, template)
        check_template(template)
        with closing(load_workbook(template, keep_links=False)) as original:
            self.layout, self.frozen = prepare_layout(original, self.blueprint, data, coverage)
            epoch = original.epoch
        payload = self.scratch / "write-data.jsonl"
        native_payload(payload, self.blueprint, data, coverage, ar_account, self.layout, epoch)
        try:
            run_excel_writer("populate", template, payload, self.stage, self.timeout)
        finally:
            payload.unlink()
        self.actual = self.inspect(self.stage, source="template")
        rows = (values[2:] for values in detail_values(data, coverage, self.layout))
        for index, values in enumerate(rows, self.layout.detail_start):
            self.check_detail_formulas(index, values)
        self.check_detail_formulas(self.layout.detail_total, total_formulas(self.layout))
        self.book = Workbook()
        self.book.remove(self.book.active)
        return self.layout, self.actual

    def check_detail_formulas(self, row, values):
        if any(self.actual.get((DETAIL, f"{column}{row}")) != value for column, value in zip("CDEF", values, strict=True)):
            raise AuditError("Excel 原生写入后的明细公式与底稿蓝图不一致。")

    def finish(self, calculation):
        self.blueprint.validate(self.book, source="report")
        reports = self.scratch / "reports.xlsx"
        self.book.save(reports)
        final = self.scratch / "excel.xlsx"
        sheet_order = tuple(spec.name for spec in self.blueprint.sheets)
        run_excel_writer("attach_reports", self.stage, reports, final, self.timeout, sheet_order)
        written = self.inspect(final)
        if written.keys() != self.actual.keys() or any(
            formula_signature(value) != formula_signature(self.actual[key]) for key, value in written.items()
        ):
            raise AuditError("整理报告工作表后，Excel 原生底稿的公式位置或内容发生变化。")
        return final


def workpaper_writer(engine, blueprint, scratch, timeout):
    writer_type = {"openpyxl": OpenpyxlWriter, "xlwings": XlwingsWriter}[engine]
    return closing(writer_type(blueprint, scratch, timeout))
