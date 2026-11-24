"""Populate source tabs and expand the AR detail without editing summary formulas."""

from collections import defaultdict
from copy import copy
from dataclasses import astuple, dataclass
from zipfile import ZipFile

from defusedxml.ElementTree import fromstring
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter, quote_sheetname
from openpyxl.workbook.properties import CalcProperties

from core.schema import AuditError, GL_FIELDS, GL_SHEET, TB_FIELDS, TB_SHEET, TB_AMOUNT_FIELDS

DETAIL = "应收账款明细表"
SUMMARY = "应收账款审定表"
AJE = "AJE"
DETAIL_FIRST = "AR_DETAIL_FIRST"
DETAIL_TOTAL = "AR_DETAIL_TOTAL"
SUMMARY_FIRST = "AR_SUMMARY_FIRST"
DETAIL_HEADERS = ("科目编码", "客商 / 辅助核算", "期初净额", "本期借方", "本期贷方", "期末余额")
SUMMARY_LABELS = ("期初净额", "本期借方", "本期贷方", "未审数", "审计调整", "审定数")
AMOUNT_FORMAT = '#,##0.00;(#,##0.00);"-"'
RAW_TB_FIELDS = (*TB_FIELDS, "account_key")
EXCEL_ROWS = 1_048_576


@dataclass(frozen=True)
class Coverage:
    root_rows: tuple[int, ...]
    detail_rows: tuple[int, ...]
    uncovered_codes: tuple[str, ...]
    tb_codes: frozenset[str]


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


def leaf_rows(rows):
    """Use terminal accounts and prefer auxiliary detail over its same-code subtotal."""
    by_code = defaultdict(list)
    for index, record in rows:
        by_code[record.account_code].append((index, record))
    levels = {code: min(record.level for _, record in group) for code, group in by_code.items()}
    rollups = {
        code[:end] for code, level in levels.items()
        for end in range(1, len(code)) if levels.get(code[:end], level) < level
    }
    for code in sorted(by_code.keys() - rollups):
        group = by_code[code]
        yield from ([item for item in group if item[1].aux_name is not None] or group)


def select_coverage(data, target):
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
    return Coverage(tuple(root_rows), tuple(index for index, _ in leaf_rows(selected)), tuple(sorted(uncovered)),
                    frozenset(row.account_code for row in data.tb))


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


def anchor_row(book, anchor, sheet, column):
    try:
        [(name, address)] = book.defined_names[anchor].destinations
        cell = book[name][address]
        if name == sheet and cell.column == column:
            return cell.row
    except (AttributeError, KeyError, TypeError, ValueError):
        pass
    raise AuditError(f"模板锚点 {anchor} 无效。")


def write_data(sheet, names, rows):
    if len(rows) + 1 > EXCEL_ROWS:
        raise AuditError("标准数据超出 Excel 行数上限。")
    if tuple(cell.value for cell in sheet[1]) != names or sheet.max_row != 1:
        raise AuditError("底层模板表须只有标准字段表头。")
    amounts = {*TB_AMOUNT_FIELDS, "debit", "credit"}
    for index, row in enumerate(rows, start=2):
        values = astuple(row)
        if sheet.title == TB_SHEET:
            values += (f"acct:{row.account_code}",)
        for column, (field, value) in enumerate(zip(names, values, strict=True), start=1):
            cell = sheet.cell(index, column)
            literal(cell, value)
            if field in amounts:
                cell.number_format = AMOUNT_FORMAT
            elif isinstance(value, str):
                cell.number_format = "@"
    sheet.auto_filter.ref = f"A1:{get_column_letter(len(names))}{len(rows) + 1}"


def build(template, data, coverage, ar_account):
    check_template(template)
    book = load_workbook(template, keep_links=False)
    write_data(book[TB_SHEET], RAW_TB_FIELDS, data.tb)
    write_data(book[GL_SHEET], GL_FIELDS, data.gl)
    start = anchor_row(book, DETAIL_FIRST, DETAIL, 1)
    total = anchor_row(book, DETAIL_TOTAL, DETAIL, 1)
    summary_start = anchor_row(book, SUMMARY_FIRST, SUMMARY, 2)
    detail = book[DETAIL]
    if total != start + 1 or detail[f"A{start}"].value != "{{DATA_START}}" or detail[f"A{total}"].value != "{{TOTAL}}":
        raise AuditError("模板须保留明细起始和合计锚点。")
    if len(coverage.detail_rows) + start > EXCEL_ROWS:
        raise AuditError("应收明细超出 Excel 行数上限。")
    scope_cell = book[SUMMARY]["B3"]
    if scope_cell.data_type == "f":
        raise AuditError("审定表 B3 是一级科目输入格，须保留为文本。")
    literal(scope_cell, ar_account)
    added = max(0, len(coverage.detail_rows) - 1)
    if added:
        total_height = detail.row_dimensions[total].height
        detail.insert_rows(total, added)
        total += added
        detail.row_dimensions[total].height = total_height
    book.defined_names[DETAIL_TOTAL].attr_text = f"{quote_sheetname(DETAIL)}!$A${total}"
    model_height = detail.row_dimensions[start].height
    for index, source in enumerate(coverage.detail_rows, start=start):
        record = data.tb[source - 2]
        detail.row_dimensions[index].height = model_height
        for column in range(1, len(DETAIL_HEADERS) + 1):
            detail.cell(index, column)._style = copy(detail.cell(start, column)._style)
        literal(detail.cell(index, 1), record.account_code)
        literal(detail.cell(index, 2), record.aux_name or record.account_name)
        detail[f"C{index}"] = f"=ROUND(Raw_TB!E{source}-Raw_TB!F{source},2)"
        detail[f"D{index}"] = f"=Raw_TB!G{source}"
        detail[f"E{index}"] = f"=Raw_TB!H{source}"
        detail[f"F{index}"] = f"=ROUND(C{index}+D{index}-E{index},2)"
    if not coverage.detail_rows:
        detail[f"A{start}"] = None
        for letter in "CDEF":
            detail[f"{letter}{start}"] = None
    detail[f"A{total}"] = "合计"
    for letter in "CDEF":
        detail[f"{letter}{total}"] = f"=ROUND(SUM({letter}{start}:{letter}{total - 1}),2)"
    book.calculation = CalcProperties(calcMode="auto", fullCalcOnLoad=True, forceFullCalc=True)
    return book, Layout(start, total, summary_start, len(coverage.detail_rows))
