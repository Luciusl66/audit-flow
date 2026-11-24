"""Build the AR template and synthetic TB and GL inputs."""

import csv
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter, quote_sheetname
from openpyxl.workbook.defined_name import DefinedName

from core.fill_workpaper import (
    ACCOUNT_KEY_PREFIX, AJE, AMOUNT_FORMAT, DETAIL, DETAIL_FIRST, DETAIL_HEADERS, DETAIL_TOTAL, RAW_TB_FIELDS,
    SUMMARY, SUMMARY_FIRST, SUMMARY_LABELS,
)
from core.schema import GL_FIELDS, GL_SHEET, TB_SHEET

ROOT = Path(__file__).resolve().parents[1]
NAVY = "263F56"
INK = "24384A"
LINE = Side(style="thin", color="CFD9E3")
TB_HEADERS = ("科目编码", "科目名称", "科目级次", "余额方向", "期初借方", "期初贷方",
              "本期借方", "本期贷方", "期末借方", "期末贷方", "往来单位")
GL_HEADERS = ("记账日期", "凭证号", "摘要", "科目编码", "借方金额", "贷方金额", "往来单位")


def font(color=INK, *, size=10, bold=False):
    return Font(name="Microsoft YaHei", size=size, color=color, bold=bold)


def columns(sheet, widths):
    sheet.sheet_view.showGridLines = False
    for index, width in enumerate(widths, 1):
        sheet.column_dimensions[get_column_letter(index)].width = width


def title(sheet, value):
    sheet["A2"] = value
    sheet["A2"].font = font(size=15, bold=True)
    sheet.row_dimensions[2].height = 25
    sheet.sheet_properties.tabColor = NAVY


def headers(sheet, row, values):
    for column, value in enumerate(values, 1):
        cell = sheet.cell(row, column, value)
        cell.font = font("FFFFFF", bold=True)
        cell.fill = PatternFill("solid", fgColor=NAVY)
        cell.alignment = Alignment(horizontal="center", vertical="center")
    sheet.row_dimensions[row].height = 24


def anchor(book, name, sheet, cell):
    book.defined_names.add(DefinedName(name, attr_text=f"{quote_sheetname(sheet.title)}!{cell}"))


def raw_sum(column):
    # Match the prefixed account_key in hidden column L, level 1 in C and an empty
    # aux_name in K. ACCOUNT_CODE excludes *, ? and ~ from the input in B3.
    return (
        f'SUMIFS(Raw_TB!${column}:${column},Raw_TB!$L:$L,"{ACCOUNT_KEY_PREFIX}"&$B$3,'
        'Raw_TB!$C:$C,1,Raw_TB!$K:$K,"")'
    )


def template(path):
    book = Workbook()
    summary = book.active
    summary.title = SUMMARY
    detail = book.create_sheet(DETAIL)
    adjustments = book.create_sheet(AJE)
    columns(summary, (25, 28))
    title(summary, SUMMARY)
    summary["A3"], summary["B3"] = "应收一级科目", "1122"
    summary["B3"].number_format = "@"
    headers(summary, 5, ("项目", "金额"))
    preset = (
        f"=ROUND({raw_sum('E')}-{raw_sum('F')},2)", f"=ROUND({raw_sum('G')},2)", f"=ROUND({raw_sum('H')},2)",
        f"=ROUND({raw_sum('I')}-{raw_sum('J')},2)",
        '=ROUND(SUMIFS(AJE!$B:$B,AJE!$A:$A,"应收账款"),2)', "=ROUND(B9+B10,2)",
    )
    for index, (name, formula) in enumerate(zip(SUMMARY_LABELS, preset, strict=True), 6):
        summary.cell(index, 1, name).font = font(bold=index >= 9)
        cell = summary.cell(index, 2, formula)
        cell.font = font(bold=index >= 9)
        cell.number_format = AMOUNT_FORMAT
        if index == 11:
            for total_cell in summary[index]:
                total_cell.fill = PatternFill("solid", fgColor="E8EFF6")
                total_cell.border = Border(top=LINE)
        summary.row_dimensions[index].height = 21
    columns(detail, (18, 29, 19, 19, 19, 19))
    title(detail, DETAIL)
    detail["A3"] = "期末余额按期初净额加借方、减贷方计算。"
    detail["A3"].font = font("566879")
    headers(detail, 5, DETAIL_HEADERS)
    detail.freeze_panes = "C6"
    detail.row_dimensions[6].height = 21
    detail.row_dimensions[7].height = 23
    for index in range(1, 7):
        cell = detail.cell(6, index)
        cell.font = font("216148" if 3 <= index <= 5 else INK)
        cell.alignment = Alignment(vertical="center", horizontal="left" if index <= 2 else "right")
        cell.border = Border(bottom=LINE)
        cell.number_format = "@" if index <= 2 else AMOUNT_FORMAT
        total = detail.cell(7, index)
        total.font = font(bold=True)
        total.border = Border(top=LINE)
        total.fill = PatternFill("solid", fgColor="E8EFF6")
        if index >= 3:
            col = get_column_letter(index)
            total.value = f"=ROUND(SUM({col}6:{col}6),2)"
            total.number_format = AMOUNT_FORMAT
    detail["A6"], detail["A7"] = "{{DATA_START}}", "{{TOTAL}}"
    detail["C6"], detail["D6"], detail["E6"], detail["F6"] = (
        "=ROUND(Raw_TB!E2-Raw_TB!F2,2)", "=Raw_TB!G2", "=Raw_TB!H2", "=ROUND(C6+D6-E6,2)",
    )
    title(adjustments, "审计调整")
    columns(adjustments, (18, 20, 36))
    adjustments["A3"] = "填写应收账款调整。正数增加审定数，负数减少审定数。"
    adjustments["A3"].font = font("566879")
    headers(adjustments, 5, ("调整项目", "调整金额", "说明"))
    for row in range(6, 11):
        adjustments.cell(row, 1, "应收账款")
        cell = adjustments.cell(row, 2)
        cell.number_format = AMOUNT_FORMAT
        cell.fill = PatternFill("solid", fgColor="FFF1CD")
        cell.font = font("16478D")
    for name, fields, widths in (
        (TB_SHEET, RAW_TB_FIELDS, (18, 23, 11, 11, 17, 17, 17, 17, 17, 17, 23, 21)),
        (GL_SHEET, GL_FIELDS, (16, 15, 28, 18, 17, 17, 25)),
    ):
        sheet = book.create_sheet(name)
        columns(sheet, widths)
        headers(sheet, 1, fields)
        sheet.freeze_panes = "B2"
        sheet.sheet_properties.tabColor = "CFAE7F"
        if name == TB_SHEET:
            sheet.column_dimensions["L"].hidden = True
    anchor(book, DETAIL_FIRST, detail, "$A$6")
    anchor(book, DETAIL_TOTAL, detail, "$A$7")
    anchor(book, SUMMARY_FIRST, summary, "$B$6")
    book.save(path)
    book.close()


def mock(folder):
    book = Workbook()
    sheet = book.active
    sheet.title = "模拟科目余额表"
    headers(sheet, 1, TB_HEADERS)
    columns(sheet, (15, 24, 11, 11, 15, 15, 15, 15, 15, 15, 22))
    for record in (
        ("1002", "银行存款", 1, "借", 2500, 0, 150, 0, 2650, 0, None),
        ("1122", "应收账款", 1, "借", "1,000.00", 50, "1,150.00", 150, "1,950.00", 0, None),
        ("1122.01", "应收账款 / 销售", 2, "借", 1000, 0, 800, 150, 1650, 0, None),
        ("1122.01", "应收账款 / 销售", 2, "借", 600, 0, 500, 125, 975, 0, "模拟甲公司"),
        ("1122.01", "应收账款 / 销售", 2, "借", 400, 0, 300, 25, 675, 0, "模拟乙公司"),
        ("1122.02", "应收账款 / 服务", 2, "贷", 0, 50, 350, 0, 300, 0, "模拟丙公司"),
        ("4001", "实收资本", 1, "贷", 0, 3450, 0, 0, 0, 3450, None),
        ("6001", "主营业务收入", 1, "贷", 0, 0, 0, 1150, 0, 1150, None),
    ):
        sheet.append(record)
    book.save(folder / "demo_tb.xlsx")
    book.close()
    with (folder / "demo_gl.csv").open("w", encoding="utf-8-sig", newline="") as output:
        csv.writer(output).writerows((
            GL_HEADERS,
            ("2025-01-02", "V-0001", "模拟暂挂", "1122.01", "1,200.50", "", "模拟甲公司"),
            ("2025-01-02", "V-0001", "模拟暂挂", "6001", "", "1,200.50", ""),
            ("2025-01-03", "V-0002", "模拟红冲", "1122.01", "(1,200.50)", "", "模拟甲公司"),
            ("2025-01-03", "V-0002", "模拟红冲", "6001", "", "(1,200.50)", ""),
            ("2025-01-05", "V-0003", "模拟赊销", "1122.01", "500.00", "", "模拟甲公司"),
            ("2025-01-05", "V-0003", "模拟赊销", "6001", "", "500.00", ""),
            ("2025-01-12", "V-0004", "模拟赊销", "1122.01", 300, "", "模拟乙公司"),
            ("2025-01-12", "V-0004", "模拟赊销", "6001", "", 300, ""),
            ("2025-01-18", "V-0005", "模拟赊销", "1122.02", 350, "", "模拟丙公司"),
            ("2025-01-18", "V-0005", "模拟赊销", "6001", "", 350, ""),
            ("2025-01-31", "V-0006", "模拟回款", "1002", 150, "", ""),
            ("2025-01-31", "V-0006", "模拟回款", "1122.01", "", 125, "模拟甲公司"),
            ("2025-01-31", "V-0006", "模拟回款", "1122.01", "", 25, "模拟乙公司"),
        ))


def main():
    (ROOT / "templates").mkdir(exist_ok=True)
    folder = ROOT / "data" / "mock"
    folder.mkdir(parents=True, exist_ok=True)
    template(ROOT / "templates" / "ar_template.xlsx")
    mock(folder)


if __name__ == "__main__":
    main()
