"""Build the bundled templates and synthetic acceptance inputs."""

from collections import defaultdict
import csv
from decimal import Decimal
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter, quote_sheetname
from openpyxl.workbook.defined_name import DefinedName
from yaml import safe_load

from core.blueprint import cell_range
from core.fill_workpaper import (
    ACCOUNT_KEY_PREFIX, AJE, AMOUNT_FORMAT, DETAIL, DETAIL_FIRST, DETAIL_HEADERS, DETAIL_TOTAL, RAW_TB_FIELDS,
    SUMMARY, SUMMARY_FIRST, SUMMARY_LABELS,
)
from core.schema import GL_FIELDS, GL_SHEET, PRIOR_TB_SHEET, TB_SHEET, ZERO
from core.workpaper_data import PERCENT_FORMAT, RAW_PRIOR_FIELDS

ROOT = Path(__file__).resolve().parents[1]
NAVY = "263F56"
INK = "24384A"
LINE = Side(style="thin", color="CFD9E3")
TB_HEADERS = ("科目编码", "科目名称", "科目级次", "余额方向", "期初借方", "期初贷方",
              "本期借方", "本期贷方", "期末借方", "期末贷方", "往来单位")
GL_HEADERS = ("记账日期", "凭证号", "摘要", "科目编码", "借方金额", "贷方金额", "往来单位")
PRIOR_HEADERS = ("科目编码", "科目名称", "上年审定借方", "上年审定贷方", "往来单位")
DEMO_TEMPLATES = {
    "1002": "cash_template.xlsx", "1122": "counterparty_template.xlsx", "6001": "revenue_template.xlsx",
    "1601": "fixed_assets_template.xlsx", "2221": "tax_template.xlsx",
}


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
    for column in "DEF":
        summary.column_dimensions[column].width = 20
    title(summary, SUMMARY)
    summary["A3"], summary["B3"] = "应收一级科目", "1122"
    summary["B3"].number_format = "@"
    headers(summary, 5, ("项目", "金额"))
    for column, label in enumerate(("上年审定数", "变动额", "变动率"), 4):
        style = summary.cell(5, 2)._style
        summary.cell(5, column, label)._style = style
    key = f'"{ACCOUNT_KEY_PREFIX}"&$B$3'
    previous = f"SUMIFS(Raw_Prior_TB!$C:$C,Raw_Prior_TB!$G:$G,{key})-SUMIFS(Raw_Prior_TB!$D:$D,Raw_Prior_TB!$G:$G,{key})"
    summary["D6"] = f'=IF(COUNTIF(Raw_Prior_TB!$G:$G,{key})=0,"",ROUND({previous},2))'
    summary["E6"] = '=IF(D6="","",ROUND(B11-D6,2))'
    summary["F6"] = '=IF(OR(D6="",D6=0),"",E6/ABS(D6))'
    summary["D6"].number_format = summary["E6"].number_format = AMOUNT_FORMAT
    summary["F6"].number_format = PERCENT_FORMAT
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
        (PRIOR_TB_SHEET, RAW_PRIOR_FIELDS, (18, 23, 22, 22, 23, 21, 21)),
    ):
        sheet = book.create_sheet(name)
        columns(sheet, widths)
        headers(sheet, 1, fields)
        sheet.freeze_panes = "B2"
        sheet.sheet_properties.tabColor = "CFAE7F"
        if name == TB_SHEET:
            sheet.column_dimensions["L"].hidden = True
        elif name == PRIOR_TB_SHEET:
            sheet.column_dimensions["F"].hidden = sheet.column_dimensions["G"].hidden = True
            anchor(book, "WP_PRIOR_FIRST", sheet, "$A$2")
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


def batch_template(path, config):
    book = Workbook()
    book.remove(book.active)
    for spec in config["sheets"]:
        if spec["source"] != "template":
            continue
        name = spec["name"]
        sheet = book.create_sheet(name)
        (address, labels), = spec["headers"].items()
        header_row = cell_range(address).min_row
        widths = tuple(18 if col == 1 else 28 if col == 2 else max(16, len(label) + 5)
                       for col, label in enumerate(labels, 1))
        columns(sheet, widths)
        headers(sheet, header_row, labels)
        sheet.freeze_panes = f"{'C' if header_row == 5 else 'B'}{header_row + 1}"
        if header_row == 5:
            title(sheet, "审计调整" if name == AJE else name)
        for key, cell in spec.get("anchors", {}).items():
            position = cell_range(cell)
            anchor(book, key, sheet, f"${get_column_letter(position.min_col)}${position.min_row}")
        for col, field in enumerate(labels, 1):
            if field in {"account_key", "balance_key", "detail_key", "month"}:
                sheet.column_dimensions[get_column_letter(col)].hidden = True
        if spec.get("role") == "summary":
            sheet["A3"] = "上年审定数"
            sheet["B3"] = '=IF(COUNTA(Raw_Prior_TB!$A:$A)>1,"已提供","未提供")'
        if name == AJE:
            sheet["A3"] = "按审定表口径填入已复核调整。正数增加审定数。"
            area = cell_range(spec["data"])
            for row in range(area.min_row, area.max_row + 1):
                sheet.cell(row, 1).number_format = "@"
                sheet.cell(row, 2).number_format = AMOUNT_FORMAT
                sheet.cell(row, 2).fill = PatternFill("solid", fgColor="FFF1CD")
        if "detail_formulas" not in spec:
            continue
        sheet["A6"], sheet["A7"] = "{{DATA_START}}", "{{TOTAL}}"
        for row in (6, 7):
            sheet.row_dimensions[row].height = 22
            for col, label in enumerate(labels, 1):
                cell = sheet.cell(row, col)
                cell.font = font(bold=row == 7)
                cell.border = Border(top=LINE) if row == 7 else Border(bottom=LINE)
                cell.alignment = Alignment(vertical="center", horizontal="left" if col < 3 else "right")
                cell.number_format = ("@" if col < 3 else PERCENT_FORMAT if "占比" in label or "变动率" in label
                                      else "General" if "校验" in label else AMOUNT_FORMAT)
                if row == 7:
                    cell.fill = PatternFill("solid", fgColor="E8EFF6")
    book.save(path)
    book.close()


def batch_mock(folder):
    chart = {
        "1002": ("银行存款", 1), "1002.01": ("人民币基本户", 2), "1002.02": ("人民币结算户", 2),
        "1122": ("应收账款", 1), "1122.01": ("应收账款 / 产品", 2), "1122.02": ("应收账款 / 服务", 2),
        "1601": ("固定资产", 1), "1601.01": ("机器设备", 2), "1601.02": ("电子设备", 2),
        "1602": ("累计折旧", 1), "1602.01": ("机器设备折旧", 2), "1602.02": ("电子设备折旧", 2),
        "2204": ("合同负债", 1),
        "2221": ("应交税费", 1), "2221.01": ("增值税", 2), "2221.01.01": ("增值税 / A 项", 3),
        "2221.01.02": ("增值税 / B 项", 3), "2221.02": ("企业所得税", 2),
        "4001": ("实收资本", 1),
        "6001": ("主营业务收入", 1), "6001.01": ("产品销售", 2), "6001.02": ("服务收入", 2),
        "6602": ("管理费用", 1),
    }
    parties = tuple(f"模拟{label}公司" for label in "甲乙丙丁戊己庚")
    a, b, c, d, e, f, g = parties
    opening = {
        ("1002.01", None): 20000, ("1002.02", None): 5000,
        ("1122.01", a): 1000, ("1122.01", b): 2000, ("1122.01", c): 800,
        ("1122.02", a): 300, ("1122.02", d): 650, ("1122.02", e): 450, ("1122.02", f): 300, ("1122.02", g): 0,
        ("1601.01", None): 10000, ("1601.02", None): 5000, ("1602.01", None): -2000, ("1602.02", None): -1000,
        ("2221.01.01", None): -100, ("2221.01.02", None): -200, ("2221.02", None): -150,
        ("2204", None): 0, ("6001.01", None): 0, ("6001.02", None): 0, ("6602", None): 0,
    }
    opening[("4001", None)] = -sum(opening.values())
    debits, credits = defaultdict(lambda: ZERO), defaultdict(lambda: ZERO)
    gl_rows = []

    def post(month, label, debit, credit, amount):
        amount = Decimal(amount).quantize(Decimal("0.01"))
        debits[debit] += amount
        credits[credit] += amount
        day = f"2025-{month:02d}-15"
        voucher = f"MOCK-{len(gl_rows) // 2 + 1:03d}"
        for key, dr, cr in ((debit, amount, ""), (credit, "", amount)):
            gl_rows.append((day, voucher, label, key[0], dr, cr, key[1] or ""))

    for month in range(1, 13):
        post(month, "模拟赊销", ("1122.01", (a, b, c)[(month - 1) % 3]), ("6001.01", None), 500 + month * 30)
        post(month, "模拟现销", ("1002.01", None), ("6001.02", None), Decimal(230 + month * 7) + Decimal("0.15"))
    post(3, "模拟往来回款", ("1002.02", None), ("1122.01", a), 350)
    post(6, "模拟客商预收", ("1002.02", None), ("1122.02", g), 75)
    post(4, "模拟购置资产", ("1601.01", None), ("1002.01", None), 1800)
    post(12, "模拟计提折旧", ("6602", None), ("1602.01", None), 450)
    post(12, "模拟计提折旧", ("6602", None), ("1602.02", None), 240)
    post(8, "模拟计提税费", ("6602", None), ("2221.01.01", None), 300)
    post(8, "模拟计提税费", ("6602", None), ("2221.01.02", None), 280)
    post(9, "模拟缴纳税费", ("2221.01.01", None), ("1002.02", None), 200)
    post(11, "模拟计提税费", ("6602", None), ("2221.02", None), 250)
    post(12, "模拟缴纳税费", ("2221.02", None), ("1002.02", None), 100)
    with (folder / "mock_gl.csv").open("w", encoding="utf-8-sig", newline="") as stream:
        csv.writer(stream).writerows((GL_HEADERS, *sorted(gl_rows, key=lambda row: (row[0], row[1]))))

    records = []

    def tb_record(code, party, keys):
        name, level = chart[code]
        start = sum((Decimal(opening[key]) for key in keys), ZERO)
        dr, cr = (sum((side[key] for key in keys), ZERO) for side in (debits, credits))
        close = start + dr - cr
        direction = "借" if close > ZERO else "贷" if close < ZERO else "平"
        return (code, name, level, direction, max(start, ZERO), max(-start, ZERO), dr, cr,
                max(close, ZERO), max(-close, ZERO), party)

    for code in chart:
        keys = [key for key in opening if key[0] == code or key[0].startswith(f"{code}.")]
        if (code, None) in opening or len(keys) != sum(key[0] == code for key in keys) or code == "1122.01":
            records.append(tb_record(code, None, keys))
        records.extend(tb_record(code, key[1], (key,)) for key in keys if key[0] == code and key[1])

    for filename, labels, mismatch in (("mock_tb.xlsx", TB_HEADERS, False),
                                       ("mock_prior_tb.xlsx", PRIOR_HEADERS, False),
                                       ("mock_prior_tb_mismatch.xlsx", PRIOR_HEADERS, True)):
        book = Workbook()
        sheet = book.active
        sheet.title = "模拟科目余额表" if labels == TB_HEADERS else "模拟上年审定余额表"
        headers(sheet, 1, labels)
        columns(sheet, (18, 27, *([17] * (len(labels) - 2))))
        for record in records:
            if labels == TB_HEADERS:
                sheet.append(record)
            else:
                code, name, _, _, debit, credit, *_, party = record
                if mismatch and (code, party) == ("1122.01", a):
                    debit -= Decimal("100.00")
                sheet.append((code, name, debit, credit, party))
        book.save(folder / filename)
        book.close()


def main():
    (ROOT / "templates").mkdir(exist_ok=True)
    folder = ROOT / "data" / "mock"
    folder.mkdir(parents=True, exist_ok=True)
    template(ROOT / "templates" / "ar_template.xlsx")
    mock(folder)
    config = safe_load((ROOT / "templates" / "blueprint.yaml").read_text(encoding="utf-8"))["workpapers"]
    for code, name in DEMO_TEMPLATES.items():
        batch_template(ROOT / "templates" / name, config[code])
    batch_mock(folder)


if __name__ == "__main__":
    main()
