"""Four audit gates. Reconciliation uses exact Decimal cents and produces local evidence."""

from dataclasses import dataclass, field
from decimal import Decimal

from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from core.fill_workpaper import DETAIL, SUMMARY, literal
from core.schema import TB_AMOUNT_FIELDS, ZERO

REPORT = "00_自动勾稽与校验报告"
DIFF_FORMAT = '#,##0.00;(#,##0.00);0.00'
FAIL_LIMIT = 100


@dataclass
class Gate:
    name: str
    scope: str
    failed: int = 0
    warnings: int = 0
    findings: list = field(default_factory=list)

    def add(self, position, message, difference=None, *, warning=False):
        if warning:
            self.warnings += 1
        else:
            self.failed += 1
        if warning or len(self.findings) < FAIL_LIMIT:
            self.findings.append((position, message, difference, warning))

    def compare(self, position, message, actual, expected):
        difference = actual - expected
        if difference != ZERO:
            self.add(position, message, difference)

    @property
    def status(self):
        return "FAILED" if self.failed else "PASS"


def check_conservation(data, source_count, source):
    gate = Gate("1. 清洗前后守恒", "原始 TB 和标准 TB 的行数及六列借贷金额。容差 0.00。")
    row_difference = len(data.tb) - source_count
    if row_difference:
        gate.add("Standard_TB", f"清洗前后数据行数不一致（行数差：{row_difference:+d}）")
    controls = {}
    for name in TB_AMOUNT_FIELDS:
        original = source[name]
        standard = sum((getattr(row, name) for row in data.tb), ZERO)
        controls[name] = original, standard, standard - original
        gate.compare(name, "清洗前后金额不守恒", standard, original)
    return gate, controls


def check_coverage(data, coverage):
    gate = Gate("2. 未映射兜底", "保留所有标准化记录，逐一告警未纳入当前应收底稿的科目。")
    for code in coverage.uncovered_codes:
        gate.add(code, "TB 科目未纳入当前应收底稿，保留在 Raw_TB", warning=True)
    for code in sorted({row.account_code for row in data.gl} - coverage.tb_codes):
        gate.add(code, "GL 科目在 TB 中缺失，保留在 Standard_GL", warning=True)
    return gate


def check_cross(data, coverage, layout, calculation):
    gate = Gate("3. 跨表勾稽", "审定表未审数、明细期末合计及 Raw_TB 一级科目期末净额。")
    root_amount = None
    if len(coverage.root_rows) != 1:
        gate.add("Raw_TB", "目标一级科目须有且仅有一条无辅助核算的余额行")
    else:
        root = data.tb[coverage.root_rows[0] - 2]
        root_amount = root.close_debit - root.close_credit
    if not coverage.detail_rows:
        gate.add(DETAIL, "未匹配到可供勾稽的明细行")

    def numeric(sheet, coordinate):
        cache = calculation.caches.get((sheet, coordinate))
        number = cache.number() if cache else None
        if number is None:
            gate.add(f"{sheet}!{coordinate}", "缺少可用于勾稽的重算金额")
        return number

    summary = numeric(SUMMARY, f"B{layout.summary_start + 3}")
    detail = numeric(DETAIL, f"F{layout.detail_total}")
    if root_amount is not None:
        if summary is not None:
            gate.compare(SUMMARY, "未审数与一级科目期末净额不一致", summary, root_amount)
        if detail is not None:
            gate.compare(DETAIL, "明细合计与一级科目期末净额不一致", detail, root_amount)
    if summary is not None and detail is not None:
        gate.compare(SUMMARY, "未审数与明细合计不一致", summary, detail)

    for line, source in enumerate(coverage.detail_rows, layout.detail_start):
        close = numeric(DETAIL, f"F{line}")
        if close is not None:
            record = data.tb[source - 2]
            gate.compare(f"{DETAIL}!F{line}", "逐行滚转与源明细期末净额不一致",
                         close, record.close_debit - record.close_credit)
    return gate, (summary, detail, root_amount)


def check_excel(layout, actual, calculation):
    gate = Gate("4. 公式健康度", "展示页预期单元格须为公式，重算后检查错误值和缺失缓存。")
    for sheet, address in sorted(layout.required_formulas()):
        formula = actual.get((sheet, address))
        if not isinstance(formula, str) or not formula.startswith("="):
            gate.add(f"{sheet}!{address}", "预期公式缺失或单元格为静态值")
    if calculation.backend in {"none", "unavailable"}:
        gate.add("公式重算", calculation.note or "缺少受支持的本地重算结果")
    for position, issue in calculation.issues:
        if calculation.backend not in {"none", "unavailable"} or issue == "公式包含 #REF!":
            gate.add(position, issue)
    return gate


def run_gates(data, source, coverage, layout, actual, calculation):
    conservation, controls = check_conservation(data, *source)
    cross, balances = check_cross(data, coverage, layout, calculation)
    gates = (conservation, check_coverage(data, coverage), cross, check_excel(layout, actual, calculation))
    return gates, controls, balances


def money(value):
    if value is None:
        return "N/A"
    return f"{value:.2f}" if value == value.quantize(Decimal("0.01")) else format(value, "f")


def console_lines(gates, controls, balances, *, show_amounts):
    lines = [f"GATE {index} {gate.status} | {gate.name[3:]} | errors={gate.failed} warnings={gate.warnings}"
             for index, gate in enumerate(gates, 1)]
    if show_amounts:
        totals = [f"  {name}: raw={money(raw)} standard={money(standard)} diff={money(diff)}"
                  for name, (raw, standard, diff) in controls.items()]
        lines[1:1] = totals
        summary, detail, root = balances
        difference = summary - detail if summary is not None and detail is not None else None
        root_diff = detail - root if detail is not None and root is not None else None
        lines.insert(-1, f"  unaudited={money(summary)} detail={money(detail)} raw_level_1={money(root)} "
                         f"unaudited_minus_detail={money(difference)} detail_minus_raw={money(root_diff)}")
    lines.append("PASS" if all(gate.status == "PASS" for gate in gates) else "FAILED")
    return lines


def write_report(book, gates, controls, balances, calculation):
    sheet = book.create_sheet(REPORT, 0)
    sheet.sheet_view.showGridLines = False
    sheet.sheet_properties.tabColor = "344A62"
    passed = all(gate.status == "PASS" for gate in gates)
    sheet.append([])
    sheet.append(["自动勾稽与校验报告"])
    sheet.append(["总体结论", "PASS" if passed else "FAILED"])
    sheet.append(["重算后端", calculation.backend, None, "记录生成时的校验结果。编辑底稿后应重新复核。",
                  calculation.note or None])
    sheet.append([])
    sheet.append(["门禁", "结果", "异常数", "告警数", "检查范围"])
    for gate in gates:
        sheet.append([gate.name, gate.status, gate.failed, gate.warnings, gate.scope])
    sheet.append([])
    sheet.append(["TB 金额字段", "原始合计", "标准合计", "差额", "比较口径"])
    for name, (raw, standard, diff) in controls.items():
        sheet.append([name, raw, standard, diff, "差额须为 0.00"])
        for cell in sheet[sheet.max_row][1:4]:
            cell.number_format = DIFF_FORMAT
    sheet.append([])
    sheet.append(["审定表未审数", "明细表期末合计", "Raw_TB 一级期末净额", "审定表－明细表", "明细表－Raw_TB"])
    summary, detail, root = balances
    sheet.append([summary, detail, root, summary - detail if summary is not None and detail is not None else None,
                  detail - root if detail is not None and root is not None else None])
    for cell in sheet[sheet.max_row]:
        cell.number_format = DIFF_FORMAT
    sheet.append([])
    sheet.append(["位置或科目", "门禁", "类别", "差异", "问题说明"])
    for number, gate in enumerate(gates, 1):
        for position, message, difference, warning in gate.findings:
            sheet.append([None, number, "告警" if warning else "异常", difference, message])
            literal(sheet.cell(sheet.max_row, 1), position)
            if difference is not None:
                sheet.cell(sheet.max_row, 4).number_format = DIFF_FORMAT
        omitted = gate.failed - sum(not warning for *_, warning in gate.findings)
        if omitted:
            sheet.append([None, number, "异常", None, f"另有 {omitted} 项异常。"])
    for index, width in enumerate((32, 22, 25, 24, 64), 1):
        sheet.column_dimensions[get_column_letter(index)].width = width
    for row in sheet.iter_rows(min_row=3):
        for cell in row:
            cell.font = Font(name="Microsoft YaHei", size=10, color="24384A")
            cell.alignment = Alignment(vertical="center")
        sheet.row_dimensions[row[0].row].height = 21
    if calculation.note:
        for cell in sheet[4][3:5]:
            cell.alignment = Alignment(vertical="center", wrap_text=True)
        sheet.row_dimensions[4].height = 35
    sheet["A2"].font = Font(name="Microsoft YaHei", size=15, bold=True, color="18324A")
    for index in (6, 12, 20, 23):
        for cell in sheet[index]:
            cell.fill = PatternFill("solid", fgColor="263F56")
            cell.font = Font(name="Microsoft YaHei", size=10, bold=True, color="FFFFFF")
    for cell in (sheet["B3"], *(sheet[f"B{row}"] for row in range(7, 11))):
        if cell.value == "FAILED":
            cell.fill = PatternFill("solid", fgColor="FCE4D9")
            cell.font = Font(name="Microsoft YaHei", size=10, bold=True, color="9C2D1C")
    sheet.freeze_panes = "B7"
    book.active = sheet
