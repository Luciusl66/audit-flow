"""Build and verify the project trial balance without duplicating ledger totals."""

from collections import defaultdict
from dataclasses import dataclass

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from core.accounts import amount_sides
from core.excel_writer import apply_openpyxl, literal
from core.fillers.base_filler import BalanceCheck, rounded
from core.reconcile import DIFF_FORMAT, Gate, check_calculated
from core.schema import AuditError, PRIOR_AMOUNT_FIELDS, TB_AMOUNT_FIELDS, TB_SHEET, ZERO
from core.workpaper_data import ACCOUNT_KEY_PREFIX, AMOUNT_FORMAT, TABLE_FORMATS, account_key, prior_lookup, raw_tables, sum_by_key

MASTER_FILE = "00_Master_TB.xlsx"
MASTER_SHEET = "Master_TB"
PROPOSALS = "建议分录"
STAGES = ("上年审定", "本期期初", "本期发生", "本期未审", "建议调整", "调整后")
PAIRS = tuple(zip("DFHJLN", "EGIKMO", strict=True))
MASTER_HEADERS = ("科目编码", "科目名称", "本次底稿", *(f"{stage}{side}" for stage in STAGES for side in ("借方", "贷方")))
PROPOSAL_HEADERS = ("异常位置", "建议类型", "借方科目", "贷方科目", "借方编码", "贷方编码",
                    "待确认金额", "借方键", "贷方键", "纳入测算金额", "纳入状态")


@dataclass(frozen=True)
class MasterLayout:
    formulas: frozenset
    checks: tuple[BalanceCheck, ...]
    differences: tuple[BalanceCheck, ...]
    duplicate_roots: tuple[str, ...]
    issues: tuple

    def required_formulas(self):
        return self.formulas


def format_headers(sheet, row, labels):
    for column, label in enumerate(labels, 1):
        cell = sheet.cell(row, column, label)
        cell.fill = PatternFill("solid", fgColor="263F56")
        cell.font = Font(name="Microsoft YaHei", bold=True, color="FFFFFF")
        cell.alignment = Alignment(horizontal="center", vertical="center")
    sheet.row_dimensions[row].height = 25
    sheet.sheet_view.showGridLines = False


def set_values(sheet, row, values):
    for column, value in enumerate(values, 1):
        cell = sheet.cell(row, column)
        literal(cell, value)
        if isinstance(value, str):
            cell.alignment = Alignment(horizontal="left", vertical="center")


def title(sheet, label, note):
    sheet["A2"], sheet["A3"] = label, note
    sheet["A2"].font = Font(name="Microsoft YaHei", size=15, bold=True, color="18324A")
    sheet.row_dimensions[2].height = 25


def add_formula(sheet, address, formula, registered, *, amount=True):
    cell = sheet[address]
    cell.value = formula
    if amount:
        cell.number_format = AMOUNT_FORMAT
    registered.add((sheet.title, address))


def write_proposals(book, ledger, findings, names, last, registered, checks):
    sheet = book.create_sheet(PROPOSALS)
    title(sheet, "待复核的建议分录", "借贷科目编码均可唯一核对且金额明确时，纳入本期期末模拟。")
    format_headers(sheet, 5, PROPOSAL_HEADERS)
    sheet.freeze_panes = "E6"
    for col, width in zip("ABCDEFGHIJK", (26, 23, 36, 36, 16, 16, 18, 14, 14, 19, 16), strict=True):
        sheet.column_dimensions[col].width = width
    sheet.column_dimensions["H"].hidden = sheet.column_dimensions["I"].hidden = True

    catalog = defaultdict(set)
    for code in ledger.current.roots:
        catalog[ledger.name(code)].add(code)
    for code, name in (names or {}).items():
        if code in ledger.current.roots:
            catalog[name].add(code)

    def resolve(name):
        matches = catalog.get(name)
        return next(iter(matches)) if matches and len(matches) == 1 else None

    debits, credits = defaultdict(lambda: ZERO), defaultdict(lambda: ZERO)
    candidates = tuple(dict.fromkeys(finding for finding in findings if finding.debit or finding.credit
                                     or finding.entry_type.startswith("RJE")))
    master_keys = f"{MASTER_SHEET}!$P$6:$P${last}"
    for row, finding in enumerate(candidates, 6):
        debit, credit = resolve(finding.debit), resolve(finding.credit)
        amount = finding.draft_amount
        set_values(sheet, row, (finding.position, finding.entry_type, finding.debit, finding.credit, debit, credit, amount))
        sheet[f"G{row}"].number_format = AMOUNT_FORMAT
        for key_column, code_column in (("H", "E"), ("I", "F")):
            code = f"{code_column}{row}"
            key = f'"{ACCOUNT_KEY_PREFIX}"&{code}'
            add_formula(sheet, f"{key_column}{row}",
                        f'=IF(AND({code}<>"",SUMPRODUCT(--EXACT({master_keys},{key}))=1),{key},"")',
                        registered, amount=False)
        eligible = f'AND(ISNUMBER(G{row}),G{row}>0,H{row}<>"",I{row}<>"",H{row}<>I{row},B{row}<>"期初滚调待复核")'
        add_formula(sheet, f"J{row}", f"=IF({eligible},ROUND(G{row},2),0)", registered)
        add_formula(sheet, f"K{row}", f'=IF(J{row}>0,"纳入模拟","待复核")', registered, amount=False)
        projected = ZERO
        if debit and credit and debit != credit and amount is not None and amount > ZERO and finding.rule_id != "OPENING_ROLLFORWARD":
            projected = amount
            debits[debit] += projected
            credits[credit] += projected
        checks.append(BalanceCheck((PROPOSALS, f"J{row}"), projected, "建议分录纳入金额与科目及金额证据不一致"))
    sheet.auto_filter.ref = f"A5:K{5 + len(candidates)}"
    return debits, credits, max(6, 5 + len(candidates))


def build_master(book: Workbook, ledger, owners, findings, names):
    data = ledger.data
    sheet = book.active
    sheet.title = MASTER_SHEET
    title(sheet, "项目总试算平衡表", "本次底稿列显示调度范围。试算涵盖完整账套，调整后为建议分录模拟。")
    format_headers(sheet, 5, MASTER_HEADERS)
    sheet.freeze_panes = "D6"
    for index, width in enumerate((15, 27, 15, *([20] * 12)), 1):
        sheet.column_dimensions[get_column_letter(index)].width = width
    sheet.column_dimensions["P"].hidden = True
    codes = sorted(ledger.current.roots.keys() | ledger.prior.roots.keys())
    total = len(codes) + 6
    if total + 2 > 1_048_576:
        raise AuditError("Master TB 超出 Excel 行数上限。")
    registered, checks = set(), []
    issues = list(ledger.grain_issues())
    totals = dict.fromkeys("DEFGHIJKLMNO", ZERO)
    debits, credits, last = write_proposals(book, ledger, findings, names, total - 1, registered, checks)
    for row, code in enumerate(codes, 6):
        set_values(sheet, row, (code, ledger.name(code), owners.get(code, "未调度")))
        literal(sheet[f"P{row}"], account_key(code))
        count, dr, cr = prior_lookup(f"A{row}", len(data.prior_tb or ()))
        for col, expression in (("D", dr), ("E", cr)):
            add_formula(sheet, f"{col}{row}", f'=IF({count}=0,"",ROUND({expression},2))', registered)
        for col, source in zip("FGHIJK", "EFGHIJ", strict=True):
            expression = sum_by_key(TB_SHEET, source, "M", f"$P{row}", len(data.tb))
            add_formula(sheet, f"{col}{row}", rounded(expression), registered)
        for col, key in (("L", "H"), ("M", "I")):
            expression = f"SUMIFS('{PROPOSALS}'!$J$6:$J${last},'{PROPOSALS}'!${key}$6:${key}${last},$P{row})"
            add_formula(sheet, f"{col}{row}", rounded(expression), registered)
        add_formula(sheet, f"N{row}", rounded(f"MAX(J{row}-K{row}+L{row}-M{row},0)"), registered)
        add_formula(sheet, f"O{row}", rounded(f"MAX(K{row}-J{row}+M{row}-L{row},0)"), registered)

        old = ledger.prior.roots.get(code, ())
        current = ledger.current.roots.get(code, ())
        prior = amount_sides(data.prior_tb, old, *PRIOR_AMOUNT_FIELDS) if old else ("", "")
        tb = tuple(sum((getattr(data.tb[line - 2], field) for line in current), ZERO) for field in TB_AMOUNT_FIELDS)
        opening_dr, opening_cr, period_dr, period_cr, closing_dr, closing_cr = tb
        rollforward = closing_dr - closing_cr - (opening_dr - opening_cr + period_dr - period_cr)
        if rollforward != ZERO:
            issues.append((f"{MASTER_SHEET}!J{row} / {code}", "期末净额与期初净额加本期净发生额不一致", rollforward))
        net = tb[4] - tb[5] + debits[code] - credits[code]
        expected = (*prior, *tb, debits[code], credits[code], max(net, ZERO), max(-net, ZERO))
        for col, amount in zip("DEFGHIJKLMNO", expected, strict=True):
            checks.append(BalanceCheck((MASTER_SHEET, f"{col}{row}"), amount, "Master TB 科目金额与独立控制数不一致"))
            if amount != "":
                totals[col] += amount

    sheet[f"A{total}"], sheet[f"A{total+1}"], sheet[f"A{total+2}"] = "合计", "借贷差额", "平衡校验"
    differences = []
    for stage, (debit, credit) in zip(STAGES, PAIRS, strict=True):
        no_prior = stage == STAGES[0] and data.prior_tb is None
        for col in (debit, credit):
            scope = f"{col}6:{col}{total - 1}"
            formula = f'=IF(COUNT({scope})=0,"",ROUND(SUM({scope}),2))' if stage == STAGES[0] else rounded(f"SUM({scope})")
            add_formula(sheet, f"{col}{total}", formula, registered)
            checks.append(BalanceCheck((MASTER_SHEET, f"{col}{total}"), "" if no_prior else totals[col],
                                       f"{stage}合计与独立控制数不一致"))
        dr, cr, diff = f"{debit}{total}", f"{credit}{total}", f"{debit}{total + 1}"
        skipped = f'{dr}=""'
        difference = f'=IF({skipped},"",ROUND({dr}-{cr},2))' if stage == STAGES[0] else rounded(f"{dr}-{cr}")
        add_formula(sheet, diff, difference, registered)
        sheet[diff].number_format = DIFF_FORMAT
        status = f"{debit}{total + 2}"
        add_formula(sheet, status, f'=IF({diff}="","SKIPPED",IF({diff}=0,"PASS","FAILED"))', registered, amount=False)
        independent_status = "SKIPPED" if no_prior else "PASS" if totals[debit] == totals[credit] else "FAILED"
        checks.append(BalanceCheck((MASTER_SHEET, status), independent_status, f"{stage}平衡状态与独立控制数不一致"))
        differences.append(BalanceCheck((MASTER_SHEET, diff), "" if no_prior else ZERO, f"{stage}借贷试算差额必须为 0.00"))
    for cell in sheet[total][:15]:
        cell.fill = PatternFill("solid", fgColor="E8EFF6")
        cell.font = Font(name="Microsoft YaHei", bold=True, color="24384A")
    sheet.auto_filter.ref = f"A5:O{total - 1}"

    for name, fields, rows, count in raw_tables(data, ledger, include_gl=False):
        raw = book.create_sheet(name)
        format_headers(raw, 1, fields)
        raw.freeze_panes = "B2"
        for col in range(1, len(fields) + 1):
            raw.column_dimensions[get_column_letter(col)].width = 23
        for col in range(len(fields) - (3 if name == TB_SHEET else 1), len(fields) + 1):
            raw.column_dimensions[get_column_letter(col)].hidden = True
        if count >= 1_048_576:
            raise AuditError("Master TB 的标准数据超出 Excel 行数上限。")
        apply_openpyxl(book, {"type": "table", "sheet": name, "first": 2, "count": count,
                              "formats": [TABLE_FORMATS.get(field, "@") for field in fields]}, rows)
    duplicates = tuple(code for code in ledger.current.roots if len(ledger.current.exact.get((code, None), ())) > 1)
    return MasterLayout(frozenset(registered), tuple(checks), tuple(differences), duplicates, tuple(issues))


def check_master(layout, calculation):
    gate = Gate("3. Master TB 总试算平衡", "全账套核对总账科目汇总粒度、余额滚动及各阶段借贷。容差 0.00。")
    for code in layout.duplicate_roots:
        gate.add(f"{TB_SHEET} / {code}", "同一总账科目存在多条无辅助核算的汇总行")
    for position, issue, difference in layout.issues:
        gate.add(position, issue, difference)
    check_calculated(gate, calculation, (*layout.checks, *layout.differences))
    return gate
