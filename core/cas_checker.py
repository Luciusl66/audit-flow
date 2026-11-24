"""Offline CAS screening. The scanner reads standardized TB rows and returns findings."""

from collections import defaultdict
from dataclasses import dataclass
from decimal import Decimal
from pathlib import Path

from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from yaml import YAMLError, safe_load

from core.fill_workpaper import leaf_rows, literal
from core.schema import AuditError, StandardTB, TB_SHEET, ZERO, account_code

CAS_SHEET = "01_准则异常与调整建议"
DEFAULT_RULES = Path(__file__).resolve().parents[1] / "rules" / "cas_general.yaml"
RULE_IDS = ("AR_CREDIT", "AP_DEBIT", "CASH_DIRECTION", "PPE_DIRECTION", "PL_UNCLOSED", "AR_ALLOWANCE")
POSTING_RULES = {"AR_CREDIT", "AP_DEBIT", "AR_ALLOWANCE"}


@dataclass(frozen=True)
class Rule:
    prefixes: tuple[str, ...]
    label: str
    basis: str
    review: str
    debit: str = ""
    credit: str = ""


@dataclass(frozen=True)
class Finding:
    rule_id: str
    position: str
    subject: str
    balance: Decimal
    detail: str
    entry_type: str = "待核对"
    debit: str = ""
    credit: str = ""
    draft_amount: Decimal | None = None
    gross: Decimal | None = None


def code_list(values):
    if not isinstance(values, (list, tuple)) or not values:
        raise ValueError
    return tuple(dict.fromkeys(account_code(value) for value in values))


def load_rules(path, account_pairs):
    try:
        with path.open(encoding="utf-8") as file:
            config = safe_load(file)
        rules = {}
        for key in RULE_IDS:
            source = config["rules"][key]
            labels = [source[field] for field in ("label", "basis", "review")]
            posting = [source[field] for field in ("debit", "credit")] if key in POSTING_RULES else []
            if not all(isinstance(value, str) and value.strip() for value in labels + posting):
                raise ValueError
            rules[key] = Rule(code_list(source["prefixes"]), *labels, *posting)
        pairs = {}
        for mapping in (config.get("allowance_pairs", {}), account_pairs or {}):
            for prefix, receivables in mapping.items():
                pairs[account_code(prefix)] = tuple(sorted(code_list(receivables)))
        return rules, pairs
    except (OSError, YAMLError, KeyError, TypeError, AttributeError, ValueError):
        raise AuditError("CAS 规则配置无效。请检查规则文件或已核对的备抵科目配对。") from None


def matches(record, prefix):
    code = record.account_code
    return code == prefix or code.startswith(prefix) and (record.level > 1 or len(prefix) == 1)


def net(record: StandardTB) -> Decimal:
    return record.close_debit - record.close_credit


def subjects(rows, prefixes):
    return leaf_rows((index, record) for index, record in rows
                     if any(matches(record, prefix) for prefix in prefixes))


def reclass_findings(rows, key, rule, ar_account):
    prefixes = (*rule.prefixes, ar_account) if key == "AR_CREDIT" else rule.prefixes
    for index, record in subjects(rows, prefixes):
        balance = net(record)
        is_reversed = balance < ZERO if key == "AR_CREDIT" else balance > ZERO
        if is_reversed:
            evidence = ("按该客商核查往来款性质。" if record.aux_name else
                        "当前行无客商标识，先取得往来明细后确定重分类金额。")
            yield Finding(key, f"{TB_SHEET}!{index}", subject(record), balance, evidence, "RJE 待确认",
                          rule.debit, rule.credit, abs(balance) if record.aux_name else None)


def subject(record):
    return " / ".join(filter(None, (record.account_code, record.account_name, record.aux_name)))


def direction_findings(rows, key, rule):
    for index, record in subjects(rows, rule.prefixes):
        balance = net(record)
        issues = []
        if balance < ZERO:
            issues.append("期末为贷方净额")
        if record.close_debit < ZERO:
            issues.append("期末借方栏为负数")
        if record.close_credit < ZERO:
            issues.append("期末贷方栏为负数")
        if issues:
            yield Finding(key, f"{TB_SHEET}!{index}", subject(record), balance,
                          "。".join(issues) + "。", "AJE 待查")


def closing_findings(rows, rule):
    for index, record in subjects(rows, rule.prefixes):
        balance = net(record)
        if balance != ZERO:
            debit, credit = ((record.account_name, "本年利润") if balance < ZERO else
                             ("本年利润", record.account_name))
            yield Finding("PL_UNCLOSED", f"{TB_SHEET}!{index}", subject(record), balance,
                          "导出结转状态待确认。", "结转分录待确认", debit, credit)


def check_allowance_scope(group, gross_rows, prefixes, rule, auxiliary=None):
    balance = sum((net(record) for _, record in group), ZERO)
    if balance >= ZERO:
        return
    position = f"{TB_SHEET}!{group[0][0]}（组内 {len(group)} 行）"
    codes = "、".join(sorted({record.account_code for _, record in group}))
    label = f"{codes}（配对应收 {'、'.join(prefixes)}）"
    if auxiliary is not None:
        label += f" / {auxiliary}"
    if not gross_rows:
        yield Finding("AR_ALLOWANCE", position, label, balance, "未取得配对范围内的应收原值。")
        return
    gross = sum((max(net(record), ZERO) for _, record in gross_rows), ZERO)
    if -balance > gross:
        excess = -balance - gross
        yield Finding("AR_ALLOWANCE", position, label, balance,
                      f"准备贷方净额超过配对范围内的应收借方原值 {excess:.2f}。原值按同范围明细筛查。",
                      "AJE 待重算", rule.debit, rule.credit, gross=gross)


def allowance_findings(rows, rule, pairs):
    key = "AR_ALLOWANCE"
    grouped = defaultdict(list)
    ordered = sorted(pairs, key=len, reverse=True)
    allowance = subjects(rows, (*rule.prefixes, *pairs))
    for index, record in allowance:
        prefix = next((code for code in ordered if matches(record, code)), None)
        if prefix is not None:
            grouped[pairs[prefix]].append((index, record))
        elif net(record) < ZERO:
            yield Finding(key, f"{TB_SHEET}!{index}", subject(record), net(record),
                          "准备归属缺少已确认配对。先按对应应收款项或组合拆分，待补足原值。")
    for prefixes, group in grouped.items():
        gross_rows = list(subjects(rows, prefixes))
        if not gross_rows or any(record.aux_name is None for _, record in group + gross_rows):
            yield from check_allowance_scope(group, gross_rows, prefixes, rule)
            continue
        allowance_by_party = defaultdict(list)
        gross_by_party = defaultdict(list)
        for index, record in group:
            allowance_by_party[record.aux_name].append((index, record))
        for index, record in gross_rows:
            gross_by_party[record.aux_name].append((index, record))
        for party, party_group in allowance_by_party.items():
            yield from check_allowance_scope(party_group, gross_by_party.get(party, []), prefixes, rule, party)


def scan_tb(tb, ar_account, *, account_pairs=None, rules_path=DEFAULT_RULES):
    """Screen the full Raw_TB population. Findings need accounting review before posting."""
    rules, pairs = load_rules(rules_path, account_pairs)
    rows = tuple(enumerate(tb, 2))
    findings = []
    for key in ("AR_CREDIT", "AP_DEBIT"):
        findings.extend(reclass_findings(rows, key, rules[key], ar_account))
    for key in ("CASH_DIRECTION", "PPE_DIRECTION"):
        findings.extend(direction_findings(rows, key, rules[key]))
    findings.extend(closing_findings(rows, rules["PL_UNCLOSED"]))
    findings.extend(allowance_findings(rows, rules["AR_ALLOWANCE"], pairs))
    return rules, tuple(findings)


def write_cas_sheet(book, rules, findings):
    sheet = book.create_sheet(CAS_SHEET, 1)
    sheet.sheet_view.showGridLines = False
    sheet.sheet_properties.tabColor = "9B682F"
    sheet["A2"] = "准则异常与调整建议"
    sheet["A2"].font = Font(name="Microsoft YaHei", size=15, bold=True, color="18324A")
    sheet["A3"], sheet["B3"] = "命中项", len(findings)
    sheet["A4"] = "按报告期与适用准则复核。分录为草案，期末净额借正贷负。"
    sheet.append(("异常类型", "位置", "科目 / 客商", "期末净额", "应收原值", "异常及复核要点",
                  "准则依据", "建议类型", "借方", "贷方", "待确认金额"))
    for finding in findings:
        rule = rules[finding.rule_id]
        record = (rule.label, finding.position, finding.subject, finding.balance, finding.gross,
                  f"{finding.detail} {rule.review}", rule.basis, finding.entry_type,
                  finding.debit, finding.credit, finding.draft_amount)
        row = sheet.max_row + 1
        for column, value in enumerate(record, 1):
            cell = sheet.cell(row, column)
            literal(cell, value)
            cell.font = Font(name="Microsoft YaHei", size=10, color="24384A")
            cell.alignment = Alignment(vertical="center", wrap_text=isinstance(value, str))
            if column in (4, 5, 11):
                cell.number_format = '#,##0.00;(#,##0.00);0.00'
        sheet.row_dimensions[row].height = 78
    if not findings:
        sheet["A6"] = "当前扫描规则未命中异常。"
    for column, width in enumerate((27, 20, 32, 16, 16, 55, 40, 19, 19, 19, 16), 1):
        sheet.column_dimensions[get_column_letter(column)].width = width
    for cell in sheet[5]:
        cell.fill = PatternFill("solid", fgColor="263F56")
        cell.font = Font(name="Microsoft YaHei", size=10, bold=True, color="FFFFFF")
        cell.alignment = Alignment(vertical="center", wrap_text=True)
    sheet.row_dimensions[5].height = 24
    sheet.freeze_panes = "C6"
    sheet.auto_filter.ref = f"A5:K{max(5, 5 + len(findings))}"
