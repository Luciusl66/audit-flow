"""Offline CAS screening. The scanner reads standardized TB rows and returns findings."""

from collections import defaultdict
from dataclasses import dataclass
from decimal import Decimal
from string import Template

from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from yaml import YAMLError, safe_load

from core.cas_routing import (ROOT, CasRoute, account_hierarchy, confirmed_account_names,
                              load_routed_rules, matches_account, route_cas)
from core.fill_workpaper import leaf_rows, literal
from core.schema import AuditError, GL_SHEET, StandardTB, TB_SHEET, ZERO, account_code

CAS_SHEET = "01_准则异常与调整建议"
DEFAULT_RULES = ROOT / "rules" / "cas_general.yaml"
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
    balance: Decimal | None
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


def scan_tb(tb, ar_account, *, account_pairs=None, rules_path=DEFAULT_RULES, exclude_reclass=()):
    """Screen the full Raw_TB population. Findings need accounting review before posting."""
    rules, pairs = load_rules(rules_path, account_pairs)
    rows = tuple(enumerate(tb, 2))
    findings = []
    for key in ("AR_CREDIT", "AP_DEBIT"):
        findings.extend(finding for finding in reclass_findings(rows, key, rules[key], ar_account)
                        if finding.position not in exclude_reclass)
    for key in ("CASH_DIRECTION", "PPE_DIRECTION"):
        findings.extend(direction_findings(rows, key, rules[key]))
    findings.extend(closing_findings(rows, rules["PL_UNCLOSED"]))
    findings.extend(allowance_findings(rows, rules["AR_ALLOWANCE"], pairs))
    return rules, tuple(findings)


@dataclass(frozen=True)
class RoutedScan:
    route: CasRoute
    rules: dict[str, Rule]
    findings: tuple[Finding, ...]
    deferred_rule_ids: tuple[str, ...]


def voucher_differences(gl, tolerance):
    totals = {}
    for index, record in enumerate(gl, 2):
        key = (record.date, record.voucher_no)
        first, total = totals.get(key, (index, ZERO))
        totals[key] = (first, total + record.debit - record.credit)
    for index, difference in totals.values():
        if abs(difference) > tolerance:
            yield f"{GL_SHEET}!{index}", "凭证借贷平衡", None, {"difference": f"{difference:.2f}"}


def post_close_balances(tb, selected_account, tolerance, *, root_only=False):
    selected = [(i, record) for i, record in enumerate(tb, 2) if selected_account(record.account_code)]
    if root_only:
        roots = [(i, record) for i, record in selected if record.level == 1 and record.aux_name is None]
        if selected and not roots:
            yield None
        selected = roots
    for index, record in leaf_rows(selected):
        balance = net(record)
        if abs(balance) > tolerance:
            debit, credit = ((record.account_name, "本年利润") if balance < ZERO else
                             ("本年利润", record.account_name))
            variables = {"account_code": record.account_code, "balance": f"{balance:.2f}",
                         "debit_account": debit, "credit_account": credit, "amount": f"{abs(balance):.2f}"}
            yield f"{TB_SHEET}!{index}", subject(record), balance, variables


def direction_observations(tb, selected_account, condition, tolerance):
    rows = ((i, row) for i, row in enumerate(tb, 2) if selected_account(row.account_code))
    for index, row in leaf_rows(rows):
        balance = net(row)
        if abs(balance) <= tolerance or condition.get("require_auxiliary") and not row.aux_name:
            continue
        direction = "credit" if balance < ZERO else "debit"
        matched = (direction == condition["trigger_direction"] if "trigger_direction" in condition
                   else direction != condition["expected_direction"])
        if matched:
            yield f"{TB_SHEET}!{index}", subject(row), balance, {
                "account_code": row.account_code, "balance": f"{balance:.2f}"}


def summary_observations(gl, selected_account, keywords):
    for index, row in enumerate(gl, 2):
        if selected_account(row.account_code):
            keyword = next((word for word in keywords if word in row.summary), None)
            if keyword is not None:
                yield f"{GL_SHEET}!{index}", row.account_code, None, {
                    "account_code": row.account_code, "keyword": keyword}


def contra_observations(tb, selected_account, gross_account, condition, tolerance, pairs):
    rows = tuple(enumerate(tb, 2))
    allowance_rows = leaf_rows((index, row) for index, row in rows if selected_account(row.account_code))
    ordered = sorted(pairs, key=len, reverse=True)
    groups = defaultdict(list)
    by_auxiliary = condition["grouping"] == "auxiliary"
    for index, row in allowance_rows:
        key = next((prefix for prefix in ordered if matches(row, prefix)), None)
        if key is not None:
            groups[pairs[key], row.aux_name if by_auxiliary else None].append((index, row))
        elif net(row) < ZERO:
            yield None
    for (codes, auxiliary), group in groups.items():
        provision = -sum((net(row) for _, row in group), ZERO)
        if provision <= ZERO:
            continue
        original = {
            index: row for index, row in subjects(rows, codes)
            if gross_account(row.account_code)
            and (not by_auxiliary or row.aux_name == auxiliary)
        }
        if ((by_auxiliary and auxiliary is None)
                or any(not any(matches(row, code) for row in original.values()) for code in codes)):
            yield None
            continue
        gross = sum((max(net(row), ZERO) for row in original.values()), ZERO)
        excess = provision - gross
        if excess > tolerance:
            index, row = group[0]
            yield f"{TB_SHEET}!{index}（组内 {len(group)} 行）", subject(row), -provision, {
                "balance": f"{-provision:.2f}", "gross": f"{gross:.2f}", "excess": f"{excess:.2f}"}


def routed_observations(data, rule, pairs, account_matches):
    condition = rule["condition"]
    check = rule["check_type"]
    def selected_account(code):
        return account_matches(rule["account_pattern"], code)

    if check == "gl_keyword_alert":
        return summary_observations(data.gl, selected_account, condition["keywords"])
    tolerance = Decimal(condition["tolerance"])
    if check == "debit_credit_balance":
        return voucher_differences(data.gl, tolerance)
    if check == "period_end_nonzero":
        return post_close_balances(data.tb, selected_account, tolerance, root_only=condition.get("root_only", False))
    if check == "contra_exceeds_gross":
        def gross_account(code):
            return account_matches(condition["gross_account_pattern"], code)

        return contra_observations(data.tb, selected_account, gross_account, condition, tolerance, pairs)
    return direction_observations(data.tb, selected_account, condition, tolerance)


def scan_routed(data, account_codes, *, complete_vouchers=False, period_closed=False,
                currency_exchange_closed=False, account_pairs=None, account_names=None, project_root=ROOT):
    """Screen selected modules with explicitly confirmed ledger and receivable scope."""
    approved_names = confirmed_account_names(account_names)
    route = route_cas(account_codes, account_names=approved_names, project_root=project_root)
    context = {"complete_vouchers": complete_vouchers, "period_closed": period_closed,
               "currency_exchange_closed": currency_exchange_closed}

    def account_matches(pattern, code):
        if pattern["type"] == "names":
            return code in route.named_codes and matches_account(pattern, code, approved_names)
        return ((not route.named_codes or route.named_codes.isdisjoint(account_hierarchy(code)))
                and matches_account(pattern, code))

    try:
        pairs = {account_code(prefix): tuple(sorted(code_list(codes)))
                 for prefix, codes in (account_pairs or {}).items()}
    except (TypeError, ValueError, AttributeError):
        raise AuditError("CAS 备抵科目配对无效。请核对已确认的应收范围。") from None
    rules, findings, deferred = {}, [], []
    for rule_id, rule in load_routed_rules(route).items():
        rules[rule_id] = Rule((), f"{rule['severity']} {rule_id}", rule["cas_ref"], "")
        condition = rule["condition"]
        if condition.get("root_only") and not any(
            account_matches(rule["account_pattern"], row.account_code) for row in data.tb
        ):
            continue
        required = condition.get("requires_context")
        if (required is not None and context[required] is not True
                or rule["check_type"] == "contra_exceeds_gross" and not pairs):
            deferred.append(rule_id)
            continue
        for observation in routed_observations(data, rule, pairs, account_matches):
            if observation is None:
                if rule_id not in deferred:
                    deferred.append(rule_id)
                continue
            position, label, balance, variables = observation
            detail = Template(rule["message_template"]).substitute(variables)
            entry = rule["suggested_entry"]
            gross = Decimal(variables["gross"]) if "gross" in variables else None
            if entry is None:
                findings.append(Finding(rule_id, position, label, balance, detail, gross=gross))
                continue
            debit, credit = (Template(entry[field]).substitute(variables) for field in ("debit", "credit"))
            amount = Decimal(Template(entry["amount"]).substitute(variables)) if entry["amount"] else None
            findings.append(Finding(rule_id, position, label, balance, detail,
                                    f"{entry['type']} 待复核", debit, credit, amount, gross))
    return RoutedScan(route, rules, tuple(findings), tuple(deferred))


def scan_counterparties(groups, ar_account, *, quantify=True):
    """Apply the existing receivable/payable rules to the same groups shown in the schedule."""
    if not groups:
        return {}, ()
    source, _ = load_rules(DEFAULT_RULES, None)
    rules, findings = {}, []
    for group in groups:
        if group.closing >= ZERO:
            continue
        code = group.account_code
        if group.normal_balance == "debit" and code in (*source["AR_CREDIT"].prefixes, ar_account):
            key = "AR_CREDIT"
        elif group.normal_balance == "credit" and code in source["AP_DEBIT"].prefixes:
            key = "AP_DEBIT"
        else:
            key = "COUNTERPARTY_DIRECTION"
        rule = source.get(key) or Rule((), "往来余额方向待核", "",
                                       "核对同一客商的合同、收付款和业务性质后确定重分类及对方科目。")
        rules[key] = rule
        balance = group.closing if group.normal_balance == "debit" else -group.closing
        label = f"{code} / {group.party}" if group.party else code
        detail = ("客商归集后余额与科目正常方向相反。" if group.party
                  else "余额与科目正常方向相反。先取得客商明细。")
        amount = abs(balance) if quantify and group.party and key in source else None
        findings.append(Finding(key, group.source, label, balance, detail, "RJE 待确认", rule.debit, rule.credit, amount))
    return rules, tuple(findings)


def opening_suggestions(differences):
    key = "OPENING_ROLLFORWARD"
    rule = Rule((), "期初滚存差异", "本期未审期初及上年审定期末",
                "核对上期已批准分录和本期入账记录。若确认漏记，按原分录核实对方科目后办理期初滚调。")
    findings = []
    for item in differences:
        counterpart = "待核实对方科目"
        debit, credit = (counterpart, item.label) if item.difference > ZERO else (item.label, counterpart)
        detail = (f"期初差异 {item.difference:.2f}。核查是否属于上期审计调整未入账。"
                  "建议分录以本期未审期初调整到上年审定期末为方向。")
        findings.append(Finding(key, item.position, item.label, None, detail, "期初滚调待复核",
                                debit, credit, abs(item.difference)))
    return ({key: rule} if findings else {}), tuple(findings)


def write_cas_sheet(book, rules, findings, *, deferred_rule_ids=()):
    sheet = book.create_sheet(CAS_SHEET, 1)
    sheet.sheet_view.showGridLines = False
    sheet.sheet_properties.tabColor = "9B682F"
    sheet["A2"] = "准则异常与调整建议"
    sheet["A2"].font = Font(name="Microsoft YaHei", size=15, bold=True, color="18324A")
    sheet["A3"], sheet["B3"] = "命中项", len(findings)
    sheet["A4"] = "按报告期与适用准则复核。分录为草案，期末净额借正贷负。"
    sheet.append(("异常类型", "位置", "科目 / 客商", "期末净额", "应收原值", "异常及复核要点",
                  "准则依据", "建议类型", "借方", "贷方", "待确认金额"))
    body_font = Font(name="Microsoft YaHei", size=10, color="24384A")
    for finding in findings:
        rule = rules[finding.rule_id]
        record = (rule.label, finding.position, finding.subject, finding.balance, finding.gross,
                  f"{finding.detail} {rule.review}", rule.basis, finding.entry_type,
                  finding.debit, finding.credit, finding.draft_amount)
        row = sheet.max_row + 1
        for column, value in enumerate(record, 1):
            cell = sheet.cell(row, column)
            literal(cell, value)
            cell.font = body_font
            cell.alignment = Alignment(vertical="center", wrap_text=isinstance(value, str))
            # GL keyword alerts have no ending balance, so the formatted cell stays blank.
            if column in (4, 5, 11):
                cell.number_format = '#,##0.00;(#,##0.00);0.00'
        sheet.row_dimensions[row].height = 78
    if not findings:
        sheet["A6"] = "已执行的规则未命中异常。" if deferred_rule_ids else "当前扫描规则未命中异常。"
    for column, width in enumerate((27, 20, 32, 16, 16, 55, 40, 19, 19, 19, 16), 1):
        sheet.column_dimensions[get_column_letter(column)].width = width
    header_fill = PatternFill("solid", fgColor="263F56")
    header_font = Font(name="Microsoft YaHei", size=10, bold=True, color="FFFFFF")
    wrapped = Alignment(vertical="center", wrap_text=True)
    for cell in sheet[5]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = wrapped
    sheet.row_dimensions[5].height = 24
    sheet.freeze_panes = "C6"
    sheet.auto_filter.ref = f"A5:K{max(5, 5 + len(findings))}"
    if deferred_rule_ids:
        header = sheet.max_row + 2
        note = "缺少上下文确认、已确认配对或一级汇总行，待数据持有人复核。"
        rows = [("待核规则", "待核说明"), *((rule_id, note) for rule_id in deferred_rule_ids)]
        for row, (label, reason) in enumerate(rows, header):
            sheet.merge_cells(f"A{row}:B{row}")
            sheet.merge_cells(f"C{row}:K{row}")
            for column, value in ((1, label), (3, reason)):
                cell = sheet.cell(row, column)
                literal(cell, value)
                cell.alignment = wrapped
                if row == header:
                    cell.fill = header_fill
                    cell.font = header_font
                else:
                    cell.font = body_font
            sheet.row_dimensions[row].height = 24
