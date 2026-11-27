"""Common raw tables, auditable summary formulas, and schedule controls."""

from abc import ABC, abstractmethod
from dataclasses import dataclass
from decimal import Decimal

from openpyxl.utils import coordinate_to_tuple, get_column_letter, quote_sheetname

from core.accounts import amount_sides, net_sides
from core.schema import PRIOR_AMOUNT_FIELDS, TB_SHEET, ZERO, AuditError
from core.workpaper_data import (
    ACCOUNT_KEY_PREFIX, AJE, TABLE_FORMATS, prior_lookup, raw_tables, sum_by_key,
)

OPEN = ("open_debit", "open_credit")
PERIOD = ("period_debit", "period_credit")
CLOSE = ("close_debit", "close_credit")


@dataclass(frozen=True)
class BalanceCheck:
    position: tuple[str, str]
    expected: Decimal | str
    message: str


@dataclass(frozen=True)
class SubjectLayout:
    formulas: frozenset
    checks: tuple[BalanceCheck, ...]
    issues: tuple
    comparison: tuple
    counterparty_groups: tuple

    def required_formulas(self):
        return self.formulas


def signed(balance, normal_balance):
    return balance if normal_balance == "debit" else -balance


def rounded(expression):
    return f"=ROUND({expression},2)"


def total_formula(section, offset):
    return rounded(f"SUM({section.cell(offset, section.first)}:{section.cell(offset, section.total - 1)})")


def sheet_ref(sheet, cell):
    if sheet.isascii() and sheet.isidentifier():
        return f"{sheet}!{cell}"
    return f"{quote_sheetname(sheet)}!{cell}"


class BaseFiller(ABC):
    comparison_column = 5  # 本期未审数

    def __init__(self, writer, ledger, codes, normal_balance):
        self.writer = writer
        self.ledger = ledger
        self.data = ledger.data
        self.codes = codes
        self.normal_balance = normal_balance
        self.checks = []
        self.issues = []
        self.detail_keys = {}
        self.counterparty_groups = ()

    def balance(self, codes, fields, normal=None):
        lines = (line for code in codes for line in self.ledger.current.roots[code])
        return signed(net_sides(self.data.tb, tuple(lines), *fields), normal or self.normal_balance)

    def account_close(self, group, normal):
        return signed(sum((row.close_debit - row.close_credit for _, row in group), ZERO), normal)

    def control(self, section, offset, row, expected, message):
        self.checks.append(BalanceCheck(section.position(offset, row), expected, message))

    def movement_formulas(self, section, row, source, normal):
        sign = ("E", "F", "G", "H") if normal == "debit" else ("F", "E", "H", "G")
        opening, opposite, increase, decrease = (source(column) for column in sign)
        c, d, e = (section.cell(offset, row) for offset in (2, 3, 4))
        return rounded(f"{opening}-{opposite}"), rounded(increase), rounded(decrease), rounded(f"{c}+{d}-{e}")

    def group_formula(self, group, column):
        return sum_by_key(TB_SHEET, column, "N", f'"{group}"', len(self.data.tb))

    def add_group(self, rows):
        key = f"grp:{len(self.detail_keys) + 1}"
        self.detail_keys.update((line, key) for line, _ in rows)
        return key

    def fill_section(self, section, rows, totals, *, formula_start=2):
        if section.width != len(totals):
            raise AuditError(f"工作表 {section.spec.name} 的列宽与原型不兼容。")
        self.writer.grid(section.spec, section.first, section.column, rows or [(None,) * section.width],
                         range(formula_start, section.width))
        self.writer.grid(section.spec, section.total, section.column, [("合计",)])
        self.writer.grid(section.spec, section.total, section.column + formula_start,
                         [totals[formula_start:]], range(section.width - formula_start))

    def movements(self, role, codes, normal, groups):
        """Fill a four-column rollforward and reconcile both detail and control totals."""
        section = self.writer.reserve(role, len(groups))
        body = []
        for row, (code, label, group) in enumerate(groups, section.first):
            if len(group) == 1:
                line = group[0][0]
                amounts = self.movement_formulas(section, row, lambda col: f"{TB_SHEET}!{col}{line}", normal)
            else:
                key = self.add_group(group)
                amounts = self.movement_formulas(section, row, lambda col: self.group_formula(key, col), normal)
            body.append((code, label, *amounts))
            self.control(section, 5, row, self.account_close(group, normal), "明细滚转与 TB 期末余额不一致")
        total = [None, None, *(total_formula(section, offset) for offset in range(2, 6))]
        self.fill_section(section, body, total)
        self.movement_controls(section, codes, normal)
        return section

    def movement_controls(self, section, codes, normal):
        self.control(section, 2, section.total, self.balance(codes, OPEN, normal), "明细期初与 TB 控制数不一致")
        lines = tuple(line for code in codes for line in self.ledger.current.roots[code])
        debit, credit = amount_sides(self.data.tb, lines, *PERIOD)
        increase, decrease = (debit, credit) if normal == "debit" else (credit, debit)
        self.control(section, 3, section.total, increase, "明细本期增加与 TB 控制数不一致")
        self.control(section, 4, section.total, decrease, "明细本期减少与 TB 控制数不一致")
        self.control(section, 5, section.total, self.balance(codes, CLOSE, normal), "明细期末与 TB 控制数不一致")

    def aje_scope(self):
        spec = self.writer.blueprint.sheet(AJE)
        anchor = spec.anchors.get("WP_AJE_FIRST")
        if (anchor is None or coordinate_to_tuple(anchor) != (spec.data.min_row, spec.data.min_col)
                or spec.protected_formulas or spec.detail_formulas or spec.data.max_col != spec.data.min_col + 2):
            raise AuditError("AJE 数据区须为科目编码、调整净额及说明三列。")
        if len(self.codes) > spec.data.max_row - spec.data.min_row + 1:
            raise AuditError("AJE 数据区不足以覆盖本底稿的科目。")
        # Seed only the account identifiers. None of the suggested entries are posted here.
        self.writer.grid(spec, spec.data.min_row, spec.data.min_col, [(code,) for code in self.codes])
        codes, amounts = (get_column_letter(spec.data.min_col + offset) for offset in (0, 1))
        first, last = spec.data.min_row, spec.data.max_row
        return f"AJE!${codes}${first}:${codes}${last}", f"AJE!${amounts}${first}:${amounts}${last}"

    def fill_summary(self):
        section = self.writer.reserve("summary", len(self.codes))
        aje_codes, aje_amounts = self.aje_scope()
        body = []
        for row, code in enumerate(self.codes, section.first):
            code_cell = section.cell(0, row)
            prior, prior_debit, prior_credit = prior_lookup(code_cell, len(self.data.prior_tb or ()))
            key = f'"{ACCOUNT_KEY_PREFIX}"&{code_cell}'
            amounts = [sum_by_key(TB_SHEET, column, "M", key, len(self.data.tb)) for column in "EFGHIJ"]
            if self.normal_balance == "credit":
                amounts = [amounts[i] for i in (1, 0, 3, 2, 5, 4)]
                prior_debit, prior_credit = prior_credit, prior_debit
            prior_cell, _, _, current, adjustment, adjusted, change = (section.cell(i, row) for i in range(2, 9))
            body.append((code, self.ledger.name(code), f'=IF({prior}=0,"",ROUND({prior_debit}-{prior_credit},2))',
                         *(rounded(f"{a}-{b}") for a, b in zip(amounts[::2], amounts[1::2], strict=True)),
                         rounded(f"SUMPRODUCT(--EXACT({aje_codes},{code_cell}),{aje_amounts})"),
                         rounded(f"{current}+{adjustment}"), f'=IF({prior_cell}="","",ROUND({adjusted}-{prior_cell},2))',
                         f'=IF(OR({prior_cell}="",{prior_cell}=0),"",{change}/ABS({prior_cell}))'))
            for offset, fields in ((3, OPEN), (4, PERIOD), (5, CLOSE)):
                self.control(section, offset, row, self.balance((code,), fields), "审定表与 TB 控制数不一致")
            lines = self.ledger.prior.roots.get(code, ())
            expected = signed(net_sides(self.data.prior_tb, lines, *PRIOR_AMOUNT_FIELDS), self.normal_balance) if lines else ""
            self.control(section, 2, row, expected, "审定表上年审定数与 Raw_Prior_TB 不一致")
            if len(self.ledger.current.exact.get((code, None), ())) > 1:
                self.issues.append((f"{TB_SHEET} / {code}", "同一总账科目存在多条无辅助核算的汇总行", None))
        row = section.total
        prior = section.cell(2, row)
        adjusted, change = section.cell(7, row), section.cell(8, row)
        prior_values = f"{section.cell(2, section.first)}:{section.cell(2, row - 1)}"
        totals = [None, None, f'=IF(COUNT({prior_values})=0,"",ROUND(SUM({prior_values}),2))',
                  *(total_formula(section, offset) for offset in range(3, 8)),
                  f'=IF({prior}="","",ROUND({adjusted}-{prior},2))',
                  f'=IF(OR({prior}="",{prior}=0),"",{change}/ABS({prior}))']
        self.fill_section(section, body, totals)
        self.control(section, self.comparison_column, row, self.comparison_amount(), "审定表合计与 TB 控制数不一致")
        return section

    def comparison_amount(self):
        return self.balance(self.codes, CLOSE)

    @abstractmethod
    def fill_detail(self):
        """Return the detail control cell comparable with the summary total."""

    def populate(self):
        summary = self.fill_summary()
        detail = self.fill_detail()
        for sheet, fields, rows, count in raw_tables(self.data, self.ledger, self.detail_keys):
            self.writer.table(sheet, fields, rows, count, TABLE_FORMATS)
        issues = (*self.issues, *self.ledger.grain_issues(set(self.codes)))
        return SubjectLayout(frozenset(self.writer.generated), tuple(self.checks), issues,
                             (summary.position(self.comparison_column, summary.total), detail,
                              self.comparison_amount()), self.counterparty_groups)
