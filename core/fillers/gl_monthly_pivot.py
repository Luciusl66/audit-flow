"""Monthly signed turnover backed by Standard_GL and independently tied to TB."""

from collections import defaultdict
from decimal import Decimal

from core.accounts import amount_sides, ancestors
from core.fillers.base_filler import BaseFiller, PERIOD, rounded, signed, total_formula
from core.schema import GL_SHEET, ZERO
from core.workpaper_data import ACCOUNT_KEY_PREFIX, raw_range, sum_by_key


class GLMonthlyPivot(BaseFiller):
    comparison_column = 4

    def comparison_amount(self):
        return self.balance(self.codes, PERIOD)

    def check_turnover(self, debit, credit, posted, roots):
        controls, parents = self.ledger.current.controls, self.ledger.parents
        debit_by_scope, credit_by_scope = defaultdict(Decimal), defaultdict(Decimal)
        postings_at_scope, postings_below = set(), set()
        for code in debit:
            scope = next(prefix for prefix in (code, *ancestors(code)) if prefix in controls)
            has_posting = code in posted
            if has_posting:
                postings_at_scope.add(scope)
            while scope is not None:
                debit_by_scope[scope] += debit[code]
                credit_by_scope[scope] += credit[code]
                scope = parents.get(scope)
                if has_posting and scope is not None:
                    postings_below.add(scope)

        coarser_only = set()
        for code in sorted(controls, key=lambda item: (len(item), item)):
            if self.ledger.owners[code] not in roots:
                continue
            if code in postings_at_scope and code in postings_below:
                self.issues.append((f"{GL_SHEET} / {code}",
                                    "GL 同时存在本科目及下级发生额，须核对明细发生额的可比口径。", None))
            parent = parents.get(code)
            if parent in postings_at_scope or parent in coarser_only:
                # GL posted at an ancestor cannot be attributed to individual children.
                coarser_only.add(code)
                continue
            source = amount_sides(self.data.tb, controls[code], *PERIOD)
            observed = debit_by_scope[code], credit_by_scope[code]
            for expected, actual, side in zip(source, observed, ("借方", "贷方"), strict=True):
                if actual != expected:
                    issue = f"GL 全年{side}与 TB 本期发生额不一致"
                    self.issues.append((f"{GL_SHEET} / {code}", issue, actual - expected))

    def fill_detail(self):
        roots = frozenset(self.codes)
        codes = {row.account_code for root in self.codes for _, row in self.ledger.leaves[root]}
        monthly = defaultdict(lambda: ZERO)
        debit, credit, years = defaultdict(Decimal), defaultdict(Decimal), set()
        posted = set()
        for row in self.data.gl:
            code = row.account_code
            if self.ledger.owner(code) not in roots:
                continue
            codes.add(code)
            years.add(row.date.year)
            debit[code] += row.debit
            credit[code] += row.credit
            if row.debit != ZERO or row.credit != ZERO:
                posted.add(code)
            monthly[code, row.date.month] += signed(row.debit - row.credit, self.normal_balance)
        if len(years) > 1:
            self.issues.append((GL_SHEET, "损益月度透视须使用同一年度序时账", None))
        self.check_turnover(debit, credit, posted, roots)
        section = self.writer.reserve("detail", len(codes))
        body = []
        gl_count = len(self.data.gl)
        dr, cr = "EF" if self.normal_balance == "debit" else "FE"
        for row, code in enumerate(sorted(codes), section.first):
            key = f'"{ACCOUNT_KEY_PREFIX}"&{section.cell(0, row)}'
            values = []
            for month in range(1, 13):
                criteria = raw_range(GL_SHEET, "I", gl_count), str(month)
                d, c = (sum_by_key(GL_SHEET, column, "H", key, gl_count, *criteria) for column in (dr, cr))
                values.append(rounded(f"{d}-{c}"))
                self.control(section, month + 1, row, monthly[code, month], "月份公式与 GL 发生额不一致")
            annual = rounded(f"SUM({section.cell(2, row)}:{section.cell(13, row)})")
            full = section.cell(14, section.total, fixed=True)
            ratio = f'=IF({full}=0,"",{section.cell(14, row)}/{full})'
            body.append((code, self.ledger.name(code), *values, annual, ratio))
        totals = [None, None, *(total_formula(section, offset) for offset in range(2, 15)),
                  f"=SUM({section.cell(15, section.first)}:{section.cell(15, section.total - 1)})"]
        self.fill_section(section, body, totals)
        self.control(section, 14, section.total, self.comparison_amount(), "GL 透视全年合计与 TB 本期发生额不一致")
        return section.position(14, section.total)
