"""Monthly signed turnover backed by Standard_GL and independently tied to TB."""

from collections import defaultdict
from decimal import Decimal

from core.accounts import amount_sides
from core.fillers.base_filler import BaseFiller, PERIOD, rounded, signed, total_formula
from core.schema import GL_SHEET, ZERO
from core.workpaper_data import ACCOUNT_KEY_PREFIX, raw_range, sum_by_key

class GLMonthlyPivot(BaseFiller):
    comparison_column = 4

    def comparison_amount(self):
        return self.balance(self.codes, PERIOD)

    def fill_detail(self):
        codes = {row.account_code for root in self.codes for _, row in self.ledger.leaves[root]}
        monthly = defaultdict(lambda: ZERO)
        debit, credit, years = defaultdict(Decimal), defaultdict(Decimal), set()
        for row in self.data.gl:
            owner = self.ledger.owner(row.account_code)
            if owner not in self.codes:
                continue
            codes.add(row.account_code)
            years.add(row.date.year)
            debit[owner] += row.debit
            credit[owner] += row.credit
            monthly[row.account_code, row.date.month] += signed(row.debit - row.credit, self.normal_balance)
        if len(years) > 1:
            self.issues.append((GL_SHEET, "损益月度透视须使用同一年度序时账", None))
        for code in self.codes:
            source = amount_sides(self.data.tb, self.ledger.current.roots[code], *PERIOD)
            for expected, observed, side in zip(source, (debit[code], credit[code]), ("借方", "贷方"), strict=True):
                if observed != expected:
                    self.issues.append((f"{GL_SHEET} / {code}", f"GL 全年{side}与 TB 本期发生额不一致", observed - expected))
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
