"""Rank scoped counterparty balances and expose the five largest concentrations."""

from collections import defaultdict
from dataclasses import dataclass
from decimal import Decimal

from core.fillers.base_filler import BaseFiller, total_formula
from core.schema import TB_SHEET

@dataclass(frozen=True)
class CounterpartyGroup:
    account_code: str
    party: str | None
    source: str
    closing: Decimal
    normal_balance: str


class CounterpartyTB(BaseFiller):
    def fill_detail(self):
        groups = defaultdict(list)
        for code in self.codes:
            for line, row in self.ledger.leaves[code]:
                groups[code, row.aux_name, row.account_code if not row.aux_name else None].append((line, row))
        ranked = sorted(groups.items(), key=lambda item: (-self.account_close(item[1], self.normal_balance),
                                                           item[0][0], item[0][1] or "", item[0][2] or ""))
        section = self.writer.reserve("detail", len(ranked))
        body, scan_groups = [], []
        for row, ((code, party, leaf), group) in enumerate(ranked, section.first):
            key = self.add_group(group)
            movement = self.movement_formulas(section, row, lambda col: self.group_formula(key, col),
                                              self.normal_balance)
            close = self.account_close(group, self.normal_balance)
            scan_groups.append(CounterpartyGroup(code, party, f"{TB_SHEET}!{group[0][0]}", close, self.normal_balance))
            label = party or f"{self.ledger.name(leaf)} [待人工核查]"
            full = section.cell(5, section.total, fixed=True)
            ratio = f'=IF({full}=0,"",{section.cell(5, row)}/{full})' if row < section.first + 5 else None
            body.append((code, label, *movement, ratio))
            self.control(section, 5, row, close, "客商滚转与 TB 期末余额不一致")
        ratios = f"{section.cell(6, section.first)}:{section.cell(6, min(section.total - 1, section.first + 4))}"
        totals = [None, None, *(total_formula(section, i) for i in range(2, 6)), f"=SUM({ratios})"]
        self.fill_section(section, body, totals)
        self.movement_controls(section, self.codes, self.normal_balance)
        self.counterparty_groups = tuple(scan_groups)
        return section.position(5, section.total)
