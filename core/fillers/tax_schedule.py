"""Roll taxes and payroll up to their second-level items."""

from collections import defaultdict

from core.fillers.base_filler import BaseFiller

class TaxSchedule(BaseFiller):
    def fill_detail(self):
        grouped = defaultdict(list)
        for code in self.codes:
            for line, row in self.ledger.leaves[code]:
                key = row.account_code
                while self.ledger.levels[key] > 2 and key in self.ledger.parents:
                    parent = self.ledger.parents[key]
                    if self.ledger.levels[parent] < 2:
                        break
                    key = parent
                grouped[key].append((line, row))
        groups = [(code, self.ledger.name(code), rows) for code, rows in sorted(grouped.items())]
        section = self.movements("detail", self.codes, self.normal_balance, groups)
        return section.position(5, section.total)
