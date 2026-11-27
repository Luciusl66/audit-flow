"""Terminal-account rollforward for cash, borrowings, and equity."""

from core.fillers.base_filler import BaseFiller


class FlatTB(BaseFiller):
    def fill_detail(self):
        groups = [(row.account_code, row.aux_name or row.account_name, ((line, row),))
                  for code in self.codes for line, row in self.ledger.leaves[code]]
        section = self.movements("detail", self.codes, self.normal_balance, groups)
        return section.position(5, section.total)
