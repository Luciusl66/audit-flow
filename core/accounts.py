"""Account hierarchy and one accounting grain per balance rollup."""

from collections import defaultdict
from dataclasses import dataclass

from core.schema import PRIOR_AMOUNT_FIELDS, PRIOR_TB_SHEET, TB_AMOUNT_FIELDS, TB_SHEET, ZERO


def leaf_rows(rows):
    """Use terminal accounts and prefer auxiliary detail over its same-code subtotal."""
    by_code = defaultdict(list)
    for index, record in rows:
        by_code[record.account_code].append((index, record))
    levels = {code: min(record.level for _, record in group) for code, group in by_code.items()}
    rollups = {
        code[:end] for code, level in levels.items()
        for end in range(1, len(code)) if levels.get(code[:end], level) < level
    }
    for code in sorted(by_code.keys() - rollups):
        group = by_code[code]
        yield from ([item for item in group if item[1].aux_name is not None] or group)


def ancestors(code):
    return (code[:end] for end in range(len(code) - 1, 0, -1))


def account_parents(levels, prior_codes=()):
    known = levels.keys() | set(prior_codes)
    parents = {}
    for code in sorted(known, key=lambda value: (len(value), value)):
        level = levels.get(code)
        for prefix in ancestors(code):
            if prefix not in known:
                continue
            parent_level = levels.get(prefix)
            if level is not None and parent_level is not None:
                child = parent_level < level
            else:
                child = (code[len(prefix)] in "._/-" and level != 1
                         or level is None and parent_level is not None)
            if child:
                parents[code] = prefix
                break
    return parents


@dataclass(frozen=True)
class BalanceIndex:
    exact: dict
    controls: dict
    roots: dict
    ambiguous_auxiliary: tuple[tuple[str, int], ...]

    @classmethod
    def build(cls, rows, parents, owners, amount_fields):
        exact = defaultdict(list)
        direct = defaultdict(list)
        for line, row in enumerate(rows, 2):
            exact[row.account_code, row.aux_name].append(line)
            direct[row.account_code].append(line)
        codes = set(direct)
        for code in direct:
            parent = parents.get(code)
            while parent is not None and parent not in codes:
                codes.add(parent)
                parent = parents.get(parent)
        children = defaultdict(list)
        controls = {}
        ambiguous = []
        for code in sorted(codes, key=lambda value: (-len(value), value)):
            if children[code]:
                for line in direct[code]:
                    row = rows[line - 2]
                    if row.aux_name is not None and any(getattr(row, field) != ZERO for field in amount_fields):
                        # The export does not say whether these are direct bookings or auxiliary rollups.
                        ambiguous.append((code, line))
                        break
            # A supplied subtotal is authoritative. With no subtotal, roll up
            # independently supplied child balances or same-code auxiliary detail.
            controls[code] = tuple(exact.get((code, None)) or children[code] or direct[code])
            children[parents.get(code)].extend(controls[code])
        return cls(exact, controls, {code: controls[code] for code in sorted(codes) if owners[code] == code},
                   tuple(ambiguous))


class Ledger:
    def __init__(self, data):
        self.data = data
        self.by_code = {}
        for line, row in enumerate(data.tb, 2):
            self.by_code.setdefault(row.account_code, []).append((line, row))
        self.levels = {code: min(row.level for _, row in rows) for code, rows in self.by_code.items()}
        self.parents = account_parents(self.levels)
        self.owners = self.find_owners(self.levels, self.parents)
        self.current = BalanceIndex.build(data.tb, self.parents, self.owners, TB_AMOUNT_FIELDS)
        grouped = defaultdict(list)
        for code, rows in self.by_code.items():
            grouped[self.owners[code]].extend(rows)
        self.members = dict(grouped)
        self.leaves = {code: tuple(leaf_rows(group)) for code, group in self.members.items()}
        prior_codes = {row.account_code for row in data.prior_tb or ()}
        self.prior_parents = account_parents(self.levels, prior_codes)
        self.prior_owners = self.find_owners(self.levels.keys() | prior_codes, self.prior_parents)
        self.prior = BalanceIndex.build(data.prior_tb or (), self.prior_parents, self.prior_owners, PRIOR_AMOUNT_FIELDS)

    def grain_issues(self, roots=None):
        for index, owners, sheet in ((self.current, self.owners, TB_SHEET),
                                     (self.prior, self.prior_owners, PRIOR_TB_SHEET)):
            for code, line in index.ambiguous_auxiliary:
                if roots is None or owners[code] in roots:
                    yield (f"{sheet}!A{line} / {code}",
                           "父科目同时存在非零辅助余额和下级科目。请核实辅助行的汇总范围后重新导入。", None)

    @staticmethod
    def find_owners(codes, parents):
        owners = {}
        for code in sorted(codes, key=lambda value: (len(value), value)):
            owners[code] = owners.get(parents.get(code), code)
        return owners

    def owner(self, code):
        if code in self.owners:
            return self.owners[code]
        for ancestor in ancestors(code):
            if ancestor in self.owners:
                return self.owners[ancestor]
        return self.prior_owners.get(code, code)

    def name(self, code):
        if code in self.by_code:
            return self.by_code[code][0][1].account_name
        lines = self.prior.exact.get((code, None)) or self.prior.controls.get(code)
        if lines:
            row = self.data.prior_tb[lines[0] - 2]
            if row.account_code == code:
                return row.account_name
        return "[待人工核查]"


def amount_sides(rows, indices, debit, credit):
    return (sum((getattr(rows[line - 2], debit) for line in indices), ZERO),
            sum((getattr(rows[line - 2], credit) for line in indices), ZERO))


def net_sides(rows, indices, debit, credit):
    debit, credit = amount_sides(rows, indices, debit, credit)
    return debit - credit
