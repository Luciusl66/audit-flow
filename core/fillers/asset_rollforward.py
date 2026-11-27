"""Link original cost, contra balances, and net carrying amounts."""

from core.blueprint import matches_code
from core.fillers.base_filler import BaseFiller, CLOSE, OPEN, rounded, sheet_ref

class AssetRollforward(BaseFiller):
    def fill_detail(self):
        pairs = self.writer.blueprint.contra_account_codes
        gross = tuple(code for code in self.codes if not matches_code(code, pairs))
        contra = tuple(code for code in self.codes if matches_code(code, pairs))
        schedules = []
        for role, codes, normal in (("gross", gross, "debit"), ("contra", contra, "credit")):
            groups = [(row.account_code, row.aux_name or row.account_name, ((line, row),))
                      for code in codes for line, row in self.ledger.leaves[code]]
            schedules.append(self.movements(role, codes, normal, groups))
        if not gross:
            self.issues.append(("原值变动表", "未取得配对的资产原值科目", None))
        net = self.writer.reserve("net", 1)
        cost, allowance = schedules
        body = ["、".join(self.codes), "资产净值"]
        totals = [None, None]
        for offset in range(2, 6):
            cost_ref, allowance_ref = (sheet_ref(section.spec.name, section.cell(offset, section.total))
                                       for section in schedules)
            body.append(rounded(f"{cost_ref}-{allowance_ref}"))
            totals.append(rounded(f"SUM({net.cell(offset, net.first)}:{net.cell(offset, net.total - 1)})"))
        exceeded = False
        for offset, fields, label in ((2, OPEN, "期初"), (5, CLOSE, "期末")):
            excess = self.balance(contra, fields, "credit") - self.balance(gross, fields, "debit")
            if excess > 0:
                self.issues.append((f"{allowance.spec.name}!{allowance.cell(offset, allowance.total)}",
                                    f"{label}备抵余额超过原值", excess))
                exceeded = True
        opening, closing = (net.cell(i, net.first) for i in (2, 5))
        body.append(f'=IF(OR({opening}<0,{closing}<0),"FAILED","PASS")')
        totals.append(f"={net.cell(6, net.first)}")
        self.fill_section(net, [tuple(body)], totals)
        self.control(net, 6, net.first, "FAILED" if exceeded else "PASS", "备抵上限公式与 TB 比较结果不一致")
        self.control(net, 5, net.total, self.comparison_amount(), "资产净值与 TB 原值及备抵净额不一致")
        return net.position(5, net.total)
