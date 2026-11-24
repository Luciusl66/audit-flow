from decimal import Decimal
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest

from openpyxl import Workbook

from core.cas_checker import CAS_SHEET, scan_tb, write_cas_sheet
from core.schema import AuditError, StandardTB, ZERO

ROOT = Path(__file__).resolve().parents[1]


def row(code, name, *, debit="0.00", credit="0.00", level=1, aux=None):
    return StandardTB(code, name, level, "借", ZERO, ZERO, ZERO, ZERO, Decimal(debit), Decimal(credit), aux)


class CasCheckerTest(unittest.TestCase):
    def test_conditional_reclass_uses_detail_and_correct_debit_credit(self):
        rows = (
            row("1122", "应收账款", credit="10"),
            row("1122.01", "应收明细", level=2, credit="10"),
            row("1122.01", "应收明细", level=2, credit="10", aux="模拟甲"),
            row("2202", "应付账款", debit="20", aux="模拟乙"),
            row("2202.01", "应付明细", level=2, debit="20"),
        )
        _, findings = scan_tb(rows, "1122")
        self.assertEqual([finding.rule_id for finding in findings], ["AR_CREDIT", "AP_DEBIT"])
        ar, ap = findings
        self.assertEqual((ar.debit, ar.credit, ar.draft_amount), ("应收账款", "合同负债", Decimal("10")))
        self.assertIn("模拟甲", ar.subject)
        self.assertEqual((ap.debit, ap.credit), ("预付款项", "应付账款"))
        self.assertIsNone(ap.draft_amount)

    def test_cash_and_fixed_assets_flag_credit_and_negative_column(self):
        rows = (
            row("1002", "银行存款", credit="15"),
            row("1601", "固定资产", debit="1", credit="-5"),
            row("1602", "累计折旧", credit="3"),
            row("6001", "收入", credit="25"),
            row("6602", "管理费用", debit="7"),
            row("6603", "财务费用"),
        )
        _, findings = scan_tb(rows, "1122")
        cash, fixed, income, expense = findings
        self.assertEqual((cash.rule_id, fixed.rule_id), ("CASH_DIRECTION", "PPE_DIRECTION"))
        self.assertIn("贷方栏为负数", fixed.detail)
        self.assertEqual((income.debit, income.credit), ("收入", "本年利润"))
        self.assertEqual((expense.debit, expense.credit), ("本年利润", "管理费用"))
        self.assertIsNone(income.draft_amount)

    def test_allowance_uses_positive_gross_within_confirmed_scope(self):
        rows = (
            row("1122", "应收账款", debit="70"),
            row("1122.01", "应收账款", level=2, debit="100", aux="模拟甲"),
            row("1122.02", "应收账款", level=2, credit="30", aux="模拟乙"),
            row("1221", "其他应收款", debit="200"),
            row("1231", "准备", credit="120"),
            row("1231.01", "应收账款准备", level=2, credit="120"),
        )
        _, findings = scan_tb(rows, "1122", account_pairs={"1231.01": ("1122",)})
        allowance = [finding for finding in findings if finding.rule_id == "AR_ALLOWANCE"]
        self.assertEqual(len(allowance), 1)
        self.assertEqual((allowance[0].balance, allowance[0].gross), (Decimal("-120"), Decimal("100")))
        self.assertIn("20.00", allowance[0].detail)
        self.assertEqual((allowance[0].debit, allowance[0].credit), ("坏账准备", "信用减值损失"))
        self.assertIsNone(allowance[0].draft_amount)

    def test_unallocated_and_missing_base_require_evidence(self):
        rows = (row("1231", "准备", credit="50"),)
        _, unallocated = scan_tb(rows, "1122")
        _, missing_base = scan_tb(rows, "1122", account_pairs={"1231": ("1122",)})
        self.assertIn("缺少已确认配对", unallocated[0].detail)
        self.assertIn("未取得", missing_base[0].detail)
        self.assertIsNone(unallocated[0].gross)
        self.assertIsNone(missing_base[0].gross)
        _, exact = scan_tb((*rows, row("1122", "应收账款", debit="50")),
                           "1122", account_pairs={"1231": ("1122",)})
        self.assertEqual(exact, ())

    def test_same_party_checks_do_not_offset_excess_against_other_customers(self):
        rows = (
            row("1122", "应收账款", debit="10", aux="模拟甲"),
            row("1122", "应收账款", debit="90", aux="模拟乙"),
            row("1231", "坏账准备", credit="20", aux="模拟甲"),
            row("1231", "坏账准备", credit="1", aux="模拟丙"),
        )
        _, findings = scan_tb(rows, "1122", account_pairs={"1231": ("1122",)})
        excess, missing = findings
        self.assertEqual(excess.gross, Decimal("10"))
        self.assertIn("模拟甲", excess.subject)
        self.assertIsNone(missing.gross)
        self.assertIn("未取得", missing.detail)

    def test_invalid_rule_file_fails_explicitly(self):
        with TemporaryDirectory(dir=ROOT / "data" / "output") as directory:
            path = Path(directory) / "rules.yaml"
            path.write_text("rules: {}", encoding="utf-8")
            with self.assertRaises(AuditError):
                scan_tb((), "1122", rules_path=path)

    def test_report_does_not_change_prior_sheets_or_evaluate_source_text(self):
        book = Workbook()
        try:
            book.active.title = "00_自动勾稽与校验报告"
            raw = book.create_sheet("Raw_TB")
            raw["I2"] = 15
            adjustments = book.create_sheet("AJE")
            adjustments["B6"] = 0
            rules, findings = scan_tb((row("2202", "=Unsafe", debit="5", aux="模拟单位"),), "1122")
            write_cas_sheet(book, rules, findings)
            self.assertEqual(book.sheetnames[1], CAS_SHEET)
            self.assertEqual((raw["I2"].value, adjustments["B6"].value), (15, 0))
            self.assertEqual(book[CAS_SHEET]["C6"].data_type, "s")
            self.assertEqual(book[CAS_SHEET]["K6"].value, Decimal("5"))
        finally:
            book.close()
