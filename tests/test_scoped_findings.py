from datetime import date
from decimal import Decimal
from types import SimpleNamespace
import unittest

from core.accounts import Ledger
from core.cas_checker import Finding, scan_routed, scan_tb
from core.schema import GL_SHEET, TB_SHEET, StandardData, StandardGL, StandardTB, ZERO
from run import finding_row, scoped_findings


class ScopedFindingsTest(unittest.TestCase):
    def test_parses_standard_and_group_positions(self):
        for position, expected in (
            ("Raw_TB!2", (TB_SHEET, 0)),
            ("Standard_GL!3", (GL_SHEET, 1)),
            ("Raw_TB!4（组内 2 行）", (TB_SHEET, 2)),
        ):
            with self.subTest(position=position):
                self.assertEqual(finding_row(position), expected)

    def test_filters_only_recognized_in_bounds_positions(self):
        def rows(*codes):
            return tuple(SimpleNamespace(account_code=code) for code in codes)

        data = SimpleNamespace(tb=rows("1122.01", "4001"), gl=rows("6001", "1122.02"))
        ledger = SimpleNamespace(owner=lambda code: code.split(".")[0])
        known = ("Raw_TB!2", "Raw_TB!3", "Standard_GL!2", "Standard_GL!3",
                 "Raw_TB!2（组内 2 行）", "Raw_TB!3（组内 2 行）")
        unknown = ("Raw_Prior_TB!3", "凭证级校验", "Raw_TB!1", "Raw_TB!0", "Raw_TB!-1", "Standard_GL!100",
                   "Raw_TB!3尾部未知文本", "Raw_TB!3（组内 不明 行）", "Raw_TB!12345678", None)
        findings = [Finding("MOCK", position, "合成测试", None, "") for position in (*known, *unknown)]
        actual = tuple(item.position for item in scoped_findings(data, ledger, {"1122"}, findings))
        self.assertEqual(actual, (known[0], known[3], known[4], *unknown))

    def test_voucher_finding_stays_with_each_affected_subject_in_either_row_order(self):
        entries = (StandardGL(date(2026, 1, 1), "合成凭证", "", "6001", Decimal("120.00"), ZERO, ""),
                   StandardGL(date(2026, 1, 1), "合成凭证", "", "1122.01", ZERO, Decimal("100.00"), ""))
        tb = tuple(StandardTB(code, "合成科目", 1, "借", ZERO, ZERO, ZERO, ZERO, ZERO, ZERO)
                   for code in ("1122", "6001", "1002"))
        for gl in (entries, entries[::-1]):
            with self.subTest(first=gl[0].account_code):
                data = StandardData(tb, gl)
                result = scan_routed(data, ("1122",), complete_vouchers=True)
                finding, = [item for item in result.findings if item.rule_id == "CAS00_GL_VOUCHER_BALANCE"]
                ledger = Ledger(data)
                for code in ("1122", "6001"):
                    self.assertEqual(tuple(scoped_findings(data, ledger, {code}, (finding,))), (finding,))
                self.assertFalse(tuple(scoped_findings(data, ledger, {"1002"}, (finding,))))

    def test_allowance_finding_stays_with_the_paired_receivable_subject(self):
        rows = (StandardTB("1231", "合成准备", 1, "贷", ZERO, ZERO, ZERO, ZERO, ZERO, Decimal("120.00")),
                StandardTB("1122", "合成应收", 1, "借", ZERO, ZERO, ZERO, ZERO, Decimal("100.00"), ZERO))
        data = StandardData(rows, ())
        ledger = Ledger(data)
        pairs = {"1231": ("1122",)}
        _, general = scan_tb(rows, "1122", account_pairs=pairs)
        routed = scan_routed(data, ("1122",), account_pairs=pairs)
        for rule_id, findings in (
            ("AR_ALLOWANCE", general),
            ("CAS22_TB_AR_ALLOWANCE_EXCEEDS_GROSS", routed.findings),
        ):
            with self.subTest(rule=rule_id):
                finding, = [item for item in findings if item.rule_id == rule_id]
                self.assertTrue(finding.position.startswith("Raw_TB!2"))
                self.assertEqual(tuple(scoped_findings(data, ledger, {"1122"}, (finding,))), (finding,))
