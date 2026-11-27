from types import SimpleNamespace
import unittest

from core.cas_checker import Finding
from core.schema import GL_SHEET, TB_SHEET
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
