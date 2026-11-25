from contextlib import closing, redirect_stderr, redirect_stdout
from io import StringIO
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest
from unittest.mock import patch

from openpyxl import Workbook, load_workbook

from adapters.demo_kingdee import Adapter
from core.cas_checker import CAS_SHEET, scan_routed, scan_tb, write_cas_sheet
from core.fill_workpaper import DETAIL, SUMMARY
from core.recalc import Cache, Calculation
from core.reconcile import REPORT
from run import ROOT, main

MOCK = ROOT / "data" / "mock"
TB = MOCK / "demo_tb.xlsx"
GL = MOCK / "demo_gl.csv"
VOUCHER_RULE = "CAS00_GL_VOUCHER_BALANCE"
PNL_RULE = "CAS00_TB_POST_CLOSE_PNL"
ALLOWANCE_RULE = "CAS22_TB_AR_ALLOWANCE_EXCEEDS_GROSS"


class CasCliTest(unittest.TestCase):
    def setUp(self):
        temporary = TemporaryDirectory(dir=ROOT / "data" / "output")
        self.addCleanup(temporary.cleanup)
        self.output = Path(temporary.name)
        self.args = ["--adapter", "demo_kingdee", "--tb", str(TB), "--gl", str(GL)]
        # Supply known demo formula caches so the CLI tests can run without an Office backend.
        summary = dict(zip((f"B{row}" for row in range(6, 12)), (950, 1150, 150, 1950, 0, 1950)))
        detail = {
            "C6": 600, "D6": 500, "E6": 125, "F6": 975,
            "C7": 400, "D7": 300, "E7": 25, "F7": 675,
            "C8": -50, "D8": 350, "E8": 0, "F8": 300, "F9": 1950,
        }
        caches = {(name, cell): Cache("n", str(value))
                  for name, values in ((SUMMARY, summary), (DETAIL, detail)) for cell, value in values.items()}
        self.calculation = Calculation("xlwings", caches, ())

    def run_workpaper(self, *options):
        with (patch("run.store", return_value=self.output),
              patch("run.recalculate", return_value=self.calculation),
              redirect_stdout(StringIO()) as stdout, redirect_stderr(StringIO()) as stderr):
            status = main([*self.args, "--out", "workpaper.xlsx", *options])
        self.assertEqual(status, 0, stderr.getvalue())
        for index in range(1, 5):
            self.assertIn(f"GATE {index} PASS", stdout.getvalue())
        return load_workbook(self.output / "workpaper.xlsx")

    def test_default_report_matches_legacy_scan(self):
        adapter = Adapter()
        data = adapter.transform(TB, GL)
        rules, findings = scan_tb(data.tb, adapter.ar_account, account_pairs=adapter.cas_allowance_pairs)
        with closing(self.run_workpaper()) as actual, closing(Workbook()) as expected:
            write_cas_sheet(expected, rules, findings)
            self.assertEqual(actual.sheetnames[:2], [REPORT, CAS_SHEET])
            self.assertEqual(list(actual[CAS_SHEET].values), list(expected[CAS_SHEET].values))
            self.assertEqual(actual[CAS_SHEET]["A6"].value, "损益类科目期末有余额")

    def test_routed_report_includes_findings_with_explicit_context(self):
        names = self.output / "names.json"
        names.write_text('{" 1122 ": " 应收账款 "}', encoding="utf-8-sig")
        with (patch("run.scan_routed", wraps=scan_routed) as scanner,
              closing(self.run_workpaper("--cas-routed", "--complete-vouchers", "--period-closed",
                                         "--currency-closed", "--account-names", str(names))) as book):
            codes = scanner.call_args.args[1]
            self.assertEqual(codes, ("1002", "1122", "1122.01", "1122.02", "4001", "6001"))
            self.assertEqual(scanner.call_args.kwargs["account_names"], {"1122": "应收账款"})
            self.assertIs(scanner.call_args.kwargs["currency_exchange_closed"], True)
            sheet = book[CAS_SHEET]
            self.assertEqual(sheet["B3"].value, 1)
            self.assertIn(PNL_RULE, sheet["A6"].value)
            self.assertEqual(sheet["K6"].value, 1150)
            self.assertEqual(sheet["A8"].value, "待核规则")
            self.assertEqual(sheet["A9"].value, ALLOWANCE_RULE)
            self.assertEqual(sheet.auto_filter.ref, "A5:K6")

    def test_unconfirmed_context_is_reported_as_deferred_only(self):
        with closing(self.run_workpaper("--cas-routed")) as book:
            sheet = book[CAS_SHEET]
            self.assertEqual(sheet["B3"].value, 0)
            self.assertEqual(sheet["A6"].value, "已执行的规则未命中异常。")
            self.assertEqual(sheet["A8"].value, "待核规则")
            self.assertEqual({cell.value for cell in sheet["A"][8:]}, {VOUCHER_RULE, PNL_RULE, ALLOWANCE_RULE})
            self.assertIn("待数据持有人复核", sheet["C9"].value)
            self.assertEqual(sheet.auto_filter.ref, "A5:K5")

    def test_invalid_account_name_file_returns_audit_error_and_no_workpaper(self):
        names = self.output / "names.json"
        for payload, message in (
            ('{"SYNTHETIC_MARKER":', "无法读取 CAS 科目名称映射"),
            ('{"BAD?CODE": "SYNTHETIC_MARKER"}', "CAS 科目名称映射无效"),
            ('{"1122": " "}', "CAS 科目名称映射无效"),
            ('[]', "CAS 科目名称映射须为"),
        ):
            with self.subTest(payload=payload):
                names.write_text(payload, encoding="utf-8")
                with (patch("run.recalculate") as recalculate,
                      redirect_stdout(StringIO()) as stdout, redirect_stderr(StringIO()) as stderr):
                    status = main([*self.args, "--cas-routed", "--account-names", str(names)])
                self.assertEqual(status, 2)
                self.assertEqual(stdout.getvalue().strip(), "FAILED")
                self.assertIn(message, stderr.getvalue())
                self.assertNotIn("SYNTHETIC_MARKER", stderr.getvalue())
                recalculate.assert_not_called()
        self.assertEqual(list(self.output.iterdir()), [names])
