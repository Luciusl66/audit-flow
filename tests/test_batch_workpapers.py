from contextlib import closing, redirect_stderr, redirect_stdout
from dataclasses import replace
from decimal import Decimal
from io import StringIO
import os
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest
from unittest.mock import patch

from openpyxl import load_workbook

from adapters.demo_kingdee import Adapter
from core.accounts import Ledger
from core.blueprint import load_blueprints
from core.cas_checker import CAS_SHEET, opening_suggestions
from core.excel_writer import ExcelWriter
from core.fillers import make_filler
from core.master_tb import MASTER_FILE, MASTER_SHEET, PROPOSALS
from core.recalc import Cache, formulas, read_caches, scan
from core.reconcile import REPORT, cached_money, check_opening, opening_differences
from core.schema import AuditError, PRIOR_TB_SHEET
from run import ROOT, main, select_subjects

MOCK = ROOT / "data" / "mock"
OUTPUT = ROOT / "data" / "output"
TB, GL = MOCK / "mock_tb.xlsx", MOCK / "mock_gl.csv"
PRIOR = MOCK / "mock_prior_tb.xlsx"


def demo_data():
    adapter = Adapter()
    return replace(adapter.transform(TB, GL), prior_tb=adapter.convert_prior_tb(PRIOR))


class BatchWorkpaperTest(unittest.TestCase):
    def test_prior_rolls_up_missing_subtotal_and_fails_for_one_cent_at_auxiliary_grain(self):
        data = demo_data()
        prior = tuple(row for row in data.prior_tb if row.account_code != "1122" or row.aux_name)
        clean = replace(data, prior_tb=prior)
        self.assertEqual(opening_differences(Ledger(clean)), ())
        # Retain the supplied subtotal to isolate the exact auxiliary account.
        altered = tuple(replace(row, prior_audited_debit=row.prior_audited_debit - Decimal("0.01"))
                        if (row.account_code, row.aux_name) == ("1122.01", "模拟甲公司") else row
                        for row in data.prior_tb)
        changed = replace(data, prior_tb=altered)
        differences = opening_differences(Ledger(changed))
        self.assertEqual([(item.code, item.auxiliary, item.difference) for item in differences],
                         [("1122.01", "模拟甲公司", Decimal("0.01"))])
        self.assertEqual(check_opening(changed, differences).status, "FAILED")
        _, (suggestion,) = opening_suggestions(differences)
        self.assertEqual(suggestion.debit, "待核实对方科目")
        self.assertIn("1122.01", suggestion.credit)
        self.assertEqual(suggestion.draft_amount, Decimal("0.01"))
        self.assertEqual(check_opening(replace(data, prior_tb=None), ()).status, "SKIPPED")

    def test_optional_prior_auxiliary_column_and_strict_source_amounts(self):
        with TemporaryDirectory(dir=OUTPUT) as directory:
            path = Path(directory) / "prior.csv"
            header = "科目编码,科目名称,上年审定借方,上年审定贷方\n"
            path.write_text(header + "001,合成科目,10.01,0\n", encoding="utf-8")
            row, = Adapter().convert_prior_tb(path)
            self.assertEqual((row.account_code, row.aux_name, row.prior_audited_debit), ("001", None, Decimal("10.01")))
            path.write_text(header + "001,合成科目,,0\n", encoding="utf-8")
            with self.assertRaisesRegex(AuditError, "第 2 行的 prior_audited_debit"):
                Adapter().convert_prior_tb(path)

    def test_dispatch_supports_prefix_lists_and_does_not_assign_a_root_twice(self):
        ledger = Ledger(demo_data())
        blueprints = load_blueprints()
        selected, owners, skipped = select_subjects("1002,1122,6001", blueprints, ledger)
        self.assertEqual(tuple(selected), ("1002", "1122", "6001"))
        self.assertNotIn("4001", owners)
        self.assertEqual(skipped, [])
        asset = replace(blueprints["1601"], account_codes=("160*",))
        selected, _, _ = select_subjects("1601", {"1601": asset}, ledger)
        self.assertEqual(selected["1601"][1], ("1601", "1602"))
        with self.assertRaisesRegex(AuditError, "范围重叠"):
            select_subjects("ALL", {"1601": asset, "1602": blueprints["1601"]}, ledger)

    def test_gl_independent_sides_fail_when_net_turnover_would_pass(self):
        data = demo_data()
        source = next(row for row in data.gl if row.account_code == "6001.01")
        cent = Decimal("0.01")
        incorrect = replace(source, debit=source.debit + cent, credit=source.credit + cent)
        changed = replace(data, gl=tuple(incorrect if row is source else row for row in data.gl))
        with TemporaryDirectory(dir=OUTPUT) as directory:
            with ExcelWriter("openpyxl", load_blueprints()["6001"], Path(directory), 30) as writer:
                layout = make_filler(writer, Ledger(changed), ("6001",)).populate()
                self.assertEqual([issue[2] for issue in layout.issues], [cent, cent])

    def test_asset_allowance_cap_and_excel_serialization_are_exact(self):
        data = demo_data()
        gross = next(row for row in data.tb if row.account_code == "1601")
        excessive = tuple(replace(row, close_credit=gross.close_debit + Decimal("0.01"))
                          if row.account_code == "1602" else row for row in data.tb)
        with TemporaryDirectory(dir=OUTPUT) as directory:
            with ExcelWriter("openpyxl", load_blueprints()["1601"], Path(directory), 30) as writer:
                layout = make_filler(writer, Ledger(replace(data, tb=excessive)), ("1601", "1602")).populate()
                self.assertIn(("期末备抵余额超过原值", Decimal("0.01")),
                              [(message, difference) for _, message, difference in layout.issues])
        self.assertEqual(cached_money(Cache("n", "26632.799999999999")), Decimal("26632.80"))
        self.assertEqual(cached_money(Cache("n", "307.14999999999998")), Decimal("307.15"))
        self.assertEqual(cached_money(Cache("n", "307.1501")), Decimal("307.1501"))


@unittest.skipUnless(os.environ.get("AUDIT_FLOW_TEST_NATIVE_EXCEL") == "1", "requires local Windows Excel")
class NativeBatchWorkpaperTest(unittest.TestCase):
    def test_all_archetypes_with_native_writer_and_master_adjustment(self):
        self.run_batch(("ALL", "--engine", "xlwings", "--prior-tb", str(PRIOR)), expected_prior=True)

    def test_first_audit_single_subject_skips_only_the_opening_gate(self):
        self.run_batch(("1122",), expected_prior=False)

    def run_batch(self, options, *, expected_prior):
        with TemporaryDirectory(dir=OUTPUT) as directory:
            output = Path(directory)
            args = ["--adapter", "demo_kingdee", "--tb", str(TB), "--gl", str(GL), "--subjects", *options]
            with (patch("run.store", return_value=output), redirect_stdout(StringIO()) as stdout,
                  redirect_stderr(StringIO()) as stderr):
                status = main(args)
            self.assertEqual(status, 0, f"{stderr.getvalue()}\n{stdout.getvalue()}")
            self.assertIn("GATE 5 PASS" if expected_prior else "GATE 5 SKIPPED (未提供上年审定数)", stdout.getvalue())
            blueprints = load_blueprints()
            codes = blueprints if expected_prior else ("1122",)
            for code in codes:
                path = output / f"{code}_workpaper.xlsx"
                with closing(load_workbook(path)) as book:
                    actual = formulas(book)
                    self.assertFalse(scan(actual, read_caches(path, actual)))
                    self.assertIn("SUMIFS", book[blueprints[code].role("summary").name]["C6"].value)
                    self.assertEqual(book[REPORT]["B11"].value,
                                     "PASS" if expected_prior else "SKIPPED (未提供上年审定数)")
                    self.assertEqual(book[PRIOR_TB_SHEET].max_row > 1, expected_prior)
            with closing(load_workbook(output / MASTER_FILE, data_only=True)) as book:
                proposal, = [row for row in book[PROPOSALS].values if row[1] == "RJE 待确认"]
                self.assertEqual(proposal[9:11], (75, "纳入模拟"))
                self.assertIn("PASS", [cell.value for cell in book[MASTER_SHEET]["N"]])
                self.assertIn("本次底稿", [cell.value for cell in book[MASTER_SHEET]["C"]])
            with closing(load_workbook(output / "1122_workpaper.xlsx", data_only=True)) as book:
                self.assertIn("RJE 待确认", [cell.value for cell in book[CAS_SHEET]["H"]])
                self.assertIsNone(book["AJE"]["B6"].value)


if __name__ == "__main__":
    unittest.main()
