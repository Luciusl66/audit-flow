from contextlib import closing, redirect_stderr, redirect_stdout
from dataclasses import replace
from decimal import Decimal
from io import StringIO
import os
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest
from unittest.mock import patch

from openpyxl import Workbook, load_workbook

from adapters.demo_kingdee import Adapter
from core.accounts import Ledger
from core.blueprint import load_blueprints
from core.cas_checker import CAS_SHEET, opening_suggestions
from core.excel_writer import ExcelWriter
from core.fillers import make_filler
from core.master_tb import MASTER_FILE, MASTER_SHEET, PROPOSALS, build_master, check_master
from core.recalc import Cache, Calculation, formulas, read_caches, scan
from core.reconcile import REPORT, cached_money, check_opening, opening_differences
from core.schema import AuditError, PRIOR_TB_SHEET, StandardData, StandardPriorTB, StandardTB, ZERO
from run import ROOT, main, publish_batch, select_subjects

MOCK = ROOT / "data" / "mock"
OUTPUT = ROOT / "data" / "output"
TB, GL = MOCK / "mock_tb.xlsx", MOCK / "mock_gl.csv"
PRIOR = MOCK / "mock_prior_tb.xlsx"


def demo_data():
    adapter = Adapter()
    return replace(adapter.transform(TB, GL), prior_tb=adapter.convert_prior_tb(PRIOR))


def synthetic_balance(code, level, amount, *, auxiliary=None, credit=False):
    dr, cr = (ZERO, Decimal(amount)) if credit else (Decimal(amount), ZERO)
    return StandardTB(code, "合成科目", level, "贷" if credit else "借", dr, cr, ZERO, ZERO, dr, cr, auxiliary)


def master_gate(ledger, owners):
    with closing(Workbook()) as book:
        layout = build_master(book, ledger, owners, (), None)
    # Keep all worksheet comparisons green to isolate the independent gate checks.
    caches = {check.position: Cache("str" if isinstance(check.expected, str) else "n", str(check.expected))
              for check in (*layout.checks, *layout.differences)}
    return check_master(layout, Calculation("mock", caches, ()))


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

    def test_opening_does_not_infer_missing_auxiliary_amount(self):
        tb = (synthetic_balance("1122", 1, "10.00"), synthetic_balance("1122", 1, "10.00", auxiliary="合成客商"))
        prior = StandardPriorTB("1122", "合成科目", Decimal("10.00"), ZERO)
        for data in (
            StandardData(tb, (), (prior,)),
            StandardData(tb[:1], (), (prior, replace(prior, aux_name="合成客商"))),
        ):
            with self.subTest(has_current_detail=len(data.tb) > 1):
                difference, = opening_differences(Ledger(data))
                self.assertIsNone(difference.difference)
                self.assertEqual(difference.auxiliary, "合成客商")
                gate = check_opening(data, (difference,))
                self.assertEqual(gate.status, "FAILED")
                self.assertIn("缺少可比", gate.findings[0][1])
                _, (finding,) = opening_suggestions((difference,))
                self.assertEqual((finding.entry_type, finding.draft_amount), ("期初明细待补", None))
                self.assertFalse(finding.debit or finding.credit)

    def test_opening_does_not_propose_the_same_amount_at_parent_and_auxiliary(self):
        tb = (synthetic_balance("1122", 1, "20.00"),
              synthetic_balance("1122.01", 2, "20.00", auxiliary="合成客商"))
        prior = (StandardPriorTB("1122", "合成科目", Decimal("10.00"), ZERO),
                 StandardPriorTB("1122.01", "合成科目", Decimal("10.00"), ZERO, "合成客商"))
        data = StandardData(tb, (), prior)
        differences = opening_differences(Ledger(data))
        self.assertEqual([item.difference for item in differences], [Decimal("10.00")] * 2)
        self.assertEqual(check_opening(data, differences).failed, 2)
        _, (total, detail) = opening_suggestions(differences)
        self.assertIsNone(total.draft_amount)
        self.assertEqual((detail.debit, detail.draft_amount), ("待核实对方科目", Decimal("10.00")))

    def test_master_flags_parent_auxiliary_with_children_in_either_period(self):
        detail = synthetic_balance("1002.01", 2, "20.00")
        equity = synthetic_balance("4001", 1, "20.00", credit=True)
        extra = synthetic_balance("1002", 1, "10.00", auxiliary="合成明细")
        clean_prior = tuple(StandardPriorTB(row.account_code, row.account_name, row.close_debit, row.close_credit)
                            for row in (detail, equity))
        extra_prior = StandardPriorTB(extra.account_code, extra.account_name, extra.close_debit, ZERO, extra.aux_name)
        for tb, prior, sheet in (
            ((extra, detail, equity), clean_prior, "Raw_TB"),
            ((detail, equity), (extra_prior, *clean_prior), PRIOR_TB_SHEET),
        ):
            with self.subTest(sheet=sheet):
                ledger = Ledger(StandardData(tb, (), prior))
                gate = master_gate(ledger, {})
                self.assertEqual(gate.failed, 1)
                self.assertEqual(gate.findings[0][0], f"{sheet}!A2 / 1002")
                self.assertIn("辅助余额", gate.findings[0][1])
                self.assertFalse(tuple(ledger.grain_issues({"4001"})))

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
                self.assertEqual([(position, amount) for position, _, amount in layout.issues],
                                 [("Standard_GL / 6001", cent)] * 2 + [("Standard_GL / 6001.01", cent)] * 2)

    def test_gl_detects_turnover_moved_between_details_with_same_total(self):
        data = demo_data()
        voucher = next(row for row in data.gl if row.account_code == "6001.01" and row.credit > ZERO)
        reallocated = tuple(replace(row, account_code="6001.02") if row is voucher else row for row in data.gl)
        with TemporaryDirectory(dir=OUTPUT) as directory:
            with ExcelWriter("openpyxl", load_blueprints()["6001"], Path(directory), 30) as writer:
                layout = make_filler(writer, Ledger(replace(data, gl=reallocated)), ("6001",)).populate()
                self.assertEqual([(position, amount) for position, _, amount in layout.issues],
                                 [("Standard_GL / 6001.01", -voucher.credit), ("Standard_GL / 6001.02", voucher.credit)])

    def test_gl_only_compares_available_grain_when_all_postings_are_at_root(self):
        data = demo_data()
        summary = tuple(replace(row, account_code="6001") if row.account_code.startswith("6001.") else row
                        for row in data.gl)
        with TemporaryDirectory(dir=OUTPUT) as directory:
            with ExcelWriter("openpyxl", load_blueprints()["6001"], Path(directory), 30) as writer:
                layout = make_filler(writer, Ledger(replace(data, gl=summary)), ("6001",)).populate()
                self.assertEqual(layout.issues, ())

    def test_gl_mixed_grains_do_not_hide_a_detail_mismatch(self):
        data = demo_data()
        voucher = next(row for row in data.gl if row.account_code == "6001.01" and row.credit > ZERO)
        mixed = tuple(replace(row, account_code="6001") if row is voucher else row for row in data.gl)
        with TemporaryDirectory(dir=OUTPUT) as directory:
            with ExcelWriter("openpyxl", load_blueprints()["6001"], Path(directory), 30) as writer:
                layout = make_filler(writer, Ledger(replace(data, gl=mixed)), ("6001",)).populate()
                self.assertEqual(len(layout.issues), 1)
                self.assertIn("GL 同时存在本科目及下级发生额", layout.issues[0][1])

    def test_master_rollforward_detects_offsetting_errors_in_unselected_roots(self):
        data = demo_data()
        hundred = Decimal("100.00")
        changed = tuple(replace(row, close_credit=row.close_credit + hundred) if row.account_code == "6001"
                        else replace(row, close_debit=row.close_debit + hundred) if row.account_code == "6602"
                        else row for row in data.tb)
        gate = master_gate(Ledger(replace(data, tb=changed)), {"1122": "1122"})
        self.assertEqual(gate.failed, 2)
        self.assertEqual([(position.rsplit(" / ", 1)[1], amount) for position, _, amount, _ in gate.findings],
                         [("6001", -hundred), ("6602", hundred)])


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

    def test_batch_becomes_visible_only_when_every_file_is_ready(self):
        with TemporaryDirectory(dir=OUTPUT) as directory:
            output = Path(directory)
            existing = output / MASTER_FILE
            existing.write_bytes(b"previous synthetic output")
            scratch = output / ".run-test"
            scratch.mkdir()
            subject = scratch / "subject.xlsx"
            subject.write_bytes(b"synthetic subject")
            master = scratch / "master.xlsx"
            staged = ((subject, "1122_workpaper.xlsx"), (master, MASTER_FILE))
            with self.assertRaises(FileNotFoundError):
                publish_batch(output, staged, True, scratch)
            self.assertEqual(list(output.glob("run_*")), [])
            self.assertEqual(list(output.glob("*.xlsx")), [existing])
            self.assertEqual(existing.read_bytes(), b"previous synthetic output")

    def test_batch_publication_uses_one_directory_and_consistent_review_suffix(self):
        for passed in (True, False):
            with self.subTest(passed=passed), TemporaryDirectory(dir=OUTPUT) as directory:
                output = Path(directory)
                scratch = output / ".run-test"
                scratch.mkdir()
                first, second = scratch / "subject.xlsx", scratch / "master.xlsx"
                first.write_bytes(b"synthetic subject")
                second.write_bytes(b"synthetic master")
                paths = publish_batch(output, ((first, "1122_workpaper.xlsx"), (second, MASTER_FILE)), passed, scratch)
                batch, = output.glob("run_*")
                self.assertEqual(set(paths), set(batch.iterdir()))
                suffix = ".xlsx" if passed else ".review.xlsx"
                self.assertEqual([path.name for path in paths], [f"1122_workpaper{suffix}", f"00_Master_TB{suffix}"])
                self.assertEqual(list(output.glob("*.xlsx")), [])


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
            batch, = output.glob("run_*")
            blueprints = load_blueprints()
            codes = blueprints if expected_prior else ("1122",)
            for code in codes:
                path = batch / f"{code}_workpaper.xlsx"
                with closing(load_workbook(path)) as book:
                    actual = formulas(book)
                    self.assertFalse(scan(actual, read_caches(path, actual)))
                    self.assertIn("SUMIFS", book[blueprints[code].role("summary").name]["C6"].value)
                    self.assertEqual(book[REPORT]["B11"].value,
                                     "PASS" if expected_prior else "SKIPPED (未提供上年审定数)")
                    self.assertEqual(book[PRIOR_TB_SHEET].max_row > 1, expected_prior)
            with closing(load_workbook(batch / MASTER_FILE, data_only=True)) as book:
                proposal, = [row for row in book[PROPOSALS].values if row[1] == "RJE 待确认"]
                self.assertEqual(proposal[9:11], (75, "纳入模拟"))
                self.assertIn("PASS", [cell.value for cell in book[MASTER_SHEET]["N"]])
                self.assertIn("本次底稿", [cell.value for cell in book[MASTER_SHEET]["C"]])
            with closing(load_workbook(batch / "1122_workpaper.xlsx", data_only=True)) as book:
                self.assertIn("RJE 待确认", [cell.value for cell in book[CAS_SHEET]["H"]])
                self.assertIsNone(book["AJE"]["B6"].value)


if __name__ == "__main__":
    unittest.main()
