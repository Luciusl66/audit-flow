from contextlib import closing, redirect_stderr, redirect_stdout
from dataclasses import replace
from io import StringIO
import os
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest
from unittest.mock import patch

from openpyxl import load_workbook
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import PatternFill

from adapters.demo_kingdee import Adapter
from core.blueprint import load_blueprint
from core.cas_checker import CAS_SHEET
from core.excel_writer import excel_literal
from core.fill_workpaper import DETAIL, DETAIL_TOTAL, SUMMARY, build, prepare_layout, select_coverage
from core.recalc import Calculation
from core.schema import AuditError, GL_SHEET
from run import ROOT, main

OUTPUT = ROOT / "data" / "output"
MOCK = ROOT / "data" / "mock"
TB, GL = MOCK / "demo_tb.xlsx", MOCK / "demo_gl.csv"
ARGS = ("--adapter", "demo_kingdee", "--tb", str(TB), "--gl", str(GL), "--out", "workpaper.xlsx")


class BlueprintTest(unittest.TestCase):
    def setUp(self):
        self.blueprint = load_blueprint()
        self.account = Adapter.ar_account
        self.data = Adapter().transform(TB, GL)
        self.coverage = select_coverage(self.data, self.account)

    def test_template_drift_stops_before_writing(self):
        for change, diagnostic in (
            (lambda book: book.create_sheet("意外工作表"), "工作表清单或顺序"),
            (lambda book: book.move_sheet(DETAIL, offset=-1), "工作表清单或顺序"),
            (lambda book: setattr(book[DETAIL]["A5"], "value", "漂移表头"), "固定表头"),
            (lambda book: setattr(book.defined_names[DETAIL_TOTAL], "attr_text", f"'{DETAIL}'!$A$8"), "模板锚点"),
            (lambda book: setattr(book[SUMMARY]["B9"], "value", 0), "保护区域缺少公式"),
        ):
            with self.subTest(diagnostic=diagnostic), closing(load_workbook(self.blueprint.template)) as book:
                change(book)
                with self.assertRaisesRegex(AuditError, diagnostic):
                    prepare_layout(book, self.blueprint, self.data, self.coverage)

    def test_openpyxl_rejects_merged_total_when_rows_must_expand(self):
        with TemporaryDirectory(dir=OUTPUT) as directory:
            template = Path(directory) / "merged.xlsx"
            with closing(load_workbook(self.blueprint.template)) as book:
                book[DETAIL].merge_cells("A7:B7")
                book.save(template)
            with self.assertRaisesRegex(AuditError, "--engine xlwings"):
                build(template, self.data, self.coverage, self.account, blueprint=self.blueprint)

    def test_missing_native_engine_reports_error_without_workpaper(self):
        with TemporaryDirectory(dir=OUTPUT) as directory:
            self.output = Path(directory)
            with patch("core.excel_writer.find_spec", return_value=None):
                self.assert_failed_generation("xlwings 写入需要", "--engine", "xlwings")

    def test_native_text_that_excel_would_truncate_is_rejected(self):
        for value in ("x" * 32_767, "𠀀" * 16_384):
            with self.subTest(length=len(value)), self.assertRaisesRegex(AuditError, "--engine openpyxl"):
                excel_literal(value)

    def test_generated_report_cannot_add_sheets_or_write_outside_blueprint(self):
        from core.reconcile import REPORT, write_report

        def out_of_scope(book, *args):
            write_report(book, *args)
            book[REPORT]["F24"] = "超范围"

        def extra_sheet(book, *args):
            write_report(book, *args)
            book.create_sheet("意外工作表")

        for renderer, message in ((extra_sheet, "工作表清单或顺序"), (out_of_scope, "报告写入超出")):
            with self.subTest(renderer=renderer.__name__), TemporaryDirectory(dir=OUTPUT) as directory:
                self.output = Path(directory)
                with (patch("run.write_report", side_effect=renderer),
                      patch("run.recalculate", return_value=Calculation("unavailable", {}, ()))):
                    self.assert_failed_generation(message)

    def assert_failed_generation(self, diagnostic, *options):
        with (patch("run.store", return_value=self.output),
              redirect_stdout(StringIO()) as stdout, redirect_stderr(StringIO()) as stderr):
            status = main([*ARGS, *options])
        self.assertEqual(status, 2)
        self.assertEqual(stdout.getvalue().strip(), "FAILED")
        self.assertIn(diagnostic, stderr.getvalue())
        self.assertEqual(list(self.output.iterdir()), [])


@unittest.skipUnless(os.environ.get("AUDIT_FLOW_TEST_NATIVE_EXCEL") == "1", "requires local Windows Excel")
class NativeExcelTest(unittest.TestCase):
    def test_native_writer_preserves_layout_and_literal_source_text(self):
        long_text = "=" + "x" * 32_765

        class SyntheticAdapter(Adapter):
            def transform(self, tb_path, gl_path):
                data = super().transform(tb_path, gl_path)
                sample = (
                    replace(data.gl[0], summary="=1+1", voucher_no="0001"),
                    replace(data.gl[1], summary="'合成单引号"),
                    replace(data.gl[2], summary=long_text),
                    *data.gl[3:],
                )
                return replace(data, gl=sample)

        blueprint = load_blueprint()
        with TemporaryDirectory(dir=OUTPUT) as directory:
            output = Path(directory)
            template = output / "synthetic_rich_template.xlsx"
            with closing(load_workbook(blueprint.template)) as book:
                book[SUMMARY].merge_cells("A2:B2")
                original_formulas = blueprint.frozen_formulas(book)
                detail = book[DETAIL]
                detail.merge_cells("A7:B7")
                detail.merge_cells("A10:B10")
                detail["A10"], detail["D10"] = "合成模板说明", "=F7"
                detail.conditional_formatting.add(
                    "F6:F7", CellIsRule(operator="lessThan", formula=["0"],
                                         fill=PatternFill("solid", fgColor="FF0000"))
                )
                book.save(template)
            with (patch("run.load_blueprint", return_value=replace(blueprint, template=template)),
                  patch("run.load_adapter", return_value=SyntheticAdapter()),
                  patch("run.store", return_value=output),
                  redirect_stdout(StringIO()) as stdout, redirect_stderr(StringIO()) as stderr):
                status = main([*ARGS, "--engine", "xlwings", "--cas-routed"])
            self.assertEqual(status, 0, stderr.getvalue())
            for index in range(1, 5):
                self.assertIn(f"GATE {index} PASS", stdout.getvalue())
            with closing(load_workbook(output / "workpaper.xlsx")) as book:
                self.assertEqual(book.sheetnames, [spec.name for spec in blueprint.sheets])
                self.assertIn("待核规则", [cell.value for cell in book[CAS_SHEET]["A"]])
                self.assertEqual(blueprint.frozen_formulas(book), original_formulas)
                self.assertIn("A2:B2", book[SUMMARY].merged_cells)
                detail = book[DETAIL]
                self.assertIn("A9:B9", detail.merged_cells)
                self.assertIn("A12:B12", detail.merged_cells)
                self.assertEqual((detail["A9"].value, detail["A12"].value), ("合计", "合成模板说明"))
                self.assertEqual(detail["D12"].value, "=F9")
                self.assertEqual(detail.conditional_formatting["F6:F9"][0].formula, ["0"])
                self.assertEqual(book[GL_SHEET]["B2"].value, "0001")
                self.assertEqual(book[GL_SHEET]["C2"].data_type, "s")
                self.assertEqual(book[GL_SHEET]["C2"].value, "=1+1")
                self.assertEqual(book[GL_SHEET]["C3"].value, "'合成单引号")
                self.assertIsNone(book[GL_SHEET]["G3"].value)
                self.assertEqual(book[GL_SHEET]["C4"].data_type, "s")
                self.assertEqual(len(book[GL_SHEET]["C4"].value), len(long_text))
                self.assertEqual(book[GL_SHEET]["C4"].value, long_text)
