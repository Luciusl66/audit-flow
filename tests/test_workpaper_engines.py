from contextlib import closing, redirect_stderr, redirect_stdout
from dataclasses import replace
from io import StringIO
import os
from pathlib import Path
from shutil import copyfile
import subprocess
import sys
from tempfile import TemporaryDirectory
from types import SimpleNamespace
import unittest
from unittest.mock import patch

from openpyxl import load_workbook
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import PatternFill

from adapters.demo_kingdee import Adapter
from core.accounts import Ledger
from core.blueprint import load_blueprint, load_blueprints
from core.cas_checker import CAS_SHEET, write_cas_sheet
from core.excel_writer import EXCEL_UNAVAILABLE, ExcelWriter, excel_literal, require_excel_writer, run_excel_writer
from core.fill_workpaper import DETAIL, DETAIL_TOTAL, SUMMARY, XlwingsWriter, build, prepare_layout, select_coverage
from core.fillers import make_filler
from core.recalc import Cache, Calculation, scan
from core.reconcile import Gate, write_report
from core.schema import AuditError, GL_SHEET, TB_AMOUNT_FIELDS, ZERO
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


class NativeWorkerBoundaryTest(unittest.TestCase):
    def test_non_windows_refuses_native_writer_before_dependency_discovery(self):
        with (patch("core.excel_writer.os", SimpleNamespace(name="posix")),
              patch("core.excel_writer.find_spec") as find_spec):
            with self.assertRaisesRegex(AuditError, "xlwings 写入需要 Windows"):
                require_excel_writer()
        find_spec.assert_not_called()

    def test_worker_commands_use_private_scratch_and_expected_sheet_order(self):
        for operation, order in (("populate", ()), ("apply_plan", ()), ("attach_reports", ("报告", SUMMARY))):
            with self.subTest(operation=operation), TemporaryDirectory(dir=OUTPUT) as directory:
                scratch = Path(directory)
                source, payload, target = (scratch / name for name in ("stage.xlsx", "payload", "final.xlsx"))

                def fake_subprocess(*args, **kwargs):
                    target.touch()
                    return subprocess.CompletedProcess(args[0], 0)

                with patch("core.excel_writer.subprocess.run", side_effect=fake_subprocess) as worker:
                    run_excel_writer(operation, source, payload, target, 7, order)
                env = {**os.environ, "TMP": directory, "TEMP": directory, "TMPDIR": directory}
                worker.assert_called_once_with(
                    [sys.executable, "-m", "core.excel_writer", operation, str(source), str(payload), str(target), *order],
                    cwd=ROOT, env=env, timeout=7, stdin=subprocess.DEVNULL,
                    stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
                )

    def test_worker_errors_reject_output_and_sanitize_diagnostics(self):
        for operation, stage in (("populate", "填充模板"), ("attach_reports", "插入报告")):
            for result, target_exists, diagnostic in (
                (EXCEL_UNAVAILABLE, False, "无法启动本机 Microsoft Excel"),
                (27, True, f"{stage}失败"),
                (0, False, f"{stage}失败"),
                (subprocess.TimeoutExpired(["private-path"], 1, stderr="private-data"), False, f"{stage}超时"),
                (OSError("private-path"), False, "无法启动本机 Excel 写入进程"),
            ):
                with self.subTest(operation=operation, diagnostic=diagnostic), TemporaryDirectory(dir=OUTPUT) as directory:
                    scratch = Path(directory)
                    target = scratch / "result.xlsx"
                    if target_exists:
                        target.touch()
                    outcome = {"side_effect": result} if isinstance(result, Exception) else {
                        "return_value": subprocess.CompletedProcess([], result)
                    }
                    with patch("core.excel_writer.subprocess.run", **outcome), self.assertRaises(AuditError) as caught:
                        run_excel_writer(operation, scratch / "stage.xlsx", scratch / "payload", target, 1)
                    self.assertIn(diagnostic, str(caught.exception))
                    self.assertNotIn("private-", str(caught.exception))


def attach_synthetic_reports(source, payload, target, sheet_order, edit=None):
    with closing(load_workbook(source)) as book, closing(load_workbook(payload)) as reports:
        for index, name in enumerate(sheet_order):
            if name in reports:
                report = book.create_sheet(name, index)
                for row in reports[name].values:
                    report.append(row)
        if edit is not None:
            edit(book)
        book.save(target)


class NativeWriterCoordinatorTest(unittest.TestCase):
    def setUp(self):
        self.blueprint = load_blueprint()
        self.adapter = Adapter()
        self.data = self.adapter.transform(TB, GL)
        self.coverage = select_coverage(self.data, self.adapter.ar_account)
        self.populate_edit = None
        self.attach_edit = None

    def fake_legacy_worker(self, operation, source, payload, target, timeout, sheet_order=()):
        if operation == "populate":
            book, _ = build(source, self.data, self.coverage, self.adapter.ar_account, blueprint=self.blueprint)
            with closing(book):
                if self.populate_edit is not None:
                    self.populate_edit(book)
                book.save(target)
        else:
            self.assertEqual(operation, "attach_reports")
            attach_synthetic_reports(source, payload, target, sheet_order, self.attach_edit)

    def test_native_population_rejects_a_changed_detail_formula(self):
        self.populate_edit = lambda book: setattr(book[DETAIL]["F6"], "value", "=0")
        with (TemporaryDirectory(dir=OUTPUT) as directory,
              patch("core.fill_workpaper.require_excel_writer"),
              patch("core.fill_workpaper.run_excel_writer", side_effect=self.fake_legacy_worker),
              closing(XlwingsWriter(self.blueprint, Path(directory), 5)) as writer):
            with self.assertRaisesRegex(AuditError, "明细公式与底稿蓝图不一致"):
                writer.populate(self.data, self.coverage, self.adapter.ar_account)

    def test_native_report_attachment_rejects_moved_or_changed_formulas(self):
        for edit in (
            lambda book: book[DETAIL].move_range("F6:F6", rows=10),
            lambda book: setattr(book[DETAIL]["F6"], "value", "=C6"),
        ):
            self.attach_edit = edit
            with (TemporaryDirectory(dir=OUTPUT) as directory,
                  patch("core.fill_workpaper.require_excel_writer"),
                  patch("core.fill_workpaper.run_excel_writer", side_effect=self.fake_legacy_worker),
                  closing(XlwingsWriter(self.blueprint, Path(directory), 5)) as writer):
                writer.populate(self.data, self.coverage, self.adapter.ar_account)
                calculation = Calculation("xlwings", {}, ())
                controls = dict.fromkeys(TB_AMOUNT_FIELDS, (ZERO, ZERO, ZERO))
                write_report(writer.book, tuple(Gate("校验", "") for _ in range(4)), controls, (None, None, None),
                             calculation)
                write_cas_sheet(writer.book, {}, ())
                with self.assertRaisesRegex(AuditError, "公式位置或内容发生变化"):
                    writer.finish(calculation)

    def test_legacy_refuses_to_publish_when_final_cache_no_longer_reconciles(self):
        def fake_caches(path, actual):
            # These controls come from the synthetic fixture and do not evaluate Excel formulas.
            caches = {key: Cache("n", "0") for key in actual}
            caches[SUMMARY, "B9"] = Cache("n", "1950" if path.name == "workpaper.xlsx" else "1950.01")
            caches.update({(DETAIL, f"F{row}"): Cache("n", value)
                           for row, value in enumerate(("975", "675", "300", "1950"), 6)})
            return Calculation("xlwings", caches, scan(actual, caches))

        with (TemporaryDirectory(dir=OUTPUT) as directory, patch("run.store", return_value=Path(directory)),
              patch("core.fill_workpaper.require_excel_writer"),
              patch("core.fill_workpaper.run_excel_writer", side_effect=self.fake_legacy_worker) as worker,
              patch("run.native_calculation", side_effect=fake_caches) as calculation,
              redirect_stdout(StringIO()) as stdout, redirect_stderr(StringIO()) as stderr):
            self.assertEqual(main([*ARGS, "--engine", "xlwings"]), 2)
            self.assertEqual(list(Path(directory).iterdir()), [])
        self.assertEqual([call.args[0] for call in worker.call_args_list], ["populate", "attach_reports"])
        self.assertEqual([call.args[0].name for call in calculation.call_args_list], ["workpaper.xlsx", "excel.xlsx"])
        self.assertEqual(stdout.getvalue().strip(), "FAILED")
        self.assertIn("最终 Excel 工作簿的公式缓存与勾稽报告不一致", stderr.getvalue())

    def test_batch_refuses_to_publish_when_final_native_formula_cache_is_missing(self):
        blueprint = load_blueprints()["1122"]
        tb, gl = MOCK / "mock_tb.xlsx", MOCK / "mock_gl.csv"
        ledger = Ledger(self.adapter.transform(tb, gl))
        with TemporaryDirectory(dir=OUTPUT) as directory:
            output = Path(directory)
            with TemporaryDirectory(dir=OUTPUT) as fake_directory, ExcelWriter(
                "openpyxl", blueprint, Path(fake_directory), 5,
            ) as local_writer:
                layout = make_filler(local_writer, ledger, ("1122",)).populate()
                local_writer.apply()
                caches = {key: Cache("n", "0") for key in local_writer.actual}
                for check in layout.checks:
                    value = check.expected
                    caches[check.position] = Cache("str" if isinstance(value, str) else "n", str(value))

                def fake_batch_worker(operation, source, payload, target, timeout, sheet_order=()):
                    if operation == "apply_plan":
                        copyfile(local_writer.stage, target)
                    else:
                        self.assertEqual(operation, "attach_reports")
                        attach_synthetic_reports(source, payload, target, sheet_order)

                def fake_caches(path, actual):
                    values = caches.copy()
                    if path.name == "excel.xlsx":
                        values.pop(layout.checks[0].position)
                    return Calculation("xlwings", values, scan(actual, values))

                args = ["--adapter", "demo_kingdee", "--tb", str(tb), "--gl", str(gl), "--subjects", "1122",
                        "--engine", "xlwings"]
                with (patch("run.store", return_value=output), patch("core.excel_writer.require_excel_writer"),
                      patch("core.excel_writer.run_excel_writer", side_effect=fake_batch_worker),
                      patch("run.native_calculation", side_effect=fake_caches) as calculation,
                      patch("run.build_master") as master,
                      redirect_stdout(StringIO()) as stdout, redirect_stderr(StringIO()) as stderr):
                    self.assertEqual(main(args), 2)
                    self.assertEqual(list(output.iterdir()), [])
        master.assert_not_called()
        self.assertEqual([call.args[0].name for call in calculation.call_args_list], ["workpaper.xlsx", "excel.xlsx"])
        self.assertEqual(stdout.getvalue().strip(), "FAILED")
        self.assertIn("最终 Excel 工作簿的公式缓存与勾稽报告不一致", stderr.getvalue())


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
