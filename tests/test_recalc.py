from contextlib import closing
from pathlib import Path
from subprocess import DEVNULL, CalledProcessError, TimeoutExpired
import sys
from tempfile import TemporaryDirectory
from types import SimpleNamespace
import unittest
from unittest.mock import patch
from xml.etree.ElementTree import fromstring, tostring
from zipfile import ZipFile

from openpyxl import Workbook

from core.recalc import Cache, FormulaMismatch, read_caches, recalculate, sheet_parts, tag
from core.reconcile import REPORT, Gate, write_report
from core.schema import AuditError, TB_AMOUNT_FIELDS, ZERO
from run import native_calculation

ROOT = Path(__file__).resolve().parents[1]
EXPECTED = {("SUMMARY", "B9"): "=1+1"}
CACHES = {("SUMMARY", "B9"): Cache("n", "2")}


def saved_formulas(path, cells):
    """Write a synthetic workbook with the requested formula XML and cached values."""
    original = path.with_name("unevaluated.xlsx")
    with closing(Workbook()) as book:
        sheet = book.active
        sheet.title = "SUMMARY"
        book.create_sheet("明细")
        for address in cells:
            sheet[address] = "=0"
        book.save(original)
    with ZipFile(original) as source, ZipFile(path, "w") as target:
        part = sheet_parts(source)["SUMMARY"]
        for item in source.infolist():
            payload = source.read(item.filename)
            if item.filename == part:
                root = fromstring(payload)
                for cell in root.iter(tag("c")):
                    attributes, text, cache = cells[cell.get("r")]
                    formula = cell.find(tag("f"))
                    formula.attrib.update(attributes)
                    formula.text = text
                    cell.find(tag("v")).text = cache
                payload = tostring(root)
            target.writestr(item, payload)


class CacheProvenanceTest(unittest.TestCase):
    def test_accepts_shared_formula_translation_and_equivalent_sheet_quotes(self):
        expected = {("SUMMARY", "B9"): "=sum('明细'!A1,$A$1)", ("SUMMARY", "B10"): "=SUM('明细'!A2,$A$1)"}
        cells = {"B9": ({"t": "shared", "si": "0", "ref": "B9:B10"}, "SUM(明细!A1,$A$1)", "2"),
                 "B10": ({"t": "shared", "si": "0"}, None, "3")}
        with TemporaryDirectory(dir=ROOT / "data" / "output") as temporary:
            path = Path(temporary) / "saved.xlsx"
            saved_formulas(path, cells)
            self.assertEqual(read_caches(path, expected),
                             {("SUMMARY", "B9"): Cache("n", "2"), ("SUMMARY", "B10"): Cache("n", "3")})
            wrong_follower = {**expected, ("SUMMARY", "B10"): "=SUM('明细'!A1,$A$1)"}
            with self.assertRaises(FormulaMismatch):
                read_caches(path, wrong_follower)

    def test_rejects_changed_formula_even_when_cached_number_matches(self):
        with TemporaryDirectory(dir=ROOT / "data" / "output") as temporary:
            source = Path(temporary) / "stage.xlsx"

            def run(*args, cwd, **kwargs):
                saved_formulas(cwd / source.name, {"B9": ({}, "1*2", "2")})

            with (patch("core.recalc.soffice_path", return_value="soffice"),
                  patch("core.recalc.subprocess.run", side_effect=run)):
                calculation = recalculate(source, EXPECTED, "libreoffice", 1)
            self.assertEqual(calculation.backend, "unavailable")
            self.assertEqual(calculation.caches, {})
            self.assertEqual(calculation.note, "libreoffice: 重算文件中的公式与待重算公式不一致")
            with self.assertRaisesRegex(AuditError, "Excel 缓存对应公式与预期公式不一致"):
                native_calculation(source.parent / "libreoffice" / source.name, EXPECTED)

    def test_formula_comparison_preserves_quoted_string_contents(self):
        with TemporaryDirectory(dir=ROOT / "data" / "output") as temporary:
            path = Path(temporary) / "saved.xlsx"
            saved_formulas(path, {"B9": ({}, 'EXACT(" A ","a")', "0")})
            with self.assertRaises(FormulaMismatch):
                read_caches(path, {("SUMMARY", "B9"): '=EXACT(" A ","A")'})


class RecalculationDiagnosticsTest(unittest.TestCase):
    def test_xlwings_recalculation_worker_failure_uses_private_directory(self):
        failure = CalledProcessError(19, ["worker"], stderr="private-data")
        with TemporaryDirectory(dir=ROOT / "data" / "output") as temporary:
            source = Path(temporary) / "stage.xlsx"
            scratch = source.parent / "xlwings"
            with (
                patch("core.recalc.os", SimpleNamespace(environ={"EXISTING": "one"}, name="nt")),
                patch("core.recalc.find_spec", return_value=object()),
                patch("core.recalc.subprocess.run", side_effect=failure) as worker,
            ):
                calculation = recalculate(source, EXPECTED, "xlwings", 9)
            worker.assert_called_once_with(
                [sys.executable, str(ROOT / "core" / "recalc.py"), "--excel-worker", str(source),
                 str(scratch / source.name)],
                cwd=scratch, env={"EXISTING": "one", **dict.fromkeys(("TMP", "TEMP", "TMPDIR"), str(scratch))},
                stdin=DEVNULL, stdout=DEVNULL, stderr=DEVNULL, check=True, timeout=9,
            )
        self.assertEqual(calculation.note, "xlwings: 引擎执行失败")
        self.assertEqual(calculation.backend, "unavailable")
        self.assertEqual(calculation.caches, {})

    def test_fallback_preserves_sanitized_failure_in_report(self):
        failure = TimeoutExpired(["xlwings", "private-path.xlsx"], 1, stderr="private-amount")
        with TemporaryDirectory(dir=ROOT / "data" / "output") as temporary:
            source = Path(temporary) / "stage.xlsx"
            with (
                patch("core.recalc.os", SimpleNamespace(name="nt", environ={})),
                patch("core.recalc.find_spec", return_value=object()),
                patch("core.recalc.soffice_path", return_value="soffice"),
                patch("core.recalc.subprocess.run", side_effect=[failure, None]),
                patch("core.recalc.read_caches", return_value=CACHES),
            ):
                calculation = recalculate(source, EXPECTED, "auto", 1)
        self.assertEqual(calculation.backend, "libreoffice")
        self.assertEqual(calculation.caches, CACHES)
        self.assertEqual(calculation.issues, ())
        self.assertEqual(calculation.note, "xlwings: 引擎运行超时")
        book = Workbook()
        try:
            controls = dict.fromkeys(TB_AMOUNT_FIELDS, (ZERO, ZERO, ZERO))
            gates = tuple(Gate("校验", "") for _ in range(4))
            write_report(book, gates, controls, (None, None, None), calculation)
            self.assertEqual(book[REPORT]["E4"].value, calculation.note)
        finally:
            book.close()

    def test_unavailable_records_each_missing_backend(self):
        with TemporaryDirectory(dir=ROOT / "data" / "output") as temporary:
            source = Path(temporary) / "stage.xlsx"
            with (
                patch("core.recalc.os", SimpleNamespace(name="nt", environ={})),
                patch("core.recalc.find_spec", return_value=None),
                patch("core.recalc.soffice_path", return_value=None),
                patch("core.recalc.subprocess.run") as run,
            ):
                calculation = recalculate(source, EXPECTED, "auto", 1)
        run.assert_not_called()
        self.assertEqual(calculation.backend, "unavailable")
        self.assertEqual(calculation.caches, {})
        self.assertEqual(calculation.note, "xlwings: 未安装 xlwings；libreoffice: 未找到 soffice")
        self.assertEqual(len(calculation.issues), 1)

    def test_missing_and_damaged_outputs_have_distinct_diagnostics(self):
        for artifact, reason in (
            (None, "未生成重算工作簿"),
            (b"invalid workbook", "重算工作簿损坏或格式无效"),
        ):
            with self.subTest(reason=reason), TemporaryDirectory(dir=ROOT / "data" / "output") as temporary:
                source = Path(temporary) / "stage.xlsx"

                def run(*args, cwd, **kwargs):
                    if artifact is not None:
                        (cwd / source.name).write_bytes(artifact)

                with (
                    patch("core.recalc.soffice_path", return_value="soffice"),
                    patch("core.recalc.subprocess.run", side_effect=run),
                ):
                    calculation = recalculate(source, EXPECTED, "libreoffice", 1)
                self.assertEqual(calculation.backend, "unavailable")
                self.assertEqual(calculation.note, f"libreoffice: {reason}")
