from dataclasses import replace
from decimal import Decimal
import json
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest
from unittest.mock import patch
from zipfile import ZipFile

from openpyxl import Workbook, load_workbook

from adapters.base_adapter import SourceCell, mapped_rows, source_rows, source_totals
from adapters.demo_kingdee import Adapter, GL_COLUMNS, TB_COLUMNS
from core.fill_workpaper import DETAIL, DETAIL_TOTAL, SUMMARY, build, select_coverage
from core.masker import profile
from core.recalc import Cache, Calculation, formulas, scan, write_caches
from core.reconcile import REPORT, Gate, check_conservation, check_coverage, check_cross, check_excel, write_report
from core.schema import AuditError, GL_SHEET, TB_SHEET, account_code, validate
from run import generate, parse_args, store
from scripts.generate_demo import raw_sum

ROOT = Path(__file__).resolve().parents[1]
MOCK = ROOT / "data" / "mock"
TEMPLATE = ROOT / "templates" / "ar_template.xlsx"


class WorkpaperTest(unittest.TestCase):
    def setUp(self):
        self.adapter = Adapter()
        self.tb_path = MOCK / "demo_tb.xlsx"
        self.gl_path = MOCK / "demo_gl.csv"
        self.data = self.adapter.transform(self.tb_path, self.gl_path)
        validate(self.data)
        self.coverage = select_coverage(self.data, self.adapter.ar_account)

    def test_adapter_headers_are_exactly_present_in_each_profile(self):
        for path, mapping in ((self.tb_path, TB_COLUMNS), (self.gl_path, GL_COLUMNS)):
            reported = {entry["header"] for entry in profile(path)["columns"]}
            self.assertTrue(set(mapping.values()) <= reported)

    def test_mapping_error_explains_exact_header_match(self):
        rows = iter([(SourceCell("科目名称 "),)])
        with self.assertRaisesRegex(AuditError, "逐字一致（含空格、全角/半角及标点）"):
            next(mapped_rows(rows, {"account_name": "科目名称"}, "TB"))

    def test_output_store_distinguishes_missing_and_redirected_directories(self):
        with TemporaryDirectory(dir=ROOT / "data" / "output") as temporary:
            root = Path(temporary)
            target = root / "data" / "output"
            with patch("run.ROOT", root):
                with self.assertRaisesRegex(AuditError, "项目数据目录缺失。"):
                    store("output")
                target.mkdir(parents=True)
                self.assertEqual(store("output"), target)
                with patch.object(type(target), "resolve", return_value=root / "redirected"):
                    with self.assertRaisesRegex(AuditError, "出于数据安全.*符号链接或重定向"):
                        store("output")

    def test_masking_keeps_shape_and_approved_direction_without_raw_text(self):
        with TemporaryDirectory(dir=ROOT / "data" / "output") as tmp:
            path = Path(tmp) / "sample.csv"
            path.write_text("摘要,往来单位,借方金额,余额方向,Private Header\n"
                            'Secret Note,Secret Name,"(8,291.50)",贷,Secret Info\n'
                            "Other Note,Secret Name,+81.7,借,Other Secret\n"
                            + "Other Note,Another Name,,Other Secret,More Secret\n" * 8
                            + "Eleventh Note,Another Name,9.33,平,More Secret\n", encoding="utf-8")
            result = profile(path)
            approved = profile(path, approved_headers={"Private Header": "summary"})
        features = result["columns"][2]["features"]
        self.assertEqual(result["rows_sampled"], 10)
        self.assertEqual(result["masked_sample"][0][:4], [None, "客商_001", "(1,000.00)", "贷"])
        self.assertEqual(result["masked_sample"][1][:4], [None, "客商_001", "+100.0", "借"])
        self.assertEqual((features["thousands_count"], features["parenthesized_negative_count"],
                          features["blank_count"]), (1, 1, 8))
        self.assertEqual(result["columns"][3]["enum_values"], ["借", "贷"])
        self.assertIsNone(result["columns"][4]["header"])
        output = json.dumps(result, ensure_ascii=False)
        for secret in ("Private Header", "Secret", "8,291", "81.7", "Eleventh Note"):
            self.assertNotIn(secret, output)
        self.assertEqual(approved["columns"][4]["header"], "Private Header")
        self.assertTrue(all(row[4] is None for row in approved["masked_sample"]))

    def test_cached_ods_formula_is_available_only_to_masked_probe(self):
        content = """<?xml version="1.0" encoding="UTF-8"?>
        <office:document-content
            xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
            xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
            xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" office:version="1.2">
          <office:body><office:spreadsheet><table:table table:name="合成">
            <table:table-row><table:table-cell office:value-type="string">
              <text:p>借方金额</text:p>
            </table:table-cell></table:table-row>
            <table:table-row><table:table-cell table:formula="of:=1+1" office:value-type="float" office:value="9">
              <text:p>9</text:p>
            </table:table-cell></table:table-row>
          </table:table></office:spreadsheet></office:body>
        </office:document-content>"""
        manifest = """<?xml version="1.0" encoding="UTF-8"?>
        <manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0">
          <manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.spreadsheet"/>
          <manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/>
        </manifest:manifest>"""
        with TemporaryDirectory(dir=ROOT / "data" / "output") as tmp:
            path = Path(tmp) / "synthetic.ods"
            with ZipFile(path, "w") as archive:
                archive.writestr("mimetype", "application/vnd.oasis.opendocument.spreadsheet")
                archive.writestr("content.xml", content)
                archive.writestr("META-INF/manifest.xml", manifest)
            result = profile(path)
            self.assertEqual(result["columns"][0]["types"], {"number_cell": 1})
            self.assertEqual(Decimal(result["masked_sample"][0][0]), Decimal("100"))
            with self.assertRaisesRegex(AuditError, "只能用于脱敏探针"), source_rows(path):
                self.fail("底稿读取不应取得只有缓存的单元格")

    def test_all_audit_inputs_reject_cached_only_formats_before_loading_adapter(self):
        with TemporaryDirectory(dir=ROOT / "data" / "output") as tmp:
            directory = Path(tmp)
            verified = [directory / f"{name}.xlsx" for name in ("tb", "gl", "prior")]
            for index, suffix in enumerate((".xls", ".xlsb", ".ods")):
                with self.subTest(input_index=index, suffix=suffix):
                    paths = list(verified)
                    paths[index] = paths[index].with_suffix(suffix)
                    for path in paths:
                        path.touch()
                    args = parse_args(["--adapter", "demo_kingdee", "--tb", str(paths[0]),
                                       "--gl", str(paths[1]), "--prior-tb", str(paths[2])])
                    with (patch("run.safe_path", side_effect=[(path, "mock") for path in paths]),
                          patch("run.load_adapter") as adapter, self.assertRaisesRegex(AuditError, "只能用于脱敏探针")):
                        generate(args)
                    adapter.assert_not_called()

    def test_conservation_rejects_a_one_cent_change_or_discarded_row(self):
        source = source_totals(self.tb_path, self.adapter.tb_columns, self.adapter.tb_source, False)
        gate, _ = check_conservation(self.data, *source)
        self.assertEqual(gate.status, "PASS")
        original = self.data.tb[0]
        altered = replace(original, close_debit=original.close_debit + Decimal("0.01"))
        changed = replace(self.data, tb=(altered, *self.data.tb[1:]))
        gate, controls = check_conservation(changed, *source)
        self.assertEqual(gate.status, "FAILED")
        self.assertEqual(controls["close_debit"][2], Decimal("0.01"))
        truncated = replace(self.data, tb=self.data.tb[1:])
        gate, _ = check_conservation(truncated, *source)
        self.assertTrue(any("行数差：-1" in issue[1] for issue in gate.findings))

    def test_report_lists_row_count_as_integer_text_without_money_format(self):
        count, totals = source_totals(self.tb_path, self.adapter.tb_columns, self.adapter.tb_source, False)
        gate, controls = check_conservation(self.data, count + 1, totals)
        book = Workbook()
        try:
            gates = (gate, *(Gate("其他校验", "") for _ in range(3)))
            write_report(book, gates, controls, (None, None, None), Calculation("xlwings", {}, ()))
            sheet = book[REPORT]
            message = next(cell for cell in sheet["E"] if "行数差" in str(cell.value))
            self.assertEqual(message.value, "清洗前后数据行数不一致（行数差：-1）")
            self.assertIsNone(sheet[f"D{message.row}"].value)
            self.assertEqual(sheet[f"D{message.row}"].number_format, "General")
        finally:
            book.close()

    def test_uncovered_records_are_reported_and_preserved(self):
        gate = check_coverage(self.data, self.coverage)
        self.assertEqual(gate.status, "PASS")
        self.assertEqual({issue[0] for issue in gate.findings}, {"1002", "4001", "6001"})
        book, _ = build(TEMPLATE, self.data, self.coverage, self.adapter.ar_account)
        try:
            self.assertEqual(book[TB_SHEET].max_row, len(self.data.tb) + 1)
            self.assertEqual(book[GL_SHEET].max_row, len(self.data.gl) + 1)
        finally:
            book.close()

    def test_level_peers_with_a_shared_prefix_are_both_included(self):
        first_detail = self.data.tb[3]
        peer = replace(first_detail, account_code=first_detail.account_code + "0")
        data = replace(self.data, tb=(*self.data.tb, peer))
        coverage = select_coverage(data, self.adapter.ar_account)
        self.assertTrue(set(self.coverage.detail_rows) < set(coverage.detail_rows))

    def test_summary_criteria_and_text_key_share_the_account_code_contract(self):
        code = account_code("001.2/3-4")
        data = replace(self.data, tb=(replace(self.data.tb[0], account_code=code), *self.data.tb[1:]))
        book, _ = build(TEMPLATE, data, select_coverage(data, code), code)
        try:
            self.assertEqual(book[SUMMARY]["B3"].value, code)
            self.assertEqual(book[TB_SHEET]["L2"].value, f"acct:{code}")
            self.assertEqual(book[TB_SHEET]["L2"].data_type, "s")
            self.assertEqual(book[SUMMARY]["B7"].value, f"=ROUND({raw_sum('G')},2)")
        finally:
            book.close()
        for wildcard in ("*", "?", "~"):
            with self.subTest(wildcard=wildcard), self.assertRaises(ValueError):
                account_code(f"001{wildcard}2")

    def test_template_formula_text_is_preserved_and_detail_expands(self):
        data = replace(self.data, gl=(replace(self.data.gl[0], summary="=1+1"), *self.data.gl[1:]))
        book, layout = build(TEMPLATE, data, self.coverage, self.adapter.ar_account)
        template = load_workbook(TEMPLATE)
        try:
            preset = {key: value for key, value in formulas(template).items() if key[0] == SUMMARY}
            actual = formulas(book)
            self.assertEqual({key: actual[key] for key in preset}, preset)
            self.assertEqual(layout.detail_total, 9)
            self.assertEqual(book[DETAIL]["F8"].value, "=ROUND(C8+D8-E8,2)")
            self.assertEqual(book[DETAIL]["F9"].value, "=ROUND(SUM(F6:F8),2)")
            self.assertEqual(book[DETAIL]["F6"].style_id, book[DETAIL]["F8"].style_id)
            self.assertEqual(list(book.defined_names[DETAIL_TOTAL].destinations), [(DETAIL, "$A$9")])
            self.assertEqual(book[GL_SHEET]["C2"].data_type, "s")
            self.assertEqual(book[GL_SHEET]["C2"].value, "=1+1")
            with TemporaryDirectory(dir=ROOT / "data" / "output") as temp:
                source, target = (Path(temp) / name for name in ("before.xlsx", "after.xlsx"))
                book.save(source)
                write_caches(source, target, {(SUMMARY, "B9"): Cache("n", "1950")})
                saved = load_workbook(target)
                cached = load_workbook(target, data_only=True)
                try:
                    self.assertEqual(formulas(saved), actual)
                    self.assertEqual(cached[SUMMARY]["B9"].value, 1950)
                finally:
                    saved.close()
                    cached.close()
        finally:
            book.close()
            template.close()

    def test_formula_health_rejects_static_values_and_all_required_errors(self):
        book, layout = build(TEMPLATE, self.data, self.coverage, self.adapter.ar_account)
        try:
            book[SUMMARY]["B9"] = 1950
            actual = formulas(book)
            codes = ("#REF!", "#VALUE!", "#NAME?", "#N/A", "#DIV/0!")
            caches = {(SUMMARY, addr): Cache("e", code)
                      for addr, code in zip(("B6", "B7", "B8", "B10", "B11"), codes, strict=True)}
            calculation = Calculation("xlwings", caches, scan(actual, caches))
            gate = check_excel(layout, actual, calculation)
            self.assertEqual(gate.status, "FAILED")
            messages = {row[1] for row in gate.findings}
            self.assertTrue(set(codes) <= messages)
            self.assertIn("预期公式缺失或单元格为静态值", messages)
            unavailable = Calculation("unavailable", {}, ())
            self.assertEqual(check_excel(layout, actual, unavailable).status, "FAILED")
        finally:
            book.close()

    def test_cross_table_difference_of_one_cent_fails(self):
        book, layout = build(TEMPLATE, self.data, self.coverage, self.adapter.ar_account)
        try:
            caches = {(SUMMARY, "B9"): Cache("n", "1950.00"), (DETAIL, "F9"): Cache("n", "1949.99")}
            caches.update({(DETAIL, cell): Cache("n", value)
                           for cell, value in (("F6", "975"), ("F7", "675"), ("F8", "300"))})
            gate, _ = check_cross(self.data, self.coverage, layout, Calculation("xlwings", caches, ()))
            self.assertEqual(gate.status, "FAILED")
            self.assertEqual({issue[2] for issue in gate.findings}, {Decimal("0.01"), Decimal("-0.01")})
        finally:
            book.close()


if __name__ == "__main__":
    unittest.main()
