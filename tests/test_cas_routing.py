from datetime import date
from decimal import Decimal
import re
import unittest
from unittest.mock import patch

from openpyxl import Workbook
from yaml import safe_load

from core.cas_checker import CAS_SHEET, scan_routed, write_cas_sheet
from core.cas_routing import ROOT, load_routed_rules, route_cas, yaml_mapping
from core.schema import AuditError, StandardData, StandardGL, StandardTB, ZERO


class CasRoutingTest(unittest.TestCase):
    def test_registry_and_module_statuses_agree(self):
        registry = yaml_mapping(ROOT / "rules" / "cas_registry.yaml")
        records = registry["standards"]
        self.assertEqual([record["cas_no"] for record in records], list(range(43)))
        for record in records:
            with self.subTest(cas_no=record["cas_no"]):
                status = "distilled" if record["priority"] == "P0高频" else "stub"
                self.assertEqual(record["status"], status)
                files = record["raw_source_pattern"]
                for pattern in files["standard"] + files["application_guidance"]:
                    re.compile(pattern)
                rules = yaml_mapping(ROOT / record["yaml_path"])
                self.assertEqual(rules["cas_no"], record["cas_no"])
                self.assertEqual(rules["name"], record["name"])
                self.assertEqual(rules["status"], status)
                if status == "stub":
                    self.assertEqual(rules["rules"], [])
                card = (ROOT / record["md_path"]).read_text(encoding="utf-8")
                metadata = safe_load(card.split("---", maxsplit=2)[1])
                self.assertEqual(metadata["status"], status)
                if status == "distilled":
                    sources = record["online_sources"]
                    self.assertEqual(metadata["source_urls"], list(sources.values()))
                    self.assertIn(f"/casg/{record['cas_no']:02d}.html", sources["guide_2024"])
                    body = card.split("## 一、核心确认与计量条件", maxsplit=1)[1]
                    self.assertGreaterEqual(len(re.findall(r"[\u3400-\u9fff]", body)), 800)
                    self.assertLessEqual(len(re.sub(r"\s", "", body)), 1500)
        self.assertEqual(records[15]["replaced_by"], 14)
        self.assertEqual(records[26]["replaced_by"], 25)
        self.assertIn("2027、2029、2030", records[30]["version_note"])
        self.assertIn("/cas/30b.html", records[30]["online_sources"]["standard_2026"])

    def test_route_exact_then_delimited_parent_without_broad_prefix(self):
        result = route_cas(("1122.01", "1122", "11220"))
        self.assertEqual(result.unmatched_codes, ("11220",))
        self.assertEqual(result.hard_rules[0].name, "cas_00_basic.yaml")
        self.assertEqual(sum(path.name == "cas_14_revenue.yaml" for path in result.hard_rules), 1)
        self.assertIn("cas_22_ar_ecl.md", {path.name for path in result.soft_refs})
        self.assertNotIn("cas_01_inventory.yaml", {path.name for path in result.hard_rules})
        self.assertEqual(route_cas(("11220",)).hard_rules, result.hard_rules[:1])

    def test_all_configured_routes_resolve_without_historical_rules(self):
        router = yaml_mapping(ROOT / "rules" / "cas_router.yaml")
        routes = router["accounts"]
        required = {"1001", "1002", "1122", "1401", "1402", "1403", "1404", "1405", "1601", "1701",
                    "2202", "2205", "6001", "6601", "6602", "6603", "6801"}
        self.assertTrue(required <= set(routes))
        for code in routes:
            with self.subTest(account_code=code):
                result = route_cas((code,))
                self.assertFalse(result.unmatched_codes)
                self.assertNotIn("cas_15_construction_contracts.yaml", {path.name for path in result.hard_rules})
                self.assertNotIn("cas_26_reinsurance_contracts.yaml", {path.name for path in result.hard_rules})
        named = {f"TEST{index}": name for index, name in enumerate(router["named_accounts"], 1)}
        combined = route_cas((*routes, *named), account_names=named)
        self.assertFalse(combined.unmatched_codes)
        registry = yaml_mapping(ROOT / "rules" / "cas_registry.yaml")
        p0 = [record for record in registry["standards"] if record["priority"] == "P0高频"]
        self.assertTrue({ROOT / record["yaml_path"] for record in p0} <= set(combined.hard_rules))
        self.assertTrue({ROOT / record["md_path"] for record in p0} <= set(combined.soft_refs))
        load_routed_rules(combined)

    def test_named_routes_follow_confirmed_meanings_without_guessing_codes(self):
        selected = {"1601": "使用权资产", "LEASEDEP": "使用权资产累计折旧", "CONTRACT": "合同负债"}
        route = route_cas(("1601.01", "LEASEDEP", "CONTRACT"), account_names=selected)
        self.assertFalse(route.unmatched_codes)
        self.assertIn("cas_21_leases.yaml", {path.name for path in route.hard_rules})
        self.assertNotIn("cas_04_fixed_assets.yaml", {path.name for path in route.hard_rules})
        self.assertEqual(route_cas(("LEASEDEP",)).unmatched_codes, ("LEASEDEP",))

        amount = Decimal("11.00")
        records = (
            StandardTB("1601", "合成使用权资产", 1, "借", ZERO, ZERO, ZERO, ZERO, ZERO, amount),
            StandardTB("LEASEDEP", "合成折旧", 1, "贷", ZERO, ZERO, ZERO, ZERO, amount, ZERO),
            StandardTB("CONTRACT", "合成合同负债", 1, "贷", ZERO, ZERO, ZERO, ZERO, amount, ZERO),
        )
        result = scan_routed(StandardData(records, ()), (*selected, "1604"), account_names=selected)
        self.assertEqual({finding.rule_id for finding in result.findings}, {
            "CAS21_TB_RIGHT_OF_USE_ASSET_DIRECTION", "CAS21_TB_LEASE_CONTRA_DIRECTION",
            "CAS14_TB_CONTRACT_LIABILITY_DIRECTION",
        })

    def test_root_closing_rules_require_confirmation_and_ignore_legitimate_detail(self):
        amount = Decimal("11.00")
        tb = (
            StandardTB("3002", "合成货币兑换", 1, "借", ZERO, ZERO, ZERO, ZERO, ZERO, ZERO),
            StandardTB("3002.USD", "合成外币明细", 2, "贷", ZERO, ZERO, ZERO, ZERO, ZERO, amount),
            StandardTB("6901", "合成以前年度损益调整", 1, "贷", ZERO, ZERO, ZERO, ZERO, ZERO, amount),
        )
        codes = ("3002", "6901")
        forex_id = "CAS19_TB_CURRENCY_EXCHANGE_UNCLOSED"
        adjustment_id = "CAS28_TB_PRIOR_PERIOD_ADJUSTMENT_UNCLOSED"
        open_scan = scan_routed(StandardData(tb, ()), codes)
        self.assertTrue({forex_id, adjustment_id} <= set(open_scan.deferred_rule_ids))
        closed_scan = scan_routed(StandardData(tb, ()), codes, period_closed=True, currency_exchange_closed=True)
        self.assertEqual({item.rule_id for item in closed_scan.findings}, {adjustment_id})
        self.assertIsNone(closed_scan.findings[0].draft_amount)
        detail_scan = scan_routed(StandardData(tb[1:], ()), codes, period_closed=True, currency_exchange_closed=True)
        self.assertIn(forex_id, detail_scan.deferred_rule_ids)

    def test_router_cannot_select_raw_sources(self):
        config = yaml_mapping(ROOT / "rules" / "cas_router.yaml")
        config["accounts"]["1002"]["soft_refs"] = ["docs/cas_raw/example.md"]
        with patch("core.cas_routing.yaml_mapping", return_value=config) as reader:
            with self.assertRaisesRegex(AuditError, "CAS 路由配置无效"):
                route_cas(("1002",))
        self.assertEqual([call.args[0].name for call in reader.call_args_list], ["cas_router.yaml"])

    def test_reads_the_selected_rule_modules_and_skips_stubs(self):
        route = route_cas(("1403", "1481"))
        with patch("core.cas_routing.yaml_mapping", wraps=yaml_mapping) as reader:
            rules = load_routed_rules(route)
        self.assertEqual([call.args[0].name for call in reader.call_args_list],
                         ["cas_00_basic.yaml", "cas_01_inventory.yaml", "cas_42_held_for_sale.yaml"])
        self.assertEqual(set(rules), {"CAS00_GL_VOUCHER_BALANCE", "CAS00_TB_POST_CLOSE_PNL",
                                      "CAS00_TB_CASH_DIRECTION", "CAS01_TB_INVENTORY_DIRECTION"})

    def test_routed_signs_keyword_and_same_scope_allowance(self):
        def tb(code, debit="0.00", credit="0.00", party=None, level=1):
            return StandardTB(code, code, level, "借", ZERO, ZERO, ZERO, ZERO,
                              Decimal(debit), Decimal(credit), party)

        data = StandardData(
            (tb("1601", credit="4.00"), tb("1122", debit="50.00", party="甲"),
             tb("1122", credit="10.00", party="乙"), tb("1122", credit="99.00"),
             tb("2202", debit="9.00", party="丙"), tb("1231", credit="51.00"),
             tb("1221", debit="100.00")),
            (StandardGL(date(2026, 11, 1), "合成凭证", "转回准备·合成", "1603", Decimal("4"), ZERO, ""),
             StandardGL(date(2026, 11, 1), "合成凭证", "转回跌价·合成", "1471", Decimal("4"), ZERO, "")),
        )
        codes = ("1601", "1603", "1122", "2202", "1231")
        allowance_id = "CAS22_TB_AR_ALLOWANCE_EXCEEDS_GROSS"
        self.assertIn(allowance_id, scan_routed(data, codes).deferred_rule_ids)
        result = scan_routed(data, codes, account_pairs={"1231": ("1122",)})
        findings = {item.rule_id: item for item in result.findings}
        self.assertEqual(set(findings), {
            "CAS04_TB_PPE_DIRECTION", "CAS08_GL_IMPAIRMENT_REVERSAL_KEYWORD",
            "CAS14_TB_AR_CONTRACT_LIABILITY", "CAS30_TB_AP_PREPAYMENT", allowance_id,
        })
        self.assertEqual(findings[allowance_id].gross, Decimal("50.00"))
        self.assertIn("差额 1.00", findings[allowance_id].detail)
        self.assertNotIn("合成", findings["CAS08_GL_IMPAIRMENT_REVERSAL_KEYWORD"].detail)
        self.assertIsNone(findings["CAS14_TB_AR_CONTRACT_LIABILITY"].draft_amount)
        self.assertEqual(findings["CAS30_TB_AP_PREPAYMENT"].debit, "预付款项")
        for pairs in ({"1231": ("1122", "1121")}, {"1231.01": ("1122",)}):
            with self.subTest(pairs=pairs):
                incomplete = scan_routed(data, codes, account_pairs=pairs)
                self.assertIn(allowance_id, incomplete.deferred_rule_ids)
                self.assertNotIn(allowance_id, {item.rule_id for item in incomplete.findings})

    def test_trial_runner_uses_confirmed_context_and_writes_findings(self):
        voucher_date = date(2026, 11, 1)
        data = StandardData(
            (StandardTB("6001", "合成收入", 1, "贷", ZERO, ZERO, ZERO, Decimal("10.00"),
                        ZERO, Decimal("10.00")),),
            (StandardGL(voucher_date, "合成凭证", "", "1002", Decimal("12.00"), ZERO, ""),
             StandardGL(voucher_date, "合成凭证", "", "6001", ZERO, Decimal("10.00"), "")),
        )
        unconfirmed = scan_routed(data, ("6001",))
        self.assertFalse(unconfirmed.findings)
        self.assertEqual(set(unconfirmed.deferred_rule_ids),
                         {"CAS00_GL_VOUCHER_BALANCE", "CAS00_TB_POST_CLOSE_PNL"})

        result = scan_routed(data, ("6001",), complete_vouchers=True, period_closed=True)
        self.assertEqual(result.deferred_rule_ids, ())
        voucher, pnl = result.findings
        self.assertIn("借贷差额为 2.00", voucher.detail)
        self.assertIsNone(voucher.draft_amount)
        self.assertEqual((pnl.debit, pnl.credit, pnl.draft_amount),
                         ("合成收入", "本年利润", Decimal("10.00")))
        book = Workbook()
        try:
            write_cas_sheet(book, result.rules, result.findings)
            self.assertEqual(book[CAS_SHEET]["B3"].value, 2)
            self.assertEqual(book[CAS_SHEET]["K7"].value, Decimal("10.00"))
        finally:
            book.close()
