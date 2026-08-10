import unittest
from pathlib import Path
from tempfile import TemporaryDirectory
from unittest.mock import MagicMock, patch

from wb_inn_extractor.models import InspectResult, ResearchRow
from wb_inn_extractor.wb_research import (
    BrowserStartupError,
    BatchInspector,
    _inspect_product_via_browser_network_api,
    _inspect_product_via_public_api,
    _open_system_chrome_session,
    _page_has_requisites_entrypoints,
    _trigger_supplier_tooltip,
    _wait_for_product_page_ready,
)


class ProductReadinessTests(unittest.TestCase):
    def test_product_content_marks_page_ready_before_title_updates(self):
        page = MagicMock()
        page.title.return_value = "Интернет-магазин Wildberries"
        page.evaluate.return_value = True

        ready = _wait_for_product_page_ready(page, nm_id=484399968, timeout_ms=1_000)

        self.assertTrue(ready)
        page.wait_for_timeout.assert_not_called()


class RequisitesTriggerTests(unittest.TestCase):
    def test_hidden_known_trigger_is_still_an_entrypoint(self):
        page = MagicMock()
        page.locator.return_value.count.return_value = 1

        with patch("wb_inn_extractor.wb_research._page_has_empty_results_state", return_value=False):
            found = _page_has_requisites_entrypoints(page, deep_mode=False)

        self.assertTrue(found)

    def test_safe_trigger_clicks_before_hovering(self):
        page = MagicMock()
        locator = MagicMock()

        with (
            patch("wb_inn_extractor.wb_research._page_has_empty_results_state", return_value=False),
            patch("wb_inn_extractor.wb_research._selector_candidates", return_value=[locator]),
            patch("wb_inn_extractor.wb_research._safe_click_seller_tooltip_trigger", return_value=True),
            patch("wb_inn_extractor.wb_research._wait_for_requisites_text", return_value="ИНН: 7725305764"),
            patch("wb_inn_extractor.wb_research._hover_like_human") as hover,
        ):
            text = _trigger_supplier_tooltip(page)

        self.assertEqual(text, "ИНН: 7725305764")
        hover.assert_not_called()


class PublicApiTests(unittest.TestCase):
    def setUp(self):
        self.row = ResearchRow(
            source_sheet="1",
            source_row_index=6,
            wb_nm_id=267843270,
            seller_name_raw="ONEENERGY",
            wb_candidate_url="https://www.wildberries.ru/catalog/267843270/detail.aspx",
        )

    def test_public_api_builds_success_result(self):
        card = {"products": [{"id": 267843270, "supplierId": 159267, "supplier": "ONEENERGY"}]}
        legal = {
            "supplierName": 'ООО "ВАНЭНЕРДЖИ"',
            "inn": "7725305764",
            "ogrn": "1167746134324",
        }
        with TemporaryDirectory() as temp_dir, patch(
            "wb_inn_extractor.wb_research._load_public_wb_json",
            side_effect=[card, legal],
        ):
            result = _inspect_product_via_public_api(6, self.row, Path(temp_dir))

        self.assertIsNotNone(result)
        self.assertEqual(result.parse_status, "SUCCESS")
        self.assertEqual(result.inn, "7725305764")
        self.assertEqual(result.ogrn, "1167746134324")
        self.assertEqual(result.entity_type, "ООО")
        self.assertEqual(result.seller_url, "https://www.wildberries.ru/seller/159267")

    def test_batch_does_not_start_browser_when_api_succeeds(self):
        api_result = InspectResult(row_number=6, url=self.row.wb_candidate_url or "", parse_status="SUCCESS")
        with (
            TemporaryDirectory() as temp_dir,
            patch("wb_inn_extractor.wb_research._inspect_product_via_public_api", return_value=api_result),
            patch("wb_inn_extractor.wb_research.sync_playwright") as playwright,
        ):
            with BatchInspector(Path(temp_dir), headful=True) as inspector:
                result = inspector.inspect_row(6, self.row)

        self.assertIs(result, api_result)
        playwright.assert_not_called()

    def test_browser_network_api_builds_success_result(self):
        card = {"products": [{"id": 267843270, "supplierId": 159267, "supplier": "ONEENERGY"}]}
        legal = {
            "supplierFullName": 'Общество с ограниченной ответственностью "ВАНЭНЕРДЖИ"',
            "inn": "7725305764",
            "ogrn": "1167746134324",
        }
        response = MagicMock(status=200)
        response.json.return_value = card
        response_info = MagicMock()
        response_info.value = response
        response_context = MagicMock()
        response_context.__enter__.return_value = response_info
        page = MagicMock()
        page.expect_response.return_value = response_context
        page.evaluate.return_value = {"status": 200, "text": __import__("json").dumps(legal)}

        with TemporaryDirectory() as temp_dir:
            result = _inspect_product_via_browser_network_api(
                page=page,
                row_number=6,
                research_row=self.row,
                artifacts_dir=Path(temp_dir),
                used_persistent_profile=True,
            )

        self.assertIsNotNone(result)
        self.assertEqual(result.parse_status, "SUCCESS")
        self.assertEqual(result.inn, "7725305764")
        self.assertTrue(result.used_persistent_profile)

    def test_failed_browser_start_resets_inspector_state(self):
        with (
            TemporaryDirectory() as temp_dir,
            patch("wb_inn_extractor.wb_research._inspect_product_via_public_api", return_value=None),
            patch("wb_inn_extractor.wb_research.sync_playwright") as playwright,
            patch.object(BatchInspector, "_restart_context", side_effect=RuntimeError("Chrome failed")),
        ):
            inspector = BatchInspector(Path(temp_dir), headful=True)
            playwright.return_value.start.return_value = MagicMock()

            with self.assertRaises(BrowserStartupError):
                inspector.inspect_row(6, self.row)

        self.assertIsNone(inspector._playwright)
        self.assertIsNone(inspector._context)

    def test_system_chrome_uses_hidden_quic_mode(self):
        playwright = MagicMock()
        browser = MagicMock()
        browser.contexts = [MagicMock()]
        playwright.chromium.connect_over_cdp.return_value = browser
        process = MagicMock()
        process.poll.return_value = 0

        with (
            TemporaryDirectory() as temp_dir,
            patch("wb_inn_extractor.wb_research._find_system_chrome", return_value=Path("C:/Chrome/chrome.exe")),
            patch("wb_inn_extractor.wb_research._reserve_local_port", return_value=9222),
            patch("wb_inn_extractor.wb_research._wait_for_chrome_debug_port"),
            patch("wb_inn_extractor.wb_research.subprocess.Popen", return_value=process) as popen,
        ):
            session = _open_system_chrome_session(playwright, Path(temp_dir) / "profile")

        args = popen.call_args.args[0]
        self.assertIn("--headless=new", args)
        self.assertIn("--enable-quic", args)
        self.assertTrue(any(arg.startswith("--user-agent=Mozilla/5.0") for arg in args))
        self.assertNotIn("--start-maximized", args)
        self.assertIs(session.context, browser.contexts[0])


if __name__ == "__main__":
    unittest.main()
