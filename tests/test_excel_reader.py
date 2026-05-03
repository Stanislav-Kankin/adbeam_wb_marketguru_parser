from pathlib import Path
from tempfile import TemporaryDirectory
import unittest

from openpyxl import Workbook

from adbeam_excel_parser.excel_reader import extract_website_rows, looks_like_website, read_excel_summary


class ExcelReaderWebsiteDetectionTests(unittest.TestCase):
    def test_looks_like_website_rejects_email_addresses(self) -> None:
        self.assertFalse(looks_like_website("buyer@example.ru"))
        self.assertFalse(looks_like_website("info@yoot.pro"))
        self.assertTrue(looks_like_website("http://example.ru"))
        self.assertTrue(looks_like_website("www.example.ru"))
        self.assertTrue(looks_like_website("example.ru/catalog"))

    def test_detected_website_column_prevents_email_row_fallback(self) -> None:
        with TemporaryDirectory() as temp_dir:
            workbook_path = Path(temp_dir) / "websites.xlsx"
            workbook = Workbook()
            worksheet = workbook.active
            worksheet.title = "Контрагенты"
            worksheet.append(["Наименование", "Электронная почта", "Ссылка на сайт"])
            worksheet.append(["Only email", "buyer@example.ru", None])
            worksheet.append(["Only site", None, "http://example.ru"])
            worksheet.append(["Both", "sales@example.ru", "example.com/catalog"])
            workbook.save(workbook_path)
            workbook.close()

            summary = read_excel_summary(workbook_path)
            website_rows = extract_website_rows(workbook_path)

        self.assertEqual(summary.website_columns, ["Ссылка на сайт"])
        self.assertEqual(summary.rows_with_websites, 2)
        self.assertIsNone(summary.preview[0].website)
        self.assertEqual([row.row_index for row in website_rows.rows], [3, 4])
        self.assertEqual([row.website for row in website_rows.rows], ["http://example.ru", "example.com/catalog"])

    def test_row_fallback_still_finds_url_when_no_website_column_exists(self) -> None:
        with TemporaryDirectory() as temp_dir:
            workbook_path = Path(temp_dir) / "fallback.xlsx"
            workbook = Workbook()
            worksheet = workbook.active
            worksheet.title = "Контрагенты"
            worksheet.append(["Наименование", "Электронная почта", "Комментарий"])
            worksheet.append(["Only email", "buyer@example.ru", None])
            worksheet.append(["Site in note", None, "example.ru"])
            workbook.save(workbook_path)
            workbook.close()

            summary = read_excel_summary(workbook_path)
            website_rows = extract_website_rows(workbook_path)

        self.assertEqual(summary.website_columns, [])
        self.assertEqual(summary.rows_with_websites, 1)
        self.assertEqual([row.row_index for row in website_rows.rows], [3])
        self.assertEqual(website_rows.rows[0].website, "example.ru")


if __name__ == "__main__":
    unittest.main()
