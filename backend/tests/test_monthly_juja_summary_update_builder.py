import unittest
import io
import openpyxl
from app.services.monthly_juja_summary_update_builder import (
    build_monthly_juja_summary_update_excel_report,
    STATIC_DAILY_WIDTHS,
    NAIROBI_DAILY_WIDTHS,
    CONSOLIDATED_DAILY_WIDTHS,
    HOURLY_WIDTHS_HSWIM_THIKA,
    HOURLY_WIDTHS_MDECK_THIKA,
)


class TestMonthlyJujaSummaryUpdateBuilder(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.buffer = build_monthly_juja_summary_update_excel_report(2026, 10)
        cls.wb = openpyxl.load_workbook(cls.buffer, data_only=False)

    def test_buffer_valid_zip_openxml(self):
        self.assertIsInstance(self.buffer, io.BytesIO)
        raw = self.buffer.getvalue()
        self.assertTrue(raw.startswith(b"PK\x03\x04"))

    def test_sheet_names_and_order(self):
        expected_sheets = [
            "M-DECK THIKA BD",
            "HSWIM THIKA BD HOURLY ",
            "M-DECK THIKA BD HOURLY",
            "M-DECK NAIROBI BD",
            "HSWIM NRB BD HOURLY",
            "M-DECK NRB BD HOURLY",
            "JUJA WB DAILY SUMMARY",
            "JUJA HSWIM summary",
            "M-Decks hourly summary",
        ]
        self.assertEqual(self.wb.sheetnames, expected_sheets)

    def test_sheet_tab_colors(self):
        # Thika sheets should have coral tab color #FFA366
        for sname in ["M-DECK THIKA BD", "HSWIM THIKA BD HOURLY ", "M-DECK THIKA BD HOURLY"]:
            ws = self.wb[sname]
            self.assertIsNotNone(ws.sheet_properties.tabColor)
            self.assertEqual(ws.sheet_properties.tabColor.rgb, "00FFA366")

        # Nairobi sheets should have peach tab color #FFC266
        for sname in ["M-DECK NAIROBI BD", "HSWIM NRB BD HOURLY", "M-DECK NRB BD HOURLY"]:
            ws = self.wb[sname]
            self.assertIsNotNone(ws.sheet_properties.tabColor)
            self.assertEqual(ws.sheet_properties.tabColor.rgb, "00FFC266")

        # Summary sheets have no tab color
        for sname in ["JUJA WB DAILY SUMMARY", "JUJA HSWIM summary", "M-Decks hourly summary"]:
            ws = self.wb[sname]
            self.assertIsNone(ws.sheet_properties.tabColor)

    def test_default_sheet_format(self):
        for sname in self.wb.sheetnames[:7]:
            ws = self.wb[sname]
            self.assertEqual(ws.sheet_format.defaultColWidth, 9.0)
            self.assertEqual(ws.sheet_format.defaultRowHeight, 14.25)

        for sname in self.wb.sheetnames[7:]:
            ws = self.wb[sname]
            self.assertEqual(ws.sheet_format.defaultColWidth, 9.0)
            self.assertEqual(ws.sheet_format.defaultRowHeight, 15.0)

    def test_juja_wb_daily_summary_black_font(self):
        ws7 = self.wb["JUJA WB DAILY SUMMARY"]
        # Check sample data rows and total row
        for r in [9, 15, 25, 39, 40]:
            for c in [2, 5, 10, 21]:
                cell = ws7.cell(r, c)
                self.assertIsNotNone(cell.font)
                self.assertEqual(cell.font.color.rgb, "00000000")

    def test_column_dimensions_match_specs(self):
        ws1 = self.wb["M-DECK THIKA BD"]
        for col_let, width in STATIC_DAILY_WIDTHS.items():
            self.assertEqual(ws1.column_dimensions[col_let].width, width)

        ws4 = self.wb["M-DECK NAIROBI BD"]
        for col_let, width in NAIROBI_DAILY_WIDTHS.items():
            self.assertEqual(ws4.column_dimensions[col_let].width, width)

        ws7 = self.wb["JUJA WB DAILY SUMMARY"]
        for col_let, width in CONSOLIDATED_DAILY_WIDTHS.items():
            self.assertEqual(ws7.column_dimensions[col_let].width, width)


if __name__ == "__main__":
    unittest.main()
