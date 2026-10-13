import tempfile
import unittest
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.chart import LineChart, Reference
from openpyxl.styles import Alignment, Font

from temp_export_updater import YearMonth, update_workbook


class ChartFooterRegressionTests(unittest.TestCase):
    def setUp(self):
        self.directory = tempfile.TemporaryDirectory()
        self.addCleanup(self.directory.cleanup)
        self.path = Path(self.directory.name) / "temp_export.xlsx"
        workbook = Workbook()
        data = workbook.active
        data.title = "Data"
        data.append(["month", "days_10", "days_20", "final"])
        for month in range(1, 10):
            data.append([datetime(2026, month, 1), month, month, month])
        chart_sheet = workbook.create_sheet("Chart")
        chart_sheet.merge_cells("A1:H1")
        chart_sheet["A1"] = "Export averages"
        for column, label in enumerate(("month", "days_10", "days_20", "final"), start=1):
            chart_sheet.cell(3, column, label)
        chart_sheet.merge_cells("A13:E16")
        chart_sheet["A13"] = "Source note"
        chart_sheet["A13"].font = Font(italic=True, color="123456")
        chart_sheet["A13"].alignment = Alignment(wrap_text=True)
        chart_sheet.row_dimensions[13].height = 42
        chart_sheet.merge_cells("G14:H14")
        chart_sheet["G14"] = "Additional footer note"
        chart = LineChart()
        chart.add_data(Reference(chart_sheet, min_col=2, max_col=4, min_row=3, max_row=12))
        chart.set_categories(Reference(chart_sheet, min_col=1, min_row=4, max_row=12))
        chart_sheet.add_chart(chart, "G3")
        workbook.save(self.path)
        workbook.close()

    @staticmethod
    def history(*months):
        return {
            "updated_at": "2026-10-14T08:16:09+09:00",
            "observations": {
                month: {"days_10": {"daily_average_exports_100m_usd": 47.5}}
                for month in months
            },
        }

    def assert_saved_layout(self, last_data_row, footer_row):
        workbook = load_workbook(self.path)
        self.addCleanup(workbook.close)
        sheet = workbook["Chart"]
        self.assertEqual(sheet["A1"].value, "Export averages")
        self.assertIn("A1:H1", {str(merged) for merged in sheet.merged_cells.ranges})
        self.assertEqual(sheet.cell(footer_row, 1).value, "Source note")
        self.assertTrue(sheet.cell(footer_row, 1).font.italic)
        self.assertEqual(sheet.cell(footer_row, 1).font.color.rgb, "00123456")
        self.assertTrue(sheet.cell(footer_row, 1).alignment.wrap_text)
        self.assertEqual(sheet.row_dimensions[footer_row].height, 42)
        self.assertEqual(sheet.cell(footer_row + 1, 7).value, "Additional footer note")
        self.assertIn(
            f"A{footer_row}:E{footer_row + 3}",
            {str(merged) for merged in sheet.merged_cells.ranges},
        )
        for row in range(4, last_data_row + 1):
            source_row = row - 2
            self.assertEqual(sheet.cell(row, 2).value, f"='Data'!B{source_row}")
            self.assertEqual(sheet.cell(row, 3).value, f"='Data'!C{source_row}")
        self.assertEqual(len(sheet._charts), 1)
        self.assertEqual(sheet._charts[0].ser[0].val.numRef.f, f"'Chart'!$B$4:$B${last_data_row}")
        self.assertEqual(sheet._charts[0].ser[0].cat.numRef.f, f"'Chart'!$A$4:$A${last_data_row}")
        return workbook

    def test_existing_month_preserves_footer_position(self):
        update_workbook(self.history("2026-09"), self.path)
        self.assert_saved_layout(12, 13)

    def test_tenth_month_moves_footer_and_keeps_pending_values_blank(self):
        update_workbook(self.history("2026-10"), self.path, [YearMonth(2026, 10)])
        workbook = self.assert_saved_layout(13, 14)
        self.assertEqual(workbook["Data"]["B11"].value, 47.5)
        self.assertIsNone(workbook["Data"]["C11"].value)
        self.assertIsNone(workbook["Data"]["D11"].value)
        self.assertIsNone(workbook["Chart"]["D13"].value)
        self.assertEqual(workbook["Data"]["E11"].value, '=IF(OR(B11="",D11=""),"",B11-D11)')

    def test_repeated_updates_only_move_footer_when_data_grows(self):
        history = self.history("2026-10")
        update_workbook(history, self.path)
        self.assertEqual(update_workbook(history, self.path), 0)
        self.assert_saved_layout(13, 14)
        update_workbook(self.history("2026-11"), self.path)
        workbook = self.assert_saved_layout(14, 15)
        self.assertEqual(workbook["Data"].max_row, 12)

    def test_multiple_new_months_crossing_year_move_entire_footer(self):
        update_workbook(
            self.history("2026-10", "2026-11", "2026-12", "2027-01", "2027-02"),
            self.path,
        )
        workbook = self.assert_saved_layout(17, 18)
        self.assertEqual(workbook["Data"]["A15"].value, datetime(2027, 2, 1))


if __name__ == "__main__":
    unittest.main()
