from pathlib import Path
from types import SimpleNamespace
import unittest
from unittest.mock import patch

from motir_export_report import Article, ExportReportError, parse_report


class ReportTableParsingTests(unittest.TestCase):
    def setUp(self):
        self.article = Article(
            report_month="2026-10",
            article_id="test",
            title="2026년 10월 수출입 동향",
            published="2026-11-01",
            department="수출입과",
            post_url="https://example.com/report",
            pdf_url="https://example.com/report.pdf",
            pdf_filename="report.pdf",
        )
        self.summary = (
            "2026년 10월 수출입 동향\n"
            "수출 1,153억 달러(+93.7%), 수입 728억 달러(+36.0%), "
            "수지 425억 달러 흑자\n"
            "일평균 수출은 76% 증가한 52.4억 달러"
        )
        self.tables = (
            Path(__file__).parent / "fixtures" / "report_tables_2026_10.txt"
        ).read_text(encoding="utf-8")

    def parse(self, tables):
        reader = SimpleNamespace(pages=[None] * 6)
        with patch(
            "motir_export_report.extract_pdf_text",
            return_value=(reader, [self.summary, tables]),
        ):
            return parse_report(self.article, b"unused PDF payload")

    def test_october_table_whitespace_preserves_values(self):
        for heading, marker in (
            ("9대 주요지역별", "□ (지역)"),
            ("9대  주요지역별", "□ (지역)"),
            ("9 대\n주요  지역별", "□\u3000( 지역 )"),
        ):
            with self.subTest(heading=heading, marker=marker):
                tables = self.tables.replace("9대  주요지역별", heading)
                tables = tables.replace("□ (지역)", marker)
                report = self.parse(tables)
                products = {row["name"]: row for row in report["products"]}
                regions = {row["name"]: row for row in report["regions"]}
                self.assertEqual(len(products), 20)
                self.assertEqual(len(regions), 9)
                self.assertEqual(products["반도체"]["export_100m_usd"], 564.4)
                self.assertEqual(products["선박"]["yoy_pct"], -43.6)
                self.assertEqual(products["생활용품"]["yoy_pct"], -13.4)
                self.assertEqual(products["섬유"]["export_100m_usd"], 9.6)
                self.assertEqual(regions["중국"]["export_100m_usd"], 269.8)
                self.assertEqual(regions["CIS"]["yoy_pct"], -31.6)
                self.assertEqual(report["summary"]["daily_average_exports_100m_usd"], 52.4)
                self.assertEqual(report["validation"]["status"], "valid")

    def test_missing_region_marker_has_actionable_error(self):
        with self.assertRaisesRegex(ExportReportError, "지역 구분 제목"):
            self.parse(self.tables.replace("□ (지역)", ""))


if __name__ == "__main__":
    unittest.main()
