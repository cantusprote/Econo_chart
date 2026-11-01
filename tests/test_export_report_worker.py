import io
import queue
from types import SimpleNamespace
import unittest
from unittest.mock import Mock, patch

from econo_ui import EconoUpdaterApp


class ExportReportWorkerTests(unittest.TestCase):
    def run_worker(self, output, return_code):
        process = SimpleNamespace(stdout=io.StringIO(output), wait=Mock(return_value=return_code))
        app = SimpleNamespace(events=queue.Queue())
        with patch("econo_ui.subprocess.Popen", return_value=process):
            EconoUpdaterApp._export_report_worker(app)
        events = []
        while not app.events.empty():
            events.append(app.events.get_nowait())
        return events

    def test_parser_failure_reaches_error_event(self):
        reason = "20대 품목·9대 지역 표가 있는 페이지를 찾지 못했습니다."
        events = self.run_worker(f"VALIDATE 1,748,879 bytes\nFAILED {reason}\n", 1)
        self.assertEqual(events[-1], ("export_error", reason))
        self.assertIn(("export_line", f"FAILED {reason}"), events)

    def test_unexpected_failure_keeps_exit_code(self):
        self.assertEqual(self.run_worker("", 2)[-1], ("export_error", "종료 코드 2"))

    def test_missing_result_is_reported_as_failure(self):
        self.assertEqual(
            self.run_worker("", 0)[-1],
            ("export_error", "처리 결과 데이터를 받지 못했습니다."),
        )

    def test_success_keeps_result_json(self):
        result = '{"report_month": "2026-10"}'
        self.assertEqual(
            self.run_worker(f"RESULT_JSON {result}\n", 0)[-1],
            ("export_done", result),
        )


if __name__ == "__main__":
    unittest.main()
