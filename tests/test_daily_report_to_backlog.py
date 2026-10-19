import tempfile
import unittest
from datetime import date
from pathlib import Path

from openpyxl import Workbook

import daily_report_to_backlog as daily


class DailyReportTests(unittest.TestCase):
    def make_workbook(self, report_rows):
        temp_dir = tempfile.TemporaryDirectory()
        self.addCleanup(temp_dir.cleanup)
        path = Path(temp_dir.name) / "daily.xlsx"

        workbook = Workbook()
        setup = workbook.active
        setup.title = daily.SETUP_SHEET
        setup.append(list(daily.SETUP_HEADERS))
        setup.append([
            "HRM",
            "[Daily Report] Bách 10/2026",
            "Task",
            "Daily Report",
            "Normal",
            "Duong Tung Bach",
            None,
            "sprint #6",
            "sprint #6",
            date(2026, 10, 1),
            date(2026, 10, 31),
            None,
            None,
            None,
            "HRM-319",
        ])

        reports = workbook.create_sheet(daily.REPORT_SHEET)
        reports.append(list(daily.REPORT_HEADERS))
        for row in report_rows:
            reports.append(row)
        workbook.save(path)
        return path

    def report_row(self, report_date=date(2026, 10, 14), hours=None):
        return [
            report_date,
            "Wednesday",
            "Retro toàn bộ dự án và bàn giao cho lead mới",
            "Đã tổng hợp trạng thái, rủi ro và hạng mục tồn đọng",
            "Không có",
            "Hỗ trợ lead mới tiếp nhận dự án",
            hours,
            "human-resource-management",
            "",
            "",
            "User-provided",
            "",
            "",
        ]

    def test_read_workbook_and_build_comment(self):
        path = self.make_workbook([self.report_row(hours=8)])
        setup, reports = daily.read_workbook(path)

        self.assertEqual(setup["Existing Issue Key"], "HRM-319")
        self.assertEqual(reports[0]["Report Date"], date(2026, 10, 14))
        body, digest = daily.comment_body(reports[0])
        self.assertIn("Daily Report 14/10/2026", body)
        self.assertIn("Retro toàn bộ dự án", body)
        self.assertIn("Actual Hours", body)
        self.assertIn(digest[:12], body)

    def test_duplicate_report_date_is_rejected(self):
        path = self.make_workbook([self.report_row(), self.report_row()])
        with self.assertRaisesRegex(ValueError, "bị trùng"):
            daily.read_workbook(path)

    def test_invalid_hours_is_rejected(self):
        path = self.make_workbook([self.report_row(hours=25)])
        with self.assertRaisesRegex(ValueError, "0..24"):
            daily.read_workbook(path)

    def test_build_issue_payload_keeps_backlog_fields(self):
        path = self.make_workbook([self.report_row()])
        setup, _ = daily.read_workbook(path)
        context = {
            "project": {"id": 1},
            "issue_type": {"id": 2},
            "priority": {"id": 3},
            "assignee": {"id": 4},
            "category": None,
            "milestone": {"id": 5},
            "version": {"id": 5},
        }

        payload = daily.build_issue_payload(setup, context)

        self.assertEqual(payload["summary"], "[Daily Report] Bách 10/2026")
        self.assertEqual(payload["issueTypeId"], 2)
        self.assertEqual(payload["priorityId"], 3)
        self.assertEqual(payload["assigneeId"], 4)
        self.assertEqual(payload["startDate"], "2026-10-01")
        self.assertEqual(payload["dueDate"], "2026-10-31")
        self.assertEqual(payload["milestoneId[]"], [5])
        self.assertEqual(payload["versionId[]"], [5])

    def test_exact_item_does_not_use_partial_matches(self):
        items = [{"id": 1, "name": "Duong Tung Bach"}, {"id": 2, "name": "Bach Nguyen"}]
        self.assertEqual(daily.exact_item(items, "Dương Tùng Bách", "Assignee")["id"], 1)
        with self.assertRaisesRegex(ValueError, "Không tìm thấy"):
            daily.exact_item(items, "Bach", "Assignee")

    def test_existing_issue_must_match_setup(self):
        path = self.make_workbook([self.report_row()])
        setup, _ = daily.read_workbook(path)
        context = {"space": {"timezone": "Asia/Ho_Chi_Minh"}, "project": {"id": 1}}
        issue = {
            "projectId": 1,
            "summary": "[Daily Report] Bách 10/2026",
            "issueType": {"name": "Task"},
            "priority": {"name": "Normal"},
            "assignee": {"name": "Duong Tung Bach"},
            "startDate": "2026-09-30T17:00:00Z",
            "dueDate": "2026-10-30T17:00:00Z",
            "category": [],
            "milestone": [{"name": "sprint #6"}],
            "versions": [{"name": "sprint #6"}],
            "estimatedHours": None,
            "actualHours": None,
            "resolution": None,
        }

        daily.validate_existing_issue(issue, setup, context)
        issue["summary"] = "Wrong issue"
        with self.assertRaisesRegex(ValueError, "không khớp"):
            daily.validate_existing_issue(issue, setup, context)

    def test_existing_comment_index_handles_markers_and_manual_headings(self):
        comments = [
            {"content": "## Daily Report 14/10/2026\n\nRetro dự án"},
            {"content": "## Daily Report 18/10/2026\n\nDone\n\n<!-- daily-report:2026-10-18:abc123 -->"},
        ]

        indexed = daily.index_existing_comments(comments)

        self.assertEqual(indexed["2026-10-14"], {"legacy"})
        self.assertEqual(indexed["2026-10-18"], {"abc123"})

    def test_state_key_separates_spaces_and_issues(self):
        path = self.make_workbook([self.report_row()])
        setup, _ = daily.read_workbook(path)

        first = daily.state_key("a.backlog.com", "HRM", setup, "HRM-319")
        other_space = daily.state_key("b.backlog.com", "HRM", setup, "HRM-319")
        other_issue = daily.state_key("a.backlog.com", "HRM", setup, "HRM-320")

        self.assertNotEqual(first, other_space)
        self.assertNotEqual(first, other_issue)


if __name__ == "__main__":
    unittest.main()
