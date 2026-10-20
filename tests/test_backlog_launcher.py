import os
import tempfile
import unittest
from contextlib import redirect_stdout
from io import StringIO
from pathlib import Path
from types import SimpleNamespace

import backlog


class BacklogLauncherTests(unittest.TestCase):
    def make_excel(self, name="Sprint 6_Final.xlsx"):
        directory = tempfile.TemporaryDirectory()
        self.addCleanup(directory.cleanup)
        path = Path(directory.name) / name
        path.touch()
        return path

    def test_normalize_excel_path_accepts_quotes_escaped_spaces_and_file_url(self):
        expected = Path("/tmp/My Report.xlsx").resolve(strict=False)
        self.assertEqual(backlog.normalize_excel_path('"/tmp/My Report.xlsx"'), expected)
        self.assertEqual(backlog.normalize_excel_path("/tmp/My\\ Report.xlsx"), expected)
        self.assertEqual(
            backlog.normalize_excel_path("file:///tmp/My%20Report.xlsx"),
            expected,
        )

    def test_infer_milestone_from_filename(self):
        self.assertEqual(backlog.infer_milestone(Path("Sprint 6_Final.xlsx")), "Sprint 6")
        self.assertEqual(backlog.infer_milestone(Path("ke-hoach-sprint-12.xlsx")), "Sprint 12")
        self.assertIsNone(backlog.infer_milestone(Path("planning.xlsx")))

    def test_task_flow_checks_then_runs_without_key_in_arguments(self):
        excel = self.make_excel()
        answers = iter(["1", str(excel), "", "y"])
        calls = []

        def fake_run(command, **kwargs):
            calls.append((command, kwargs))
            return SimpleNamespace(returncode=0)

        result = backlog.main(
            input_func=lambda _prompt: next(answers),
            secret_func=lambda _prompt: "top-secret-key",
            run_func=fake_run,
        )

        self.assertEqual(result, 0)
        self.assertEqual(len(calls), 2)
        self.assertEqual(calls[0][0][-3:], ["--milestone", "Sprint 6", "--check"])
        self.assertIn("--run", calls[1][0])
        self.assertNotIn("top-secret-key", " ".join(calls[1][0]))
        self.assertEqual(calls[1][1]["env"]["BACKLOG_API_KEY"], "top-secret-key")
        self.assertNotEqual(os.environ.get("BACKLOG_API_KEY"), "top-secret-key")

    def test_detected_milestone_can_be_overridden(self):
        excel = self.make_excel()
        answers = iter(["1", str(excel), "Sprint 7", "n"])
        calls = []

        def fake_run(command, **kwargs):
            calls.append((command, kwargs))
            return SimpleNamespace(returncode=0)

        result = backlog.main(
            input_func=lambda _prompt: next(answers),
            secret_func=lambda _prompt: self.fail("cancelled flow must not ask for a key"),
            run_func=fake_run,
        )

        self.assertEqual(result, 0)
        self.assertEqual(calls[0][0][-3:], ["--milestone", "Sprint 7", "--check"])
        self.assertEqual(len(calls), 1)

    def test_daily_flow_can_replace_managed_comments(self):
        excel = self.make_excel("Daily Report.xlsx")
        answers = iter(["2", str(excel), "y", "y"])
        calls = []

        def fake_run(command, **kwargs):
            calls.append((command, kwargs))
            return SimpleNamespace(returncode=0)

        result = backlog.main(
            input_func=lambda _prompt: next(answers),
            secret_func=lambda _prompt: "secret",
            run_func=fake_run,
        )

        self.assertEqual(result, 0)
        self.assertEqual(len(calls), 2)
        self.assertEqual(calls[0][0][-1], "--check")
        self.assertIn("--replace-managed-comments", calls[1][0])
        self.assertEqual(calls[1][0][-1], "--run")

    def test_failed_preflight_does_not_request_key_or_run_upload(self):
        excel = self.make_excel()
        answers = iter(["1", str(excel), ""])
        calls = []

        def fake_run(command, **kwargs):
            calls.append((command, kwargs))
            return SimpleNamespace(returncode=2)

        def fail_secret(_prompt):
            self.fail("API Key must not be requested after a failed check")

        result = backlog.main(
            input_func=lambda _prompt: next(answers),
            secret_func=fail_secret,
            run_func=fake_run,
        )

        self.assertEqual(result, 2)
        self.assertEqual(len(calls), 1)

    def test_interrupted_upload_warns_that_backlog_may_be_partial(self):
        excel = self.make_excel()
        answers = iter(["1", str(excel), "", "y"])
        call_count = 0

        def fake_run(_command, **_kwargs):
            nonlocal call_count
            call_count += 1
            if call_count == 1:
                return SimpleNamespace(returncode=0)
            raise KeyboardInterrupt

        output = StringIO()
        with redirect_stdout(output):
            result = backlog.main(
                input_func=lambda _prompt: next(answers),
                secret_func=lambda _prompt: "secret",
                run_func=fake_run,
            )

        self.assertEqual(result, 130)
        self.assertIn("có thể đã nhận một phần dữ liệu", output.getvalue())


if __name__ == "__main__":
    unittest.main()
