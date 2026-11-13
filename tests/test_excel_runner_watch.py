"""Fixtures for tools/excel_runner_watch.py (#295)."""

from __future__ import annotations

import json
import re
import sys
import tempfile
import unittest
from datetime import datetime, timezone
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "tools"))

import check_issues  # noqa: E402
import excel_runner_watch as W  # noqa: E402

REGRESSION = "Run complete VBA regression suite through Excel COM"
EXPORT = "Export selected grids through Excel COM"
LOST = ["The self-hosted runner lost communication with the server. Verify the machine is running."]
NOW = datetime(2026, 11, 13, 12, 0, tzinfo=timezone.utc)


def step(name: str, conclusion: str | None = "success") -> dict[str, object]:
    return {"name": name, "conclusion": conclusion}


def job(steps: list[dict[str, object]], conclusion: str = "failure", job_id: int = 7,
        name: str = "Run Excel/VBA regression harness") -> dict[str, object]:
    return {"id": job_id, "name": name, "conclusion": conclusion, "steps": steps}


def run(attempt: int = 1, conclusion: str = "failure") -> dict[str, object]:
    return {"id": 1, "run_attempt": attempt, "conclusion": conclusion}


class Retry(unittest.TestCase):
    def decide(self, jobs: list[dict[str, object]], notes: dict[str, list[str]] | None = None,
               attempt: int = 1, conclusion: str = "failure") -> bool:
        return W.retry_decision(run(attempt, conclusion), jobs, notes or {})[0]

    def test_failures_before_excel_are_rerun_once(self) -> None:
        setup = job([step("Set up job", "failure"), step(REGRESSION, "skipped")])
        checkout = job([step("Set up job"), step("Check out repository", "failure"),
                        step(REGRESSION, "skipped")])
        for failed in (setup, checkout):
            self.assertTrue(self.decide([failed]))
            self.assertFalse(self.decide([failed], attempt=2), "never a second automatic re-run")

    def test_a_lost_runner_is_rerun(self) -> None:
        lost = job([step("Set up job"), step("Check out repository"), step(REGRESSION, None)])
        self.assertTrue(self.decide([lost], {"7": LOST}))
        self.assertFalse(self.decide([lost], {}), "no annotation, no proof of a lost runner")
        self.assertFalse(self.decide([lost], {"8": LOST}), "the annotation of another job")

    def test_harness_and_export_failures_are_never_rerun(self) -> None:
        regression = job([step("Set up job"), step("Check out repository"),
                          step(REGRESSION, "failure")])
        export = job([step("Set up job"), step("Check out repository"), step(REGRESSION),
                      step(EXPORT, "failure")])
        both = job([step("Check out repository", "failure"), step(REGRESSION, "failure")])
        unknown = job([step("Set up job"), step("Upload exact-SHA VBA certification evidence",
                                                "failure")])
        for failed in (regression, export, both, unknown):
            self.assertFalse(self.decide([failed]))
            self.assertFalse(self.decide([failed], {"7": LOST}),
                             "a lost-runner note does not excuse a failed harness step")

    def test_every_failed_job_must_be_infrastructure(self) -> None:
        infra = job([step("Set up job", "failure")], job_id=7, name="select")
        harness = job([step(REGRESSION, "failure")], job_id=8)
        self.assertFalse(self.decide([infra, harness]))
        self.assertTrue(self.decide([infra, job([step(REGRESSION)], "success", 8)]))

    def test_other_conclusions_are_left(self) -> None:
        setup = job([step("Set up job", "failure")])
        for conclusion in ("success", "cancelled", "timed_out", None):
            self.assertFalse(self.decide([setup], conclusion=conclusion))  # type: ignore[arg-type]
        self.assertFalse(self.decide([job([step(REGRESSION)], "success")]))
        cancelled = job([step("Set up job"), step(REGRESSION, None)], "cancelled")
        self.assertFalse(self.decide([cancelled], {"7": LOST}))


def queued_run(run_id: int, minutes: int, status: str = "queued") -> dict[str, object]:
    created = NOW.timestamp() - minutes * 60
    return {"id": run_id, "name": "Excel VBA Regression", "head_branch": "feat/x",
            "status": status,
            "created_at": datetime.fromtimestamp(created, timezone.utc).isoformat()}


def excel_job(status: str, minutes: int | None = None) -> dict[str, object]:
    out: dict[str, object] = {"name": "Run Excel/VBA regression harness", "status": status,
                              "labels": ["self-hosted", "Windows", "X64", "excel"]}
    if minutes is not None:
        created = NOW.timestamp() - minutes * 60
        out["created_at"] = datetime.fromtimestamp(created, timezone.utc).isoformat()
    return out


class Watch(unittest.TestCase):
    def test_a_quiet_queue_reports_nothing(self) -> None:
        runs = [queued_run(1, 10), queued_run(2, 20, "in_progress"), queued_run(3, 500, "completed")]
        jobs = {"2": [excel_job("in_progress")]}
        self.assertEqual(W.findings(runs, jobs, NOW), [])

    def test_a_run_without_jobs_is_reported_after_the_limit(self) -> None:
        self.assertEqual(W.findings([queued_run(1, W.STUCK_MINUTES)], {}, NOW), [])
        found = W.findings([queued_run(1, W.STUCK_MINUTES + 1)], {}, NOW)
        self.assertEqual(len(found), 1)
        self.assertIn("no job was created", found[0])
        self.assertIn("run 1", found[0])

    def test_an_offline_runner_is_reported(self) -> None:
        runs = [queued_run(1, W.OFFLINE_MINUTES + 5, "in_progress")]
        found = W.findings(runs, {"1": [excel_job("queued")]}, NOW)
        self.assertEqual(len(found), 1)
        self.assertIn("probably offline", found[0])
        hosted = dict(excel_job("queued"), labels=["ubuntu-latest"])
        self.assertEqual(W.findings(runs, {"1": [hosted]}, NOW), [])

    def test_a_busy_runner_gets_the_longer_limit(self) -> None:
        runs = [queued_run(1, W.OFFLINE_MINUTES + 5), queued_run(2, 5, "in_progress")]
        jobs = {"1": [excel_job("queued")], "2": [excel_job("in_progress")]}
        self.assertEqual(W.findings(runs, jobs, NOW), [])
        runs[0] = queued_run(1, W.BUSY_MINUTES + 5)
        found = W.findings(runs, jobs, NOW)
        self.assertEqual(len(found), 1)
        self.assertIn("held the runner", found[0])

    def test_the_job_wait_counts_from_the_job(self) -> None:
        runs = [queued_run(1, W.OFFLINE_MINUTES + 30, "in_progress")]
        self.assertEqual(W.findings(runs, {"1": [excel_job("queued", minutes=10)]}, NOW), [])

    def test_old_runs_are_left_out_and_bad_times_fail(self) -> None:
        self.assertEqual(W.findings([queued_run(1, W.MAX_AGE_HOURS * 60 + 1)], {}, NOW), [])
        with self.assertRaises(W.WatchError):
            W.findings([dict(queued_run(1, 90), created_at="yesterday")], {}, NOW)

    def test_the_alert_passes_issue_triage(self) -> None:
        lines = W.findings([queued_run(1, 90)], {}, NOW)
        issue = W.alert(lines, NOW, "2026-12-18")
        self.assertEqual(issue["title"], W.ALERT_TITLE)
        self.assertIn("run 1", issue["body"])
        self.assertEqual(check_issues.target_end_date(issue["body"]), "2026-11-20")
        self.assertIn("Planned start date: 2026-11-13 (UTC check date)", issue["body"])
        self.assertIsNotNone(check_issues.dependency_section(issue["body"]))
        self.assertTrue(re.search(r"^- \[ \] ", issue["body"], re.M))
        near = W.alert(lines, NOW, "2026-11-15")
        self.assertIn("Target end date: 2026-11-14 (Europe/Rome)", near["body"])
        with self.assertRaises(W.WatchError):
            W.alert(lines, NOW, "2026-11-13")

    def test_command_line(self) -> None:
        with tempfile.TemporaryDirectory() as d:
            root = Path(d)
            (root / "jobs").mkdir()
            (root / "runs.json").write_text(json.dumps({"workflow_runs": [queued_run(1, 90)]}))
            out = root / "alert.json"
            args = ["watch", "--runs", str(root / "runs.json"), "--jobs-dir", str(root / "jobs"),
                    "--now", "2026-11-13T12:00:00Z", "--milestone-due", "2026-12-18",
                    "--output", str(out)]
            self.assertEqual(W.main(args), 0)
            self.assertEqual(json.loads(out.read_text())["title"], W.ALERT_TITLE)
            out.unlink()
            (root / "runs.json").write_text(json.dumps({"workflow_runs": []}))
            self.assertEqual(W.main(args), 0)
            self.assertFalse(out.exists())
            (root / "runs.json").write_text("not json")
            self.assertEqual(W.main(args), 1)


class StepNames(unittest.TestCase):
    def test_the_classified_steps_exist_in_the_workflow(self) -> None:
        workflow = (ROOT / ".github/workflows/excel-vba-regression.yml").read_text(encoding="utf-8")
        for name in W.HARNESS_STEPS | (W.INFRASTRUCTURE_STEPS - {"Set up job"}):
            self.assertIn(f"- name: {name}\n", workflow, name)


if __name__ == "__main__":
    unittest.main()
