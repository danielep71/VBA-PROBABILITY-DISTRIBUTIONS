#!/usr/bin/env python3
"""Watch the self-hosted Excel runner and retry its infrastructure failures (#295).

Every Excel regression and export runs on one self-hosted Windows runner. The
GITHUB_TOKEN of a workflow cannot list runners, so this tool reads what a
workflow can read, the runs and their jobs, and draws two kinds of conclusion.

retry
    Decide whether a failed run of the Excel workflow may be re-run once,
    automatically. Only an INFRASTRUCTURE failure qualifies:

      * the failed step is job set-up or the checkout, before Excel starts; or
      * no step failed and GitHub annotated the job with a lost runner
        ("lost communication with the server", "received a shutdown signal").

    A failure of the regression or export step is a HARNESS failure and is
    never re-run automatically, whatever else happened. A run is re-run at
    most once: only attempt 1 qualifies.

watch
    Report what needs a person:

      * a run still queued with no job created after STUCK_MINUTES. GitHub
        sometimes never creates the jobs; such a run cannot be cancelled
        through the API, and starts only when the pull request is closed and
        reopened or a new commit is pushed;
      * a job for the self-hosted runner still queued after OFFLINE_MINUTES
        while no job is running on that runner (the runner is offline), or
        after BUSY_MINUTES even if one is.

    Runs older than MAX_AGE_HOURS are left out: by then the report is noise.
    The alert issue it writes passes Issue triage (tools/check_issues.py),
    like the traffic alert: the workflow adds assignee, labels and milestone.

Usage, from a workflow that fetched the JSON with gh:
    python3 tools/excel_runner_watch.py retry --run run.json --jobs jobs.json \\
        --annotations annotations.json
    python3 tools/excel_runner_watch.py watch --runs runs.json --jobs-dir jobs/ \\
        --now 2026-11-13T12:00:00Z --milestone-due 2026-12-18 --output alert.json

retry prints "rerun" or "leave" and the reason. watch writes the alert file
only when something needs attention. Both exit 0 unless the input is unusable.
"""

from __future__ import annotations

import argparse
import json
import sys
from collections.abc import Mapping, Sequence
from datetime import date, datetime, timedelta, timezone
from pathlib import Path
from typing import Any

INFRASTRUCTURE_STEPS = frozenset({"Set up job", "Check out repository"})
HARNESS_STEPS = frozenset({
    "Run complete VBA regression suite through Excel COM",
    "Export selected grids through Excel COM",
})
LOST_RUNNER = ("lost communication with the server", "received a shutdown signal")

STUCK_MINUTES = 30
OFFLINE_MINUTES = 60
BUSY_MINUTES = 180
MAX_AGE_HOURS = 48
ALERT_DAYS = 7
ALERT_TITLE = "[CI]: The Excel runner needs attention"

Json = Mapping[str, Any]


class WatchError(ValueError):
    """The input cannot be interpreted."""


# --------------------------------------------------------------------------
# retry
# --------------------------------------------------------------------------
def classify_job(job: Json, annotations: Sequence[str]) -> str:
    """"ok", "infrastructure" or "harness" for one job of a failed run."""
    if job.get("conclusion") not in ("failure", "cancelled", "timed_out"):
        return "ok"
    steps = [s for s in job.get("steps") or [] if isinstance(s, Mapping)]
    failed = [str(s.get("name")) for s in steps if s.get("conclusion") == "failure"]
    if any(name in HARNESS_STEPS for name in failed):
        return "harness"
    if failed:
        return "infrastructure" if all(n in INFRASTRUCTURE_STEPS for n in failed) else "harness"
    if job.get("conclusion") == "failure" and any(
            marker in text for text in annotations for marker in LOST_RUNNER):
        return "infrastructure"
    return "harness"


def retry_decision(run: Json, jobs: Sequence[Json],
                   annotations: Mapping[str, Sequence[str]]) -> tuple[bool, str]:
    if run.get("conclusion") != "failure":
        return False, f"run concluded {run.get('conclusion')!r}, not a failure"
    if run.get("run_attempt") != 1:
        return False, f"attempt {run.get('run_attempt')}: a run is re-run at most once"
    kinds = {str(job.get("name")): classify_job(job, annotations.get(str(job.get("id")), ()))
             for job in jobs}
    harness = sorted(name for name, kind in kinds.items() if kind == "harness")
    if harness:
        return False, "not an infrastructure failure: " + ", ".join(harness)
    infrastructure = sorted(name for name, kind in kinds.items() if kind == "infrastructure")
    if not infrastructure:
        return False, "no failed job to classify"
    return True, "infrastructure failure before or outside the harness: " + ", ".join(infrastructure)


# --------------------------------------------------------------------------
# watch
# --------------------------------------------------------------------------
def parse_time(text: Any) -> datetime:
    if not isinstance(text, str):
        raise WatchError(f"not a timestamp: {text!r}")
    try:
        return datetime.fromisoformat(text.replace("Z", "+00:00")).astimezone(timezone.utc)
    except ValueError as exc:
        raise WatchError(f"not a timestamp: {text!r}") from exc


def minutes_since(now: datetime, text: Any) -> float:
    return (now - parse_time(text)).total_seconds() / 60


def self_hosted(job: Json) -> bool:
    return "self-hosted" in [str(label) for label in job.get("labels") or []]


def findings(runs: Sequence[Json], jobs: Mapping[str, Sequence[Json]], now: datetime) -> list[str]:
    """One line per run or job that needs a person, oldest first."""
    running = any(self_hosted(job) and job.get("status") == "in_progress"
                  for listed in jobs.values() for job in listed)
    limit = BUSY_MINUTES if running else OFFLINE_MINUTES
    found: list[tuple[float, str]] = []
    for run in runs:
        waiting = run.get("status") in ("queued", "pending", "waiting", "requested")
        # A run whose first job finished is in_progress while its Excel job
        # still waits for the runner.
        if not waiting and run.get("status") != "in_progress":
            continue
        age = minutes_since(now, run.get("created_at"))
        if age > MAX_AGE_HOURS * 60:
            continue
        listed = jobs.get(str(run.get("id")), ())
        label = f"run {run.get('id')} ({run.get('name')}, branch {run.get('head_branch')})"
        if not listed:
            if waiting and age > STUCK_MINUTES:
                found.append((-age, f"{label}: queued {int(age)} min and no job was created. "
                              "Close and reopen the pull request, or push a new commit to "
                              "the branch."))
            continue
        for job in listed:
            if not (self_hosted(job) and job.get("status") == "queued"):
                continue
            waited = minutes_since(now, job.get("created_at") or run.get("created_at"))
            if waited > limit:
                state = ("another job has held the runner all that time" if running
                         else "no job is running on the runner: it is probably offline")
                found.append((-waited, f"{label}: job \"{job.get('name')}\" has waited "
                              f"{int(waited)} min; {state}."))
    return [text for _, text in sorted(found)]


def alert(lines: Sequence[str], now: datetime, milestone_due: str) -> dict[str, str]:
    try:
        due = date.fromisoformat(milestone_due)
    except ValueError as exc:
        raise WatchError(f"milestone due date is not a date: {milestone_due!r}") from exc
    today = now.date()
    target = min(today + timedelta(days=ALERT_DAYS), due - timedelta(days=1))
    if target < today:
        raise WatchError(f"milestone due {due} leaves no target date after {today}")
    body = [
        "## Findings",
        "",
        f"The scheduled runner check of {now.strftime('%Y-%m-%d %H:%M')} UTC found:",
        "",
        *[f"- {line}" for line in lines],
        "",
        "## What to do",
        "",
        "- [ ] Check that the runner machine is on, awake and connected, and that the "
        "runner process is running (`docs/EXCEL_VBA_CI.md`, \"Hardened runner set-up\").",
        "- [ ] For a run with no job: close and reopen its pull request, or push a new commit.",
        "",
        "The check edits this issue while the condition lasts and closes it when it clears.",
        "",
        "## Dependencies",
        "",
        "None.",
        "",
        "## Planning",
        "",
        f"- Planned start date: {today.isoformat()} (UTC check date)",
        f"- Target end date: {target.isoformat()} (Europe/Rome)",
        "",
    ]
    return {"title": ALERT_TITLE, "body": "\n".join(body)}


# --------------------------------------------------------------------------
# command line
# --------------------------------------------------------------------------
def load(path: str) -> Any:
    try:
        return json.loads(Path(path).read_text(encoding="utf-8"))
    except (OSError, ValueError) as exc:
        raise WatchError(f"cannot read {path}: {exc}") from exc


def listed(data: Any, key: str) -> list[Json]:
    items = data.get(key) if isinstance(data, Mapping) else data
    if not isinstance(items, list):
        raise WatchError(f"expected a list of {key}")
    return [item for item in items if isinstance(item, Mapping)]


def main(argv: Sequence[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    sub = ap.add_subparsers(dest="command", required=True)
    r = sub.add_parser("retry")
    r.add_argument("--run", required=True)
    r.add_argument("--jobs", required=True)
    r.add_argument("--annotations", required=True,
                   help="JSON object: job id -> list of annotation messages")
    w = sub.add_parser("watch")
    w.add_argument("--runs", required=True)
    w.add_argument("--jobs-dir", required=True, help="one <run id>.json per queued run")
    w.add_argument("--now", required=True)
    w.add_argument("--milestone-due", required=True)
    w.add_argument("--output", required=True)
    a = ap.parse_args(argv)
    try:
        if a.command == "retry":
            run = load(a.run)
            notes = load(a.annotations)
            if not isinstance(run, Mapping) or not isinstance(notes, Mapping):
                raise WatchError("run and annotations must be JSON objects")
            rerun, reason = retry_decision(run, listed(load(a.jobs), "jobs"), notes)
            print(("rerun" if rerun else "leave") + ": " + reason)
            return 0
        now = parse_time(a.now)
        runs = listed(load(a.runs), "workflow_runs")
        jobs = {str(run.get("id")): listed(load(str(path)), "jobs")
                for run in runs
                for path in [Path(a.jobs_dir) / f"{run.get('id')}.json"] if path.exists()}
        lines = findings(runs, jobs, now)
        for line in lines:
            print("FINDING: " + line)
        if lines:
            Path(a.output).write_text(json.dumps(alert(lines, now, a.milestone_due)),
                                      encoding="utf-8")
        else:
            print("The Excel runner queue needs no attention.")
        return 0
    except WatchError as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
