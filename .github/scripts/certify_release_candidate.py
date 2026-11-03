"""Certify one exact release commit from its hosted Excel regression run (#76).

Used by .github/workflows/release.yml on every run, tag push and rehearsal
alike. The retained record in the tree (benchmark/excel_regression_record.json)
names an earlier Excel candidate, because a commit can never contain a record
of itself; it proves that the certified VBA bytes are unchanged, not that Excel
ran the release commit. That proof comes from the artifact of an Excel VBA
Regression run on the exact commit:

1. ``--candidate`` must be a full SHA equal to the checked-out ``HEAD^{commit}``.
2. Successful ``push`` and ``workflow_dispatch`` runs of the policy's Excel
   workflow in this repository whose head commit is the candidate are listed
   through the GitHub API, newest first. Pull request runs are never used: they
   execute a synthetic merge commit, not the commit they report.
3. Each run's ``excel-vba-regression-<run id>`` artifact is downloaded and
   ``benchmark/check_excel_certification.py`` validates its record with
   ``--candidate-sha`` and ``--evidence-ref`` both set to the candidate,
   ``--require-pass`` and ``--log-directory``. A record for any other commit
   fails, even one whose VBA bytes are identical. The record must also name the
   self-hosted runner and the run it was downloaded from.
4. The first run that passes certifies the candidate: CERTIFIED. Otherwise the
   candidate is UNCERTIFIED, with every reason listed. A tag push then fails; a
   rehearsal reports a warning, and the workflow never attests or releases it.

The result goes to the job summary and, as ``certification`` and
``excel_run``, to the step outputs. Run locally with an authenticated gh:

    python3 .github/scripts/certify_release_candidate.py \\
        --repository owner/name --candidate "$(git rev-parse 'HEAD^{commit}')" \\
        --mode rehearsal --work-dir /tmp/excel-certification
"""
from __future__ import annotations

import argparse
import json
import os
import re
import shutil
import subprocess
import sys
from collections.abc import Callable, Sequence
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

SHA40 = re.compile(r"^[0-9a-f]{40}$")
POLICY_PATH = ".github/excel-evidence-policy.json"
CHECKER = "benchmark/check_excel_certification.py"
RECORD = "excel-certification.json"
EVENTS = frozenset({"push", "workflow_dispatch"})
RUNNER_CLASS = "self-hosted-excel"
CERTIFIED = "CERTIFIED"
UNCERTIFIED = "UNCERTIFIED"

Gh = Callable[[Sequence[str]], subprocess.CompletedProcess[bytes]]


class CandidateError(Exception):
    """The candidate cannot be identified, so nothing is looked up."""


def run_gh(args: Sequence[str]) -> subprocess.CompletedProcess[bytes]:
    try:
        return subprocess.run(["gh", *args], capture_output=True, check=False)
    except OSError as exc:
        return subprocess.CompletedProcess(["gh", *args], 127, b"", str(exc).encode())


def git(root: Path, *args: str) -> str:
    proc = subprocess.run(["git", *args], cwd=root, capture_output=True, text=True,
                          check=False)
    if proc.returncode:
        raise CandidateError(proc.stderr.strip() or f"git {' '.join(args)} failed")
    return proc.stdout.strip()


def one_line(text: str) -> str:
    return " ".join(text.split())


def checked_out_candidate(root: Path, candidate: str) -> str:
    """The candidate, after proving it is the commit checked out at ``root``."""
    if not SHA40.fullmatch(candidate):
        raise CandidateError(f"candidate {candidate!r} is not a full lower-case commit SHA")
    head = git(root, "rev-parse", "--verify", "HEAD^{commit}")
    if head != candidate:
        raise CandidateError(f"checked-out commit {head} is not the candidate {candidate}")
    return candidate


def excel_workflow(root: Path, candidate: str) -> str:
    """The Excel workflow path that the candidate's evidence policy names."""
    try:
        policy = json.loads(git(root, "show", f"{candidate}:{POLICY_PATH}"))
    except ValueError as exc:
        raise CandidateError(f"{POLICY_PATH} at {candidate} is not JSON: {exc}") from exc
    workflow = policy.get("workflow") if isinstance(policy, dict) else None
    if not isinstance(workflow, str) or not re.fullmatch(r"\.github/workflows/[\w.-]+\.ya?ml",
                                                         workflow):
        raise CandidateError(f"{POLICY_PATH} at {candidate} names no workflow file")
    return workflow


def select_runs(payload: Any, repository: str, workflow: str,
                candidate: str) -> list[dict[str, Any]]:
    """Successful push/dispatch runs of ``workflow`` on ``candidate``, newest first."""
    runs = payload.get("workflow_runs") if isinstance(payload, dict) else None
    if not isinstance(runs, list):
        raise CandidateError("the GitHub API returned no workflow_runs list")
    chosen: list[dict[str, Any]] = []
    for run in runs:
        if not isinstance(run, dict):
            continue
        head_repository = run.get("head_repository")
        path = run.get("path")
        if (type(run.get("id")) is int
                and run.get("head_sha") == candidate
                and isinstance(path, str)
                and (path == workflow or path.startswith(workflow + "@"))
                and run.get("event") in EVENTS
                and run.get("status") == "completed"
                and run.get("conclusion") == "success"
                and isinstance(head_repository, dict)
                and head_repository.get("full_name") == repository):
            chosen.append(run)
    return sorted(chosen, key=lambda run: int(run["id"]), reverse=True)


def list_runs(gh: Gh, repository: str, workflow: str, candidate: str) -> list[dict[str, Any]]:
    endpoint = (f"repos/{repository}/actions/workflows/{Path(workflow).name}/runs"
                f"?head_sha={candidate}&status=success&per_page=100")
    proc = gh(["api", endpoint])
    if proc.returncode:
        raise CandidateError("listing the Excel runs failed: "
                             + one_line(proc.stderr.decode("utf-8", "replace")))
    try:
        payload = json.loads(proc.stdout.decode("utf-8"))
    except ValueError as exc:
        raise CandidateError(f"the GitHub API returned invalid JSON: {exc}") from exc
    return select_runs(payload, repository, workflow, candidate)


def validate_run(gh: Gh, root: Path, repository: str, candidate: str, run_id: int,
                 directory: Path) -> str | None:
    """Why the run's artifact does not certify the candidate, or None if it does."""
    shutil.rmtree(directory, ignore_errors=True)
    directory.mkdir(parents=True)
    name = f"excel-vba-regression-{run_id}"
    proc = gh(["run", "download", str(run_id), "--repo", repository, "--name", name,
               "--dir", str(directory)])
    if proc.returncode:
        return (f"artifact {name} could not be downloaded: "
                + one_line(proc.stderr.decode("utf-8", "replace")))
    record_path = directory / RECORD
    if not record_path.is_file():
        return f"artifact {name} holds no {RECORD}"
    check = subprocess.run(
        [sys.executable, str(root / CHECKER), "--record", str(record_path),
         "--candidate-sha", candidate, "--evidence-ref", candidate,
         "--require-pass", "--log-directory", str(directory)],
        cwd=root, capture_output=True, text=True, check=False)
    if check.returncode:
        detail = [line.strip(" -") for line in check.stdout.splitlines()[1:]]
        return "record rejected: " + one_line(" ".join(detail) or check.stderr)
    # The checker has validated the record's schema.
    runner = json.loads(record_path.read_text(encoding="utf-8-sig"))["runner"]
    named = runner["workflow"]["run_id"] if isinstance(runner["workflow"], dict) else None
    if runner["class"] != RUNNER_CLASS or named != run_id:
        return (f"record names runner class {runner['class']} and run {named}, "
                f"not {RUNNER_CLASS} run {run_id}")
    return None


@dataclass
class Outcome:
    candidate: str
    workflow: str
    run_id: int | None = None
    reasons: list[str] = field(default_factory=list)

    @property
    def status(self) -> str:
        return CERTIFIED if self.run_id is not None else UNCERTIFIED


def certify(gh: Gh, root: Path, repository: str, workflow: str, candidate: str,
            work_dir: Path) -> Outcome:
    outcome = Outcome(candidate, workflow)
    try:
        runs = list_runs(gh, repository, workflow, candidate)
    except CandidateError as exc:
        outcome.reasons.append(str(exc))
        return outcome
    if not runs:
        outcome.reasons.append(
            f"no successful push or workflow_dispatch run of {workflow} in {repository} "
            f"has head commit {candidate}")
    for run in runs:
        run_id = int(run["id"])
        reason = validate_run(gh, root, repository, candidate, run_id, work_dir / str(run_id))
        if reason is None:
            outcome.run_id = run_id
            break
        outcome.reasons.append(f"run {run_id}: {reason}")
    return outcome


def summary(outcome: Outcome, mode: str, run_url: str | None) -> str:
    lines = [f"Exact-candidate Excel certification of `{outcome.candidate}`: "
             f"**{outcome.status}**", ""]
    if outcome.status == CERTIFIED:
        lines.append(f"Validated the record of {outcome.workflow} run "
                     f"[{outcome.run_id}]({run_url}) with `--candidate-sha` and "
                     "`--evidence-ref` set to this commit, `--require-pass` and the "
                     "retained log.")
    else:
        lines += [f"- {reason}" for reason in outcome.reasons]
        lines.append("")
        if mode == "tag":
            lines.append("The release stops here. Run the Excel workflow on this exact "
                         "commit, then re-run this workflow.")
        else:
            lines.append("Rehearsal only: this run creates no attestation and no release, "
                         "and a tag push of this commit would fail here.")
    return "\n".join(lines) + "\n"


def append(path: str | None, text: str) -> None:
    if path:
        with open(path, "a", encoding="utf-8") as f:
            f.write(text)


def main(argv: Sequence[str] | None = None, gh: Gh = run_gh) -> int:
    ap = argparse.ArgumentParser(description=(__doc__ or "").split("\n")[0])
    ap.add_argument("--root", default=".", help="the checked-out repository (default: .)")
    ap.add_argument("--repository", required=True, help="owner/name whose runs to search")
    ap.add_argument("--candidate", required=True,
                    help="full SHA of the checked-out commit to certify")
    ap.add_argument("--mode", required=True, choices=("tag", "rehearsal"),
                    help="tag: fail when uncertified; rehearsal: report it as a warning")
    ap.add_argument("--work-dir", required=True, help="where run artifacts are downloaded")
    ap.add_argument("--keep-dir", default=None,
                    help="copy the certifying artifact (record and log) here")
    ap.add_argument("--server-url", default=os.environ.get("GITHUB_SERVER_URL",
                                                           "https://github.com"))
    ap.add_argument("--github-output", default=os.environ.get("GITHUB_OUTPUT"))
    ap.add_argument("--summary", default=os.environ.get("GITHUB_STEP_SUMMARY"))
    args = ap.parse_args(argv)
    root = Path(args.root).resolve()

    try:
        candidate = checked_out_candidate(root, args.candidate)
        workflow = excel_workflow(root, candidate)
    except CandidateError as exc:
        print(f"::error::{one_line(str(exc))}")
        return 2

    outcome = certify(gh, root, args.repository, workflow, candidate, Path(args.work_dir))
    run_url = None
    if outcome.run_id is not None:
        run_url = f"{args.server_url}/{args.repository}/actions/runs/{outcome.run_id}"
        if args.keep_dir:
            shutil.copytree(Path(args.work_dir) / str(outcome.run_id), args.keep_dir,
                            dirs_exist_ok=True)
    append(args.github_output,
           f"certification={outcome.status}\nexcel_run={outcome.run_id or ''}\n")
    text = summary(outcome, args.mode, run_url)
    append(args.summary, text)
    print(text, end="")
    if outcome.status == CERTIFIED:
        return 0
    level = "error" if args.mode == "tag" else "warning"
    print(f"::{level}::{candidate} has no valid Excel certification of its own: "
          + one_line("; ".join(outcome.reasons)))
    return 1 if args.mode == "tag" else 0


if __name__ == "__main__":
    sys.exit(main())
