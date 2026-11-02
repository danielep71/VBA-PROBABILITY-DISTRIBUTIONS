#!/usr/bin/env python3
"""Run every static check of the repository with one command, locally and in CI.

Every check runs, even when an earlier one fails, so one run reports every
problem. The run writes, under test-results/static-checks/:

- one <check>.log per check, with its command and output;
- report.json: the commit, whether the working tree is dirty, the pull request
  head, the Python version and platform, the SHA-256 of every tracked file, and
  each check's command, status and exit code;
- summary.md, which --summary also appends to a file such as
  $GITHUB_STEP_SUMMARY.

The checks keep writing their own evidence (test-results/*.json and *.md) as
before. A check whose tool is not installed (Ruff, actionlint) is skipped and
reported as such; --strict, which CI passes, makes it fail instead. The exit
code is 1 when any check failed.

The run covers the repository only: it compiles no VBA, runs no Excel and
certifies no figure.

To add a check, add one Check entry to the list returned by checks().
"""

from __future__ import annotations

import argparse
import datetime
import hashlib
import importlib.util
import json
import os
import platform
import shutil
import subprocess
import sys
from collections.abc import Callable, Sequence
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parents[1]
RESULTS_DIR = "test-results"
REPORT_DIR = f"{RESULTS_DIR}/static-checks"
SCOPE = "Repository checks only: no VBA compilation, no Excel run, no numerical certification"
TIMEOUT_SECONDS = 600
STATUSES = ("pass", "fail", "skip")


@dataclass(frozen=True)
class Context:
    """What the checks need to know about the run."""

    root: Path
    # Candidate revision and optional comparison base of the committed-range checks.
    head: str = "HEAD"
    base: str | None = None
    pull_request_head: str | None = None
    # Explicit actionlint executable; without one, actionlint is looked up on PATH.
    actionlint: str | None = None


@dataclass(frozen=True)
class Check:
    """One command, run from the repository root.

    missing: why the check cannot run here (its tool is not installed), if it cannot.
    output:  a path, relative to the root, that also receives the command's output.
    summary: a Markdown file the command writes; it is added to the run summary.
    reports: other files the command writes, such as its JSON report.

    Every path in output, summary and reports is deleted before the command runs,
    so a file left by an earlier run can never stand in for this run's evidence.
    """

    name: str
    command: list[str]
    missing: str | None = None
    output: str | None = None
    summary: str | None = None
    reports: tuple[str, ...] = ()


@dataclass
class Result:
    name: str
    status: str
    command: list[str]
    exit_code: int | None
    detail: str = ""
    log: str = field(default="", repr=False)


# ---------------------------------------------------------------------------
# The list of checks.

def module_missing(name: str, hint: str) -> str | None:
    """Why the Python module `name` cannot run, or None if it is installed."""
    if importlib.util.find_spec(name) is not None:
        return None
    return f"{name} is not installed; {hint}"


def checks(context: Context) -> list[Check]:
    """Every static check, in the order it runs. Add a new check here."""
    python = sys.executable
    results = RESULTS_DIR
    whitespace = [python, "tools/check_committed_whitespace.py", "--root", ".",
                  "--mode", "committed", "--head", context.head,
                  "--output", f"{results}/committed-whitespace.json",
                  "--summary", f"{results}/committed-whitespace.md"]
    if context.base:
        whitespace += ["--base", context.base]
    # An explicit actionlint must exist there: CI pins it and never falls back to PATH.
    actionlint = shutil.which(context.actionlint or "actionlint")
    actionlint_missing = None if actionlint else (
        f"actionlint is not an executable at {context.actionlint}" if context.actionlint
        else "actionlint is not on PATH; install it or pass --actionlint")
    return [
        Check("whitespace-self-test",
              [python, "tools/check_committed_whitespace.py", "--root", ".", "--self-test"]),
        Check("whitespace", whitespace, summary=f"{results}/committed-whitespace.md",
              reports=(f"{results}/committed-whitespace.json",)),
        Check("vba-jumps-self-test",
              [python, "tools/check_vba_jumps.py", "--root", ".", "--self-test"]),
        Check("vba-jumps",
              [python, "tools/check_vba_jumps.py", "--root", ".",
               "--output", f"{results}/vba-jumps.json", "--summary", f"{results}/vba-jumps.md"],
              summary=f"{results}/vba-jumps.md", reports=(f"{results}/vba-jumps.json",)),
        # VBA source rules; their fixtures are in tests/test_check_vba.py (python-tests).
        Check("vba-source", [python, "tools/check_vba_source.py", "--root", "."]),
        Check("vba-conditionals", [python, "tools/check_vba_conditionals.py", "--root", "."]),
        Check("vba-visibility", [python, "tools/check_vba_visibility.py", "--root", "."]),
        Check("release-semantics-self-test",
              [python, "tools/check_release_semantics.py", "--root", ".", "--self-test"]),
        Check("release-semantics",
              [python, "tools/check_release_semantics.py", "--root", ".",
               "--output", f"{results}/release-semantics.json",
               "--summary", f"{results}/release-semantics.md"],
              summary=f"{results}/release-semantics.md",
              reports=(f"{results}/release-semantics.json",)),
        # Tracked binaries: approved assets/ images, or a reviewed SHA-256 in the allowlist.
        Check("binaries", [python, "tools/check_binaries.py", "--root", "."]),
        Check("release-bundle-self-test",
              [python, ".github/scripts/build_release_bundle.py", "--self-test"]),
        Check("python-tests", [python, "-m", "unittest", "discover", "-s", "tests", "-v"]),
        # Rules come from ruff.toml: syntax errors and pyflakes only.
        Check("ruff",
              [python, "-m", "ruff", "check", "--no-cache", "--output-format", "concise", "."],
              missing=module_missing("ruff", "install it with "
                                     "python -m pip install -r tools/requirements-lint-ci.txt"),
              output=f"{results}/ruff.txt"),
        # -config-file makes the self-hosted runner label declaration explicit.
        Check("actionlint",
              [actionlint or context.actionlint or "actionlint", "-no-color",
               "-config-file", ".github/actionlint.yaml"],
              missing=actionlint_missing,
              output=f"{results}/actionlint.txt"),
    ]


# ---------------------------------------------------------------------------
# Running, reporting.

def git(root: Path, *args: str) -> str:
    return subprocess.run(["git", *args], cwd=root, capture_output=True, text=True,
                          check=True).stdout


def run_check(check: Check, root: Path, strict: bool) -> Result:
    # A stale file from an earlier run must not pass for this run's evidence.
    for stale in (check.output, check.summary, *check.reports):
        if stale:
            (root / stale).unlink(missing_ok=True)
    if check.missing:
        status = "fail" if strict else "skip"
        return Result(check.name, status, check.command, None, check.missing, check.missing)
    try:
        completed = subprocess.run(check.command, cwd=root, capture_output=True, text=True,
                                   timeout=TIMEOUT_SECONDS, check=False)
    except (OSError, subprocess.TimeoutExpired) as error:
        return Result(check.name, "fail", check.command, None,
                      f"could not run: {type(error).__name__}", str(error))
    output, code = completed.stdout + completed.stderr, completed.returncode
    if check.output:
        path = root / check.output
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(output, encoding="utf-8")
    return Result(check.name, "pass" if code == 0 else "fail", check.command, code,
                  "" if code == 0 else f"exit code {code}", output)


def tracked_hashes(root: Path) -> dict[str, str]:
    """SHA-256 of the working-tree bytes of every tracked file."""
    names = [name for name in git(root, "ls-files", "-z").split("\0") if name]
    return {name: hashlib.sha256((root / name).read_bytes()).hexdigest()
            for name in names if (root / name).is_file()}


def build_report(context: Context, results: Sequence[Result], strict: bool,
                 dirty: bool) -> dict[str, Any]:
    root = context.root
    return {
        "schema_version": 1,
        "tool": "tools/run_static_checks.py",
        "status": "fail" if any(r.status == "fail" for r in results) else "pass",
        "strict": strict,
        "commit": git(root, "rev-parse", "HEAD").strip(),
        "dirty": dirty,
        "candidate": context.head,
        "base": context.base,
        "pull_request_head": context.pull_request_head,
        "generated_at": datetime.datetime.now(datetime.timezone.utc)
        .isoformat(timespec="seconds"),
        "python": platform.python_version(),
        "python_implementation": platform.python_implementation(),
        "platform": platform.platform(),
        "scope": SCOPE,
        "checks": [{"name": r.name, "status": r.status, "command": r.command,
                    "exit_code": r.exit_code, "detail": r.detail, "log": f"{r.name}.log"}
                   for r in results],
        "files_sha256": tracked_hashes(root),
    }


def summary(report: dict[str, Any], details: Sequence[str] = ()) -> str:
    marks = {"pass": "✅ pass", "fail": "❌ fail", "skip": "⏭️ skip"}
    head = report["pull_request_head"]
    lines = [f"## Static checks: {report['status'].upper()}", "",
             f"- Commit: `{report['commit']}`"
             + (" (dirty working tree)" if report["dirty"] else ""),
             f"- Pull request head: `{head}`" if head else "- Pull request head: none",
             f"- Python {report['python']} on {report['platform']}"
             + ("; strict: a missing tool fails" if report["strict"]
                else "; not strict: a missing tool is skipped"),
             f"- Scope: {report['scope']}.", "",
             "| Check | Result | Exit code | Detail |", "| --- | --- | --- | --- |"]
    lines += [f"| `{check['name']}` | {marks[check['status']]} | "
              f"{'' if check['exit_code'] is None else check['exit_code']} | "
              f"{check['detail'].replace('|', '/')} |" for check in report["checks"]]
    lines += ["", f"Logs and `report.json`, with the SHA-256 of every tracked file, are in "
              f"`{REPORT_DIR}/`; CI keeps them as the run's artifact."]
    text = "\n".join(lines) + "\n"
    for detail in details:
        text += "\n" + detail.rstrip("\n") + "\n"
    return text


def run(context: Context, listed: Sequence[Check], strict: bool,
        summary_path: Path | None = None) -> int:
    root = context.root
    out = root / REPORT_DIR
    out.mkdir(parents=True, exist_ok=True)
    for stale in out.glob("*.log"):
        stale.unlink()
    dirty = bool(git(root, "status", "--porcelain"))
    results: list[Result] = []
    details: list[str] = []
    in_ci = os.environ.get("GITHUB_ACTIONS") == "true"
    for check in listed:
        result = run_check(check, root, strict)
        results.append(result)
        (out / f"{check.name}.log").write_text(f"$ {' '.join(check.command)}\n{result.log}",
                                               encoding="utf-8")
        print(f"{check.name}: {result.status.upper()}"
              + (f" ({result.detail})" if result.detail else ""), flush=True)
        if result.status == "fail":
            if result.log:
                print(result.log.rstrip(), flush=True)
            if in_ci:
                print(f"::error title=Static check {check.name}::{result.detail}", flush=True)
        if check.summary and result.status != "skip" and (root / check.summary).is_file():
            details.append((root / check.summary).read_text(encoding="utf-8"))
    report = build_report(context, results, strict, dirty)
    (out / "report.json").write_text(json.dumps(report, indent=2) + "\n", encoding="utf-8")
    text = summary(report, details)
    (out / "summary.md").write_text(text, encoding="utf-8")
    if summary_path is not None:
        with summary_path.open("a", encoding="utf-8") as handle:
            handle.write(text)
    counts = {status: sum(r.status == status for r in results) for status in STATUSES}
    print(f"static checks: {report['status'].upper()} ({counts['pass']} passed, "
          f"{counts['fail']} failed, {counts['skip']} skipped; report in {REPORT_DIR}/)")
    if counts["skip"]:
        print("Skipped checks did not run; install their tools, or pass --strict to fail them.")
    return 1 if report["status"] == "fail" else 0


def optional(text: str | None) -> str | None:
    return text or None


def main(argv: Sequence[str] | None = None,
         listed: Callable[[Context], list[Check]] = checks) -> int:
    parser = argparse.ArgumentParser(description=(__doc__ or "").splitlines()[0])
    parser.add_argument("--root", type=Path, default=ROOT)
    parser.add_argument("--head", default=os.environ.get("CANDIDATE_SHA") or "HEAD",
                        type=lambda text: text or "HEAD",
                        help="candidate revision of the committed-whitespace check "
                             "(default: $CANDIDATE_SHA, else HEAD)")
    parser.add_argument("--base", default=os.environ.get("BASE_SHA"), type=optional,
                        help="comparison base of the committed-whitespace check "
                             "(default: $BASE_SHA, else none)")
    parser.add_argument("--pr-head", default=os.environ.get("PR_HEAD_SHA"), type=optional,
                        help="pull request head recorded in the report "
                             "(default: $PR_HEAD_SHA, else none)")
    parser.add_argument("--actionlint", default=os.environ.get("ACTIONLINT"), type=optional,
                        help="actionlint executable (default: $ACTIONLINT, else PATH)")
    parser.add_argument("--strict", action="store_true",
                        help="fail a check whose tool is missing instead of skipping it, "
                             "as CI does")
    parser.add_argument("--summary", type=Path, help="also append the summary to this file")
    args = parser.parse_args(argv)
    context = Context(args.root.resolve(), args.head, optional(args.base),
                      optional(args.pr_head), optional(args.actionlint))
    return run(context, listed(context), args.strict, args.summary)


if __name__ == "__main__":
    sys.exit(main())
