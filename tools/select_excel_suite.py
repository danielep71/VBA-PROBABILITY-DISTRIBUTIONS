#!/usr/bin/env python3
"""Choose the Excel suite level for a pull request: fast or full (#299).

Two test procedures take almost the whole Excel regression, because they
drive the double-double Gamma kernel (ci timings, #298). The FAST level skips
them (.github/excel-fast-policy.json) and finishes in about a minute. It is a
smoke test of the runner and the harness, never certification evidence.

FAST is chosen only for a pull request whose complete diff cannot change what
Excel executes or how the run is driven. Everything else is FULL, and so is
anything the selection cannot classify:

  * any event other than pull_request (push, dispatch, schedule);
  * a missing, zero or non-ancestor base, or an empty diff;
  * a changed path that is a VBA module (*.bas, *.cls, *.frm) anywhere, lies
    under ci/, is an Excel policy (.github/excel-*.json), is the Excel
    workflow, or is this selection tool or its policy.

One exception keeps a pull request from repeating a run that already exists
(2026-11-15). When the head carries benchmark/excel_regression_record.json
for a candidate commit that

  * is an ancestor of the head,
  * passed its complete regression with the assertion count the evidence
    policy at the head expects, and
  * differs from the head in no path that can change the Excel run,

then what Excel would execute at the head is what the record certifies, and
the level is FAST: a smoke test of the runner instead of a second full run.
A record that is missing, malformed, failed, for another commit, or older
than a VBA, runner or policy change does not qualify, and the level is FULL.

The post-merge run on the release branch and every export are push or
dispatch events, so they are always full: a wrong record cannot keep a
change from the full suite, only postpone it to the merge.

    python3 tools/select_excel_suite.py --event pull_request --base SHA [--head HEAD]
"""
from __future__ import annotations

import argparse
import json
import os
import re
import subprocess

FULL_PREFIXES = ("ci/",)
FULL_SUFFIXES = (".bas", ".cls", ".frm")
FULL_FILES = frozenset({
    ".github/workflows/excel-vba-regression.yml",
    "tools/select_excel_suite.py",
})
EXCEL_POLICY = re.compile(r"^\.github/excel-[^/]+\.json$")
RECORD = "benchmark/excel_regression_record.json"
POLICY = ".github/excel-evidence-policy.json"
SHA = re.compile(r"[0-9a-f]{40}")


def needs_full(path: str) -> bool:
    return (path.startswith(FULL_PREFIXES) or path.lower().endswith(FULL_SUFFIXES)
            or path in FULL_FILES or bool(EXCEL_POLICY.match(path)))


def git(*args: str) -> str:
    return subprocess.check_output(["git", *args], text=True, stderr=subprocess.PIPE)


def changed_paths(base: str, head: str) -> list[str]:
    """Every path touched between base and head, both sides of a rename."""
    out = git("diff", "--name-status", "-M", f"{base}...{head}")
    paths: list[str] = []
    for line in out.splitlines():
        parts = line.split("\t")
        paths.extend(p for p in parts[1:] if p)
    return paths


def certified_candidate(head: str) -> tuple[str | None, str]:
    """The candidate the head's Excel record certifies for the head, or None and why not."""
    try:
        record = json.loads(git("show", f"{head}:{RECORD}"))
        policy = json.loads(git("show", f"{head}:{POLICY}"))
    except (subprocess.CalledProcessError, OSError, ValueError):
        return None, "no readable Excel record or policy at the head"
    if not isinstance(record, dict) or not isinstance(policy, dict):
        return None, "the Excel record or policy is not an object"
    candidate = record.get("candidate_sha")
    if not isinstance(candidate, str) or not SHA.fullmatch(candidate):
        return None, "the Excel record names no candidate commit"
    stages = record.get("stages")
    harness = record.get("harness")
    if not isinstance(stages, dict) or not isinstance(harness, dict):
        return None, "the Excel record has no stages or harness"
    for stage in ("import", "compile", "regression", "cleanup"):
        entry = stages.get(stage)
        if not isinstance(entry, dict) or entry.get("status") != "PASS":
            return None, f"the Excel record's {stage} stage is not PASS"
    expected = policy.get("expected_assertions")
    if type(expected) is not int or expected < 1:
        return None, "the evidence policy states no assertion count"
    if (harness.get("assertions"), harness.get("passed"), harness.get("failed")) != (expected, expected, 0):
        return None, "the Excel record's assertion count is not the policy's"
    try:
        subprocess.check_call(["git", "merge-base", "--is-ancestor", candidate, head],
                              stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        out = git("diff", "--name-status", "-M", candidate, head)
    except (subprocess.CalledProcessError, OSError):
        return None, "the recorded candidate is not an ancestor of the head"
    later = sorted(p for line in out.splitlines() for p in line.split("\t")[1:] if p and needs_full(p))
    if later:
        return None, f"{len(later)} path(s) changed after the recorded candidate, first {later[0]}"
    return candidate, "certified"


def select(event: str, base: str, head: str = "HEAD") -> tuple[str, str]:
    if event != "pull_request":
        return "full", f"event {event or 'unknown'}: full suite"
    if not re.fullmatch(r"[0-9a-fA-F]{40}", base or "") or set(base) == {"0"}:
        return "full", "missing comparison base: full suite"
    try:
        git("merge-base", base, head)
        paths = changed_paths(base, head)
    except (subprocess.CalledProcessError, OSError):
        return "full", "comparison base is not reachable: full suite"
    if not paths:
        return "full", "empty change set: full suite"
    hits = sorted(p for p in paths if needs_full(p))
    if hits:
        candidate, why = certified_candidate(head)
        if candidate is not None:
            return "fast", (f"{len(hits)} path(s) can change the Excel run, and the Excel record "
                            f"certifies them at {candidate[:7]} with nothing later: fast suite")
        return "full", (f"{len(hits)} path(s) can change the Excel run, first {hits[0]}; "
                        f"{why}: full suite")
    return "fast", f"{len(paths)} changed path(s), none can change the Excel run: fast suite"


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    ap.add_argument("--event", required=True)
    ap.add_argument("--base", default="")
    ap.add_argument("--head", default="HEAD")
    a = ap.parse_args(argv)
    suite, reason = select(a.event, a.base, a.head)
    print(f"{suite}: {reason}")
    output = os.environ.get("GITHUB_OUTPUT")
    if output:
        with open(output, "a", encoding="utf-8") as f:
            f.write(f"suite={suite}\n")
    summary = os.environ.get("GITHUB_STEP_SUMMARY")
    if summary:
        with open(summary, "a", encoding="utf-8") as f:
            f.write(f"Excel suite level: **{suite}** ({reason})\n")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
