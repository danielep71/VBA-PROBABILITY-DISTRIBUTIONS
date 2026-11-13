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

A pull request that changes VBA or the runner therefore runs the full suite
on every push, including its final head. The post-merge run on the release
branch and every export are push or dispatch events, so they are full too.

    python3 tools/select_excel_suite.py --event pull_request --base SHA [--head HEAD]
"""
from __future__ import annotations

import argparse
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
        return "full", f"{len(hits)} path(s) can change the Excel run, first {hits[0]}: full suite"
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
