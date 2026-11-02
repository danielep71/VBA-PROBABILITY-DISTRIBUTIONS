#!/usr/bin/env python3
"""Negative controls for check_module_lists.py and the derived module inventory."""
from __future__ import annotations

import contextlib
import io
import json
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
import _tracked_modules as T  # noqa: E402
import check_module_lists as M  # noqa: E402

FAILS: list[str] = []
SOURCES = ["src/A.bas", "src/B.bas", "tests/T.bas"]


def check(condition: bool, message: str) -> None:
    if not condition:
        FAILS.append(message)


def git(root: Path, *args: str) -> None:
    subprocess.run(["git", "-C", str(root), *args], check=True, capture_output=True)


def write(root: Path, rel: str, text: str = 'Attribute VB_Name = "M"\n') -> Path:
    path = root / rel
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(text, encoding="utf-8")
    return path


def write_policy(root: Path, sources: list[str]) -> None:
    write(root, M.POLICY_PATH, json.dumps({"schema_version": 1, "regression_sources": sources}, indent=2) + "\n")


def run(root: Path) -> tuple[int, str]:
    output = io.StringIO()
    with contextlib.redirect_stdout(output):
        rc = M.main(["--root", str(root)])
    return rc, output.getvalue()


def expect(root: Path, label: str, passes: bool, needle: str = "") -> None:
    rc, output = run(root)
    if passes:
        check(rc == 0, f"{label} must pass, got: {output.strip()}")
    else:
        check(rc != 0, f"{label} must fail")
        check(needle in output, f"{label} must report {needle!r}, got: {output.strip()}")


def make_repo() -> Path:
    root = Path(tempfile.mkdtemp())
    git(root, "init", "-q")
    for rel in SOURCES:
        write(root, rel)
    # Neither a non-.bas file nor a module outside src/ and tests/ is a regression source.
    write(root, "tests/test_helper.py", "pass\n")
    write(root, "benchmark/Exporter.bas")
    write_policy(root, SOURCES)
    git(root, "add", "-A")
    return root


root = make_repo()
try:
    expect(root, "baseline", True)
    check(T.regression_modules(root) == SOURCES, "regression modules must be the tracked src/tests .bas files")
    check(T.production_modules(root) == ["src/A.bas", "src/B.bas"], "production modules must be the tracked src .bas files")

    write_policy(root, list(reversed(SOURCES)))
    expect(root, "same list in another order", True)

    # A tracked module missing from the list.
    write_policy(root, ["src/A.bas", "tests/T.bas"])
    expect(root, "tracked src module missing from the list", False,
           "tracked module src/B.bas is missing from regression_sources")
    write_policy(root, ["src/A.bas", "src/B.bas"])
    expect(root, "tracked tests module missing from the list", False,
           "tracked module tests/T.bas is missing from regression_sources")

    # A tracked module in a nested directory is still in scope.
    write(root, "src/sub/N.bas")
    git(root, "add", "--", "src/sub/N.bas")
    write_policy(root, SOURCES)
    expect(root, "tracked nested module missing from the list", False,
           "tracked module src/sub/N.bas is missing from regression_sources")
    git(root, "rm", "-q", "--cached", "--", "src/sub/N.bas")
    shutil.rmtree(root / "src" / "sub")
    expect(root, "nested module removed again", True)

    # An extra listed module that is not tracked.
    write_policy(root, SOURCES + ["src/Ghost.bas"])
    expect(root, "listed module that does not exist", False,
           "regression_sources lists src/Ghost.bas, which is not a tracked .bas module")
    write(root, "src/Ghost.bas")
    expect(root, "listed module present but untracked", False, "(present but not tracked)")
    (root / "src" / "Ghost.bas").unlink()

    # A tracked module outside src/ and tests/ must not be listed.
    write_policy(root, SOURCES + ["benchmark/Exporter.bas"])
    expect(root, "listed module outside src/ and tests/", False,
           "regression_sources lists benchmark/Exporter.bas")

    # A module present on disk but untracked escapes both lists unless reported.
    write_policy(root, SOURCES)
    write(root, "tests/Scratch.bas")
    check("tests/Scratch.bas" not in T.regression_modules(root), "untracked module must not be picked up")
    expect(root, "unlisted module present but untracked", False,
           "tests/Scratch.bas is present but not tracked")
    git(root, "add", "--", "tests/Scratch.bas")
    expect(root, "newly tracked module that is not listed", False,
           "tracked module tests/Scratch.bas is missing from regression_sources")
    git(root, "rm", "-q", "--cached", "--", "tests/Scratch.bas")
    (root / "tests" / "Scratch.bas").unlink()

    # Upper-case extensions are modules too.
    write(root, "src/Upper.BAS")
    git(root, "add", "--", "src/Upper.BAS")
    expect(root, "tracked .BAS module missing from the list", False,
           "tracked module src/Upper.BAS is missing from regression_sources")
    git(root, "rm", "-q", "--cached", "--", "src/Upper.BAS")
    (root / "src" / "Upper.BAS").unlink()

    # Malformed lists.
    write_policy(root, SOURCES + ["src/A.bas"])
    expect(root, "duplicate listed module", False, "lists src/A.bas more than once")
    write_policy(root, [])
    expect(root, "empty list", False, "regression_sources must be a non-empty array")
    write(root, M.POLICY_PATH, "{not json")
    expect(root, "unreadable policy", False, "cannot read")
    write_policy(root, SOURCES)
    expect(root, "baseline restored", True)

    # Without Git the tracked set is unknown: fail closed, never glob.
    no_git = Path(tempfile.mkdtemp())
    try:
        for rel in SOURCES + [M.POLICY_PATH]:
            write(no_git, rel, (root / rel).read_text(encoding="utf-8"))
        expect(no_git, "root that is not a Git work tree", False, "cannot list the tracked modules")
        shutil.copytree(no_git, root / "nested")
        expect(root / "nested", "root below the Git top level", False, "not the top level")
        shutil.rmtree(root / "nested")
    finally:
        shutil.rmtree(no_git, ignore_errors=True)
finally:
    shutil.rmtree(root, ignore_errors=True)

if FAILS:
    print("FAIL: module-list fixtures")
    for failure in FAILS:
        print("  - " + failure)
    raise SystemExit(1)
print("PASS: module-list fixtures (a tracked module missing from regression_sources, a listed "
      "module that is absent, untracked or outside src/ and tests/, a present-but-untracked "
      "module, duplicates, a malformed policy and a non-Git root fail; reordering passes)")
