"""Tests for tools/run_static_checks.py, in temporary Git repositories.

The run is tested with small checks injected in place of the real list, so the
test suite never runs itself.
"""

from __future__ import annotations

import contextlib
import hashlib
import io
import json
import os
import platform
import shutil
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path
from typing import Any
from unittest import mock

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "tools"))

import run_static_checks as R  # noqa: E402
from run_static_checks import Check, Context  # noqa: E402

PASSING = [sys.executable, "-c", "print('fine')"]
FAILING = [sys.executable, "-c", "import sys; sys.exit('bad input')"]


@unittest.skipIf(shutil.which("git") is None, "git is not installed")
class RepositoryTest(unittest.TestCase):
    def setUp(self) -> None:
        directory = tempfile.TemporaryDirectory(prefix="static-checks-")
        self.addCleanup(directory.cleanup)
        self.root = Path(directory.name)
        self.git("init", "-q")
        self.git("config", "user.name", "Fixture")
        self.git("config", "user.email", "fixture@example.invalid")
        self.git("config", "core.autocrlf", "false")
        self.git("config", "commit.gpgsign", "false")
        self.out = self.root / R.REPORT_DIR

    def git(self, *args: str) -> str:
        return subprocess.run(["git", *args], cwd=self.root, capture_output=True, text=True,
                              check=True).stdout

    def write(self, files: dict[str, str | bytes]) -> None:
        for name, content in files.items():
            path = self.root / name
            path.parent.mkdir(parents=True, exist_ok=True)
            if isinstance(content, bytes):
                path.write_bytes(content)
            else:
                path.write_text(content, encoding="utf-8")

    def commit(self, files: dict[str, str | bytes] | None = None) -> None:
        self.write(files or {})
        self.git("add", "-A")
        self.git("commit", "-qm", "Fixture", "--allow-empty")

    def run_checks(self, listed: list[Check], strict: bool = False,
                   summary: Path | None = None,
                   context: Context | None = None) -> tuple[int, str, dict[str, Any]]:
        with contextlib.redirect_stdout(io.StringIO()) as printed, \
                mock.patch.dict(os.environ, {"GITHUB_ACTIONS": ""}):
            code = R.run(context or Context(self.root), listed, strict, summary)
        report = json.loads((self.out / "report.json").read_text(encoding="utf-8"))
        return code, printed.getvalue(), report

    @staticmethod
    def statuses(report: dict[str, Any]) -> dict[str, tuple[str, int | None]]:
        return {check["name"]: (check["status"], check["exit_code"])
                for check in report["checks"]}


class RunTest(RepositoryTest):
    def test_passing_check(self) -> None:
        self.commit({"README.md": "x\n"})
        code, printed, report = self.run_checks([Check("ok", PASSING)])
        self.assertEqual(code, 0)
        self.assertIn("ok: PASS", printed)
        self.assertEqual(report["status"], "pass")
        self.assertEqual(self.statuses(report), {"ok": ("pass", 0)})
        self.assertEqual((self.out / "ok.log").read_text(encoding="utf-8"),
                         f"$ {' '.join(PASSING)}\nfine\n")

    def test_failing_check_fails_the_run_and_later_checks_still_run(self) -> None:
        self.commit({"README.md": "x\n"})
        listed = [Check("broken", FAILING), Check("after", PASSING)]
        code, printed, report = self.run_checks(listed)
        self.assertEqual(code, 1)
        self.assertIn("broken: FAIL (exit code 1)", printed)
        self.assertIn("bad input", printed)
        self.assertIn("after: PASS", printed)
        self.assertEqual(report["status"], "fail")
        self.assertEqual(self.statuses(report), {"broken": ("fail", 1), "after": ("pass", 0)})
        self.assertIn("bad input", (self.out / "broken.log").read_text(encoding="utf-8"))

    def test_command_that_cannot_start_fails(self) -> None:
        self.commit()
        code, printed, report = self.run_checks(
            [Check("absent", [str(self.root / "no-such-tool")])])
        self.assertEqual(code, 1)
        self.assertEqual(self.statuses(report), {"absent": ("fail", None)})
        self.assertIn("could not run", printed)

    def test_missing_tool_is_skipped_unless_strict(self) -> None:
        self.commit()
        listed = [Check("needs-tool", ["tool"], missing="tool is not on PATH"),
                  Check("ok", PASSING)]
        code, printed, report = self.run_checks(listed)
        self.assertEqual(code, 0)
        self.assertIn("needs-tool: SKIP (tool is not on PATH)", printed)
        self.assertIn("pass --strict to fail them", printed)
        self.assertEqual(report["status"], "pass")
        self.assertIs(report["strict"], False)
        self.assertEqual(self.statuses(report), {"needs-tool": ("skip", None),
                                                 "ok": ("pass", 0)})
        self.assertIn("tool is not on PATH", (self.out / "needs-tool.log").read_text("utf-8"))

    def test_strict_mode_fails_a_missing_tool(self) -> None:
        self.commit()
        listed = [Check("needs-tool", ["tool"], missing="tool is not on PATH"),
                  Check("ok", PASSING)]
        code, printed, report = self.run_checks(listed, strict=True)
        self.assertEqual(code, 1)
        self.assertIn("needs-tool: FAIL (tool is not on PATH)", printed)
        self.assertEqual(report["status"], "fail")
        self.assertIs(report["strict"], True)
        self.assertEqual(self.statuses(report), {"needs-tool": ("fail", None),
                                                 "ok": ("pass", 0)})

    def test_report_records_the_run_and_every_tracked_file(self) -> None:
        tracked: dict[str, str | bytes] = {"README.md": "x\n", "src/module.bas": b"A\r\nB\r\n",
                                           "data.bin": b"\x00\x01binary"}
        self.commit(tracked)
        self.write({"untracked.txt": "y\n"})
        context = Context(self.root, head="HEAD", base="0" * 40, pull_request_head="f" * 40)
        listed = [Check("ok", PASSING), Check("broken", FAILING)]
        _, _, report = self.run_checks(listed, context=context)
        self.assertEqual(report["schema_version"], 1)
        self.assertEqual(report["commit"], self.git("rev-parse", "HEAD").strip())
        self.assertIs(report["dirty"], True)
        self.assertEqual(report["pull_request_head"], "f" * 40)
        self.assertEqual((report["candidate"], report["base"]), ("HEAD", "0" * 40))
        self.assertEqual(report["python"], platform.python_version())
        self.assertEqual(report["platform"], platform.platform())
        expected = {name: hashlib.sha256(content.encode("utf-8") if isinstance(content, str)
                                         else content).hexdigest()
                    for name, content in tracked.items()}
        self.assertEqual(report["files_sha256"], expected)
        self.assertEqual(report["checks"][1],
                         {"name": "broken", "status": "fail", "command": FAILING,
                          "exit_code": 1, "detail": "exit code 1", "log": "broken.log"})

    def test_clean_tree_and_no_pull_request(self) -> None:
        self.commit({"README.md": "x\n"})
        _, _, report = self.run_checks([Check("ok", PASSING)])
        self.assertIs(report["dirty"], False)
        self.assertIsNone(report["pull_request_head"])
        self.assertIn("- Pull request head: none",
                      (self.out / "summary.md").read_text(encoding="utf-8"))

    def test_summary_output_and_check_evidence(self) -> None:
        self.commit()
        writes_md = [sys.executable, "-c",
                     "import pathlib; pathlib.Path('evidence.md').write_text('## Evidence\\n')"]
        listed = [Check("ok", PASSING, output="results/ok.txt"),
                  Check("evidence", writes_md, summary="evidence.md"),
                  Check("broken", FAILING, summary="stale.md"),
                  Check("needs-tool", ["tool"], missing="tool is not on PATH",
                        output="results/tool.txt")]
        self.write({"stale.md": "## Stale\n", "results/tool.txt": "stale\n"})
        step_summary = self.root / "step-summary.md"
        step_summary.write_text("before\n", encoding="utf-8")
        self.run_checks(listed, summary=step_summary)
        self.assertEqual((self.root / "results/ok.txt").read_text(encoding="utf-8"), "fine\n")
        # Files from an earlier run are removed, not reported as this run's evidence.
        self.assertFalse((self.root / "stale.md").exists())
        self.assertFalse((self.root / "results/tool.txt").exists())
        text = step_summary.read_text(encoding="utf-8")
        self.assertTrue(text.startswith("before\n## Static checks: FAIL"))
        self.assertIn("| `ok` | ✅ pass | 0 |  |", text)
        self.assertIn("| `broken` | ❌ fail | 1 | exit code 1 |", text)
        self.assertIn("| `needs-tool` | ⏭️ skip |  | tool is not on PATH |", text)
        self.assertIn("\n## Evidence\n", text)
        self.assertNotIn("Stale", text)
        self.assertIn("no Excel run", text)
        self.assertEqual((self.out / "summary.md").read_text(encoding="utf-8"),
                         text.removeprefix("before\n"))

    def test_stale_json_report_is_removed_when_the_checker_fails_first(self) -> None:
        # A checker that fails before writing its JSON report must not leave the
        # passing report of an earlier run in place.
        self.commit()
        self.write({"results/check.json": '{"status": "pass"}\n'})
        listed = [Check("broken", FAILING, summary="results/check.md",
                        reports=("results/check.json",))]
        code, _, report = self.run_checks(listed)
        self.assertEqual(code, 1)
        self.assertEqual(self.statuses(report), {"broken": ("fail", 1)})
        self.assertFalse((self.root / "results/check.json").exists())

    def test_every_json_report_of_the_check_list_is_tracked(self) -> None:
        context = Context(self.root, head="HEAD")
        for check in R.checks(context):
            for argument, value in zip(check.command, check.command[1:]):
                if argument == "--output" and value.endswith(".json"):
                    self.assertIn(value, check.reports, check.name)

    def test_logs_of_an_earlier_run_are_removed(self) -> None:
        self.commit()
        self.run_checks([Check("old", PASSING)])
        self.run_checks([Check("new", PASSING)])
        self.assertEqual(sorted(path.name for path in self.out.glob("*.log")), ["new.log"])

    def test_github_annotation_for_a_failed_check(self) -> None:
        self.commit()
        with contextlib.redirect_stdout(io.StringIO()) as printed, \
                mock.patch.dict(os.environ, {"GITHUB_ACTIONS": "true"}):
            R.run(Context(self.root), [Check("broken", FAILING)], False)
        self.assertIn("::error title=Static check broken::exit code 1", printed.getvalue())


class MainTest(RepositoryTest):
    def call(self, argv: list[str], environment: dict[str, str]) -> tuple[int, Context]:
        seen: list[Context] = []

        def listed(context: Context) -> list[Check]:
            seen.append(context)
            return [Check("ok", PASSING)]

        names = ("CANDIDATE_SHA", "BASE_SHA", "PR_HEAD_SHA", "ACTIONLINT", "GITHUB_ACTIONS")
        cleared = {name: "" for name in names}
        with contextlib.redirect_stdout(io.StringIO()), \
                mock.patch.dict(os.environ, {**cleared, **environment}):
            code = R.main(["--root", str(self.root), *argv], listed)
        return code, seen[0]

    def test_defaults_without_ci_environment(self) -> None:
        self.commit()
        code, context = self.call([], {})
        self.assertEqual(code, 0)
        self.assertEqual((context.root, context.head, context.base, context.pull_request_head,
                          context.actionlint), (self.root.resolve(), "HEAD", None, None, None))

    def test_ci_environment_and_arguments(self) -> None:
        self.commit()
        _, context = self.call([], {"CANDIDATE_SHA": "a" * 40, "BASE_SHA": "b" * 40,
                                    "PR_HEAD_SHA": "c" * 40, "ACTIONLINT": "/opt/actionlint"})
        self.assertEqual((context.head, context.base, context.pull_request_head,
                          context.actionlint), ("a" * 40, "b" * 40, "c" * 40, "/opt/actionlint"))
        _, context = self.call(["--head", "d" * 40, "--base", "", "--pr-head", "e" * 40],
                               {"BASE_SHA": "b" * 40})
        self.assertEqual((context.head, context.base, context.pull_request_head),
                         ("d" * 40, None, "e" * 40))

    def test_strict_flag_reaches_the_run(self) -> None:
        self.commit()

        def listed(context: Context) -> list[Check]:
            return [Check("needs-tool", ["tool"], missing="tool is not on PATH")]

        with contextlib.redirect_stdout(io.StringIO()), \
                mock.patch.dict(os.environ, {"GITHUB_ACTIONS": ""}):
            self.assertEqual(R.main(["--root", str(self.root)], listed), 0)
            self.assertEqual(R.main(["--root", str(self.root), "--strict"], listed), 1)


class CheckListTest(unittest.TestCase):
    """The real list: names, the arguments CI relies on, and missing tools."""

    def by_name(self, context: Context) -> dict[str, Check]:
        listed = R.checks(context)
        names = [check.name for check in listed]
        self.assertEqual(len(names), len(set(names)), "check names must be unique")
        return {check.name: check for check in listed}

    def test_every_static_check_is_listed(self) -> None:
        checks = self.by_name(Context(ROOT))
        self.assertEqual(list(checks), [
            "whitespace-self-test", "whitespace", "vba-jumps-self-test", "vba-jumps",
            "vba-source", "vba-conditionals", "vba-visibility", "release-semantics-self-test",
            "release-semantics", "binaries", "release-bundle-self-test", "python-tests", "ruff",
            "mypy", "actionlint", "actionlint-probe"])
        for check in checks.values():
            script = next((part for part in check.command if part.endswith(".py")), None)
            if script is not None:
                self.assertTrue((ROOT / script).is_file(), script)
        self.assertEqual(checks["python-tests"].command[1:],
                         ["-m", "unittest", "discover", "-s", "tests", "-v"])
        self.assertIn("--self-test", checks["release-bundle-self-test"].command)
        self.assertEqual(checks["vba-jumps"].command[-4:],
                         ["--output", "test-results/vba-jumps.json",
                          "--summary", "test-results/vba-jumps.md"])

    def test_committed_whitespace_takes_head_and_base(self) -> None:
        command = self.by_name(Context(ROOT))["whitespace"].command
        self.assertEqual(command[command.index("--head") + 1], "HEAD")
        self.assertNotIn("--base", command)
        command = self.by_name(Context(ROOT, head="a" * 40, base="b" * 40))["whitespace"].command
        self.assertEqual(command[command.index("--head") + 1], "a" * 40)
        self.assertEqual(command[-2:], ["--base", "b" * 40])

    def test_missing_tools(self) -> None:
        with mock.patch("importlib.util.find_spec", return_value=None), \
                mock.patch("shutil.which", return_value=None):
            checks = self.by_name(Context(ROOT))
        self.assertIn("ruff is not installed", str(checks["ruff"].missing))
        self.assertIn("mypy is not installed", str(checks["mypy"].missing))
        self.assertIn("actionlint is not on PATH", str(checks["actionlint"].missing))
        self.assertEqual(checks["actionlint-probe"].missing, checks["actionlint"].missing)
        missing = {name for name, check in checks.items() if check.missing}
        self.assertEqual(missing, {"ruff", "mypy", "actionlint", "actionlint-probe"})

    def test_mypy_reads_its_configuration(self) -> None:
        # Scope and strictness live in mypy.ini, so the command passes no paths.
        self.assertEqual(self.by_name(Context(ROOT))["mypy"].command[1:], ["-m", "mypy"])
        self.assertTrue((ROOT / "mypy.ini").is_file())

    def test_explicit_actionlint(self) -> None:
        with tempfile.TemporaryDirectory(prefix="static-checks-tool-") as directory:
            path = Path(directory) / "actionlint"
            path.write_text("#!/bin/sh\n", encoding="utf-8")
            path.chmod(0o755)
            with mock.patch.dict(os.environ, {"PATH": ""}):
                check = self.by_name(Context(ROOT, actionlint=str(path)))["actionlint"]
            self.assertIsNone(check.missing)
            self.assertEqual(check.command, [str(path), "-no-color",
                                             "-config-file", ".github/actionlint.yaml"])
            # The probe runs exactly the command of the actionlint check.
            with mock.patch.dict(os.environ, {"PATH": ""}):
                probe = self.by_name(Context(ROOT, actionlint=str(path)))["actionlint-probe"]
            self.assertIsNone(probe.missing)
            self.assertEqual(probe.command[1:3], ["tools/probe_actionlint.py", "--"])
            self.assertEqual(probe.command[3:], check.command)
            # A pinned path that is absent is missing, even with actionlint on PATH.
            absent = str(Path(directory) / "absent" / "actionlint")
            with mock.patch.dict(os.environ, {"PATH": directory}):
                check = self.by_name(Context(ROOT, actionlint=absent))["actionlint"]
            self.assertEqual(check.missing, f"actionlint is not an executable at {absent}")
            self.assertEqual(check.command[0], absent)

if __name__ == "__main__":
    unittest.main()
