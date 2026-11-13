"""Fixtures for the Excel suite levels (#299): tools/select_excel_suite.py,
.github/excel-fast-policy.json and their wiring in the runner and workflow."""

from __future__ import annotations

import json
import re
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "tools"))

import select_excel_suite as S  # noqa: E402


def git(root: Path, *args: str) -> str:
    return subprocess.run(
        ["git", "-C", str(root), "-c", "user.name=fixture", "-c", "user.email=f@example.invalid",
         "-c", "commit.gpgsign=false", *args],
        check=True, capture_output=True, text=True).stdout.strip()


class Selection(unittest.TestCase):
    def setUp(self) -> None:
        self.tmp = tempfile.TemporaryDirectory()
        self.root = Path(self.tmp.name)
        git(self.root, "init", "-q")
        for rel in ("src/M_A.bas", "tests/M_T.bas", "tests/test_x.py", "docs/A.md",
                    "ci/Run.ps1", "benchmark/tool.py", ".github/excel-evidence-policy.json"):
            path = self.root / rel
            path.parent.mkdir(parents=True, exist_ok=True)
            path.write_text("base\n", encoding="utf-8")
        git(self.root, "add", "-A")
        git(self.root, "commit", "-qm", "base")
        self.base = git(self.root, "rev-parse", "HEAD")

    def tearDown(self) -> None:
        self.tmp.cleanup()

    def level(self, changes: dict[str, str | None], event: str = "pull_request",
              base: str | None = None) -> str:
        for rel, text in changes.items():
            path = self.root / rel
            if text is None:
                path.unlink()
            else:
                path.parent.mkdir(parents=True, exist_ok=True)
                path.write_text(text, encoding="utf-8")
        git(self.root, "add", "-A")
        git(self.root, "commit", "-qm", "change", "--allow-empty")
        import os
        cwd = os.getcwd()
        os.chdir(self.root)
        try:
            return S.select(event, self.base if base is None else base)[0]
        finally:
            os.chdir(cwd)

    def test_documentation_and_python_only_is_fast(self) -> None:
        self.assertEqual(self.level({"docs/A.md": "new\n", "benchmark/tool.py": "new\n",
                                     "tests/test_x.py": "new\n"}), "fast")

    def test_production_module_is_full(self) -> None:
        self.assertEqual(self.level({"src/M_A.bas": "new\n", "docs/A.md": "new\n"}), "full")

    def test_test_module_is_full(self) -> None:
        self.assertEqual(self.level({"tests/M_T.bas": "new\n"}), "full")

    def test_a_vba_module_anywhere_is_full(self) -> None:
        self.assertEqual(self.level({"benchmark/study/M_PROBE.bas": "new\n"}), "full")
        self.assertEqual(self.level({"benchmark/study/C_THING.CLS": "new\n"}), "full")

    def test_runner_policy_workflow_and_selector_are_full(self) -> None:
        for rel in ("ci/Run.ps1", ".github/excel-evidence-policy.json",
                    ".github/excel-fast-policy.json",
                    ".github/workflows/excel-vba-regression.yml",
                    "tools/select_excel_suite.py"):
            with self.subTest(rel=rel):
                self.assertEqual(self.level({rel: "changed again\n" + rel}), "full")

    def test_deleted_and_renamed_modules_are_full(self) -> None:
        self.assertEqual(self.level({"src/M_A.bas": None}), "full")
        git(self.root, "mv", "tests/M_T.bas", "docs/M_T.txt")
        self.assertEqual(self.level({}), "full")

    def test_other_events_and_unusable_bases_are_full(self) -> None:
        change = {"docs/A.md": "new\n"}
        self.assertEqual(self.level(change, event="push"), "full")
        self.assertEqual(self.level({}, event="workflow_dispatch"), "full")
        self.assertEqual(self.level({}, event=""), "full")
        self.assertEqual(self.level({}, base=""), "full")
        self.assertEqual(self.level({}, base="0" * 40), "full")
        self.assertEqual(self.level({}, base="f" * 40), "full")
        self.assertEqual(self.level({}, base="main"), "full")

    def test_empty_change_set_is_full(self) -> None:
        import os
        cwd = os.getcwd()
        os.chdir(self.root)
        try:
            self.assertEqual(S.select("pull_request", self.base)[0], "full")
        finally:
            os.chdir(cwd)


class Wiring(unittest.TestCase):
    def setUp(self) -> None:
        self.fast = json.loads((ROOT / ".github/excel-fast-policy.json").read_text(encoding="utf-8"))
        self.full = json.loads((ROOT / ".github/excel-evidence-policy.json").read_text(encoding="utf-8"))
        self.runner = (ROOT / "ci/Run-ExcelVbaTests.ps1").read_text(encoding="utf-8")
        self.workflow = (ROOT / ".github/workflows/excel-vba-regression.yml").read_text(encoding="utf-8")

    def test_each_skipped_test_is_called_exactly_once(self) -> None:
        module = (ROOT / "tests/M_STATS_PROBDIST_TEST.bas").read_text(encoding="ascii").splitlines()
        self.assertTrue(self.fast["skip_tests"])
        self.assertEqual(len(self.fast["skip_tests"]), len(set(self.fast["skip_tests"])))
        for name in self.fast["skip_tests"]:
            calls = [line for line in module if line.strip() == name]
            self.assertEqual(len(calls), 1, name)
            self.assertTrue(any(re.match(rf"Private Sub {name}\(\)", line) for line in module), name)

    def test_fast_count_is_a_strict_subset(self) -> None:
        self.assertEqual(set(self.fast), {"schema_version", "skip_tests", "expected_assertions"})
        self.assertGreater(self.fast["expected_assertions"], 0)
        self.assertLess(self.fast["expected_assertions"], self.full["expected_assertions"])

    def test_runner_implements_the_fast_level(self) -> None:
        self.assertIn("[switch]$Fast", self.runner)
        self.assertIn("excel-fast-policy.json", self.runner)
        self.assertIn("expected exactly one call of $testName", self.runner)
        self.assertIn("if ($Fast) { $expectedAssertions = [int]$fastPolicy.expected_assertions }",
                      self.runner)
        self.assertIn("FAST level, not certification evidence", self.runner)
        self.assertIn("if ($Fast) { $RunStudies = $false }", self.runner)

    def test_workflow_passes_fast_only_from_the_selection_or_a_plain_dispatch(self) -> None:
        self.assertIn("python3 tools/select_excel_suite.py --event", self.workflow)
        self.assertIn("needs: select", self.workflow)
        self.assertIn("!cancelled()", self.workflow)
        self.assertEqual(self.workflow.count("&& '-Fast' || '' }}"), 1)
        fast = [self.workflow.split("&& '-Fast' || '' }}")[0].rsplit("${{", 1)[1]]
        self.assertIn("github.event_name == 'pull_request' && needs.select.outputs.suite == 'fast'",
                      fast[0])
        self.assertIn("inputs.suite == 'fast' && !(inputs.export_main", fast[0])
        self.assertNotIn("'push'", fast[0])


if __name__ == "__main__":
    unittest.main()
