"""Fixtures for the CI test-coverage guard (tools/check_ci_test_coverage.py)."""

from __future__ import annotations

import contextlib
import io
import subprocess
import sys
import tempfile
import textwrap
import unittest
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "tools"))

import check_ci_test_coverage as guard  # noqa: E402

UNIT = "import unittest\n\nclass T(unittest.TestCase):\n    def test_ok(self):\n        pass\n"
SCRIPT = "print('PASS')\n"


def workflow(steps: str, *, on: str = "pull_request:\n    branches: [main]",
             job_extra: str = "") -> str:
    body = textwrap.indent(textwrap.dedent(steps).strip("\n"), " " * 6)
    return (f"name: CI\non:\n  {on}\npermissions:\n  contents: read\njobs:\n"
            f"  check:\n    name: Check\n    runs-on: ubuntu-24.04\n{job_extra}"
            f"    steps:\n      - uses: actions/checkout@v0 # pinned in real workflows\n{body}\n")


def shim(*entries: str) -> str:
    rows = "".join(f"    ({entry!r},),\n" for entry in entries)
    return ("import subprocess, sys\ncommands = (\n" + rows + "    ('check_live.py',),\n)\n"
            "for command in commands:\n    subprocess.run([sys.executable, *command])\n")


class FixtureRepository:
    def __init__(self, files: dict[str, str]) -> None:
        self.tmp = tempfile.TemporaryDirectory()
        self.root = Path(self.tmp.name)
        self.git("init", "-q")
        self.write(files)

    def git(self, *args: str) -> None:
        subprocess.run(["git", "-C", str(self.root), *args], check=True,
                       stdout=subprocess.DEVNULL, stderr=subprocess.PIPE)

    def write(self, files: dict[str, str]) -> None:
        for path, text in files.items():
            target = self.root / path
            target.parent.mkdir(parents=True, exist_ok=True)
            target.write_text(text, encoding="utf-8")
        self.git("add", "-A")

    def check(self, **kwargs: Any) -> dict[str, Any]:
        kwargs.setdefault("metadata_only", lambda paths: False)
        return guard.run_check(self.root, **kwargs)


class YamlReaderTests(unittest.TestCase):
    def test_block_style_subset(self) -> None:
        data = guard.load_yaml(textwrap.dedent("""\
            on:
              pull_request:
                branches: [main, "release/**"]  # comment
            jobs:
              a:
                steps:
                  - name: literal
                    run: |
                      echo one
                      # kept: inside a block scalar

                      echo two
                  - run: >-
                      python x.py
                      --flag
                  - name: "quoted: value"
                    if: ${{ always() &&
                      !cancelled() }}
            """))
        steps = data["jobs"]["a"]["steps"]
        self.assertEqual(data["on"]["pull_request"]["branches"], ["main", "release/**"])
        self.assertEqual(steps[0]["run"], "echo one\n# kept: inside a block scalar\n\necho two\n")
        self.assertEqual(steps[1]["run"], "python x.py --flag")
        self.assertEqual(steps[2]["name"], "quoted: value")
        self.assertEqual(steps[2]["if"], "${{ always() && !cancelled() }}")

    def test_unsupported_or_ambiguous_yaml_is_an_error(self) -> None:
        for text in ("a: 1\na: 2\n", "a: &anchor 1\n", "a:\n  - b\n c: 1\n", "a:\n\tb: 1\n"):
            with self.subTest(text=text), self.assertRaises(guard.CoverageError):
                guard.load_yaml(text)

    def test_every_repository_workflow_parses(self) -> None:
        for path in sorted((ROOT / ".github" / "workflows").glob("*.y*ml")):
            with self.subTest(path=path.name):
                self.assertIsInstance(guard.load_yaml(path.read_text(encoding="utf-8")), dict)


class CoverageTests(unittest.TestCase):
    def repo(self, files: dict[str, str]) -> FixtureRepository:
        repository = FixtureRepository(files)
        self.addCleanup(repository.tmp.cleanup)
        return repository

    def findings(self, report: dict[str, Any]) -> dict[str, str]:
        return {item["path"]: item["message"] for item in report["findings"]}

    def test_referenced_file_passes_and_unreferenced_file_is_reported(self) -> None:
        repository = self.repo({
            ".github/workflows/ci.yml": workflow("""\
                - name: Run tests
                  working-directory: benchmark
                  run: python test_listed.py
                """),
            "benchmark/test_listed.py": SCRIPT,
        })
        self.assertEqual(repository.check()["status"], "pass")
        repository.write({"benchmark/test_unlisted.py": SCRIPT})
        report = repository.check()
        self.assertEqual(report["status"], "fail")
        self.assertEqual(list(self.findings(report)), ["benchmark/test_unlisted.py"])
        self.assertIn("no CI step or runner list", self.findings(report)["benchmark/test_unlisted.py"])

    def test_untracked_files_are_not_inventoried(self) -> None:
        repository = self.repo({".github/workflows/ci.yml": workflow("- run: echo none")})
        (repository.root / "test_scratch.py").write_text(SCRIPT, encoding="utf-8")
        self.assertEqual(repository.check()["status"], "pass")

    def test_path_filtered_or_push_only_workflows_give_no_credit(self) -> None:
        for on in ("pull_request:\n    paths: [tools/**]", "push:\n    branches: [main]",
                   "pull_request:\n    branches: [develop]",
                   "pull_request:\n    types: [labeled]"):
            with self.subTest(on=on):
                repository = self.repo({
                    ".github/workflows/ci.yml": workflow("- run: python test_a.py", on=on),
                    "test_a.py": SCRIPT,
                })
                self.assertEqual(list(self.findings(repository.check())), ["test_a.py"])

    def test_unittest_discovery_credits_testcase_modules_in_packages(self) -> None:
        repository = self.repo({
            ".github/workflows/ci.yml": workflow("- run: python -m unittest discover -s tests"),
            "tests/test_unit.py": UNIT,
            "tests/test_script_only.py": SCRIPT,
            "tests/nested/test_not_a_package.py": UNIT,
            "tests/pkg/__init__.py": "",
            "tests/pkg/test_in_package.py": UNIT,
        })
        findings = self.findings(repository.check())
        self.assertEqual(sorted(findings), ["tests/nested/test_not_a_package.py",
                                            "tests/test_script_only.py"])
        self.assertIn("defines no unittest.TestCase", findings["tests/test_script_only.py"])

    def test_unittest_pattern_and_named_modules(self) -> None:
        repository = self.repo({
            ".github/workflows/ci.yml": workflow("""\
                - run: python -m unittest discover -s tests -p 'test_a*.py'
                - run: python3 -m unittest tests.test_b tests/test_c.py
                """),
            "tests/test_a1.py": UNIT,
            "tests/test_b.py": UNIT,
            "tests/test_c.py": UNIT,
            "tests/test_d.py": UNIT,
        })
        self.assertEqual(list(self.findings(repository.check())), ["tests/test_d.py"])

    def test_runner_list_credits_entries_except_in_reduced_mode(self) -> None:
        files = {
            "benchmark/test_evidence_tools.py": shim("test_tool.py", "study/test_model.py"),
            "benchmark/test_tool.py": SCRIPT,
            "benchmark/study/test_model.py": SCRIPT,
        }
        full = self.repo({**files, ".github/workflows/ci.yml": workflow("""\
            - working-directory: benchmark
              run: python test_evidence_tools.py
            """)})
        self.assertEqual(full.check()["status"], "pass")
        reduced = self.repo({**files, ".github/workflows/ci.yml": workflow("""\
            - working-directory: benchmark
              run: python test_evidence_tools.py --checks-only
            """)})
        # The runner itself still runs; its fixture entries do not.
        self.assertEqual(sorted(self.findings(reduced.check())),
                         ["benchmark/study/test_model.py", "benchmark/test_tool.py"])

    def test_selected_runs_require_a_change_to_trigger_full_proofs(self) -> None:
        steps = """\
            - working-directory: benchmark
              env:
                FULL_PROOFS: ${{ needs.select.outputs.full }}
              run: |
                if [[ "$FULL_PROOFS" == "true" ]]; then
                  python test_evidence_tools.py
                else
                  python test_evidence_tools.py --checks-only
                fi
            """
        repository = self.repo({
            ".github/workflows/ci.yml": workflow(steps),
            "benchmark/test_evidence_tools.py": shim("test_tool.py"),
            "benchmark/test_tool.py": SCRIPT,
        })
        report = repository.check()
        self.assertEqual(report["status"], "pass")
        self.assertEqual({row["level"] for row in report["tests"]}, {"with full proofs"})
        report = repository.check(metadata_only=lambda paths: paths == ["benchmark/test_tool.py"])
        self.assertIn("metadata-only", self.findings(report)["benchmark/test_tool.py"])

    def test_selection_mapping_is_the_real_one(self) -> None:
        repository = self.repo({
            ".github/workflows/ci.yml": workflow(
                "- run: python tests/test_check_issues.py\n- run: python tests/test_other.py",
                job_extra="    if: needs.select.outputs.full == 'true'\n"),
            "tests/test_check_issues.py": SCRIPT,
            "tests/test_other.py": SCRIPT,
        })
        report = guard.run_check(repository.root)
        self.assertEqual(list(self.findings(report)), ["tests/test_check_issues.py"])

    def test_unrecognized_conditions_give_no_credit(self) -> None:
        cases = {
            "step if": "- if: github.event_name == 'push'\n  run: python test_a.py",
            "or-list": "- run: test -f x || python test_a.py",
            "shell if": "- run: |\n    if [[ -n \"$X\" ]]; then python test_a.py; fi",
            "unaggregated continue-on-error": "- continue-on-error: true\n  run: python test_a.py",
        }
        for name, steps in cases.items():
            with self.subTest(name):
                repository = self.repo({".github/workflows/ci.yml": workflow(steps),
                                        "test_a.py": SCRIPT})
                self.assertEqual(list(self.findings(repository.check())), ["test_a.py"])

    def test_test_run_only_when_full_proofs_are_off_is_reported(self) -> None:
        # Review finding: the else branch runs only on metadata-only changes,
        # so editing the test (which selects full proofs) would skip it.
        scripts = {
            "else branch": 'if [[ "$FULL_PROOFS" == "true" ]]; then\n  echo full\n'
                           "else\n  python test_tool.py\nfi",
            "elif branch": 'if [[ "$FULL_PROOFS" == "true" ]]; then\n  echo full\n'
                           "elif true; then\n  python test_tool.py\nfi",
            "negated test": 'if [[ "$FULL_PROOFS" != "true" ]]; then\n  python test_tool.py\nfi',
            "unproven test": 'if [[ -n "$FULL_PROOFS" ]]; then\n  python test_tool.py\nfi',
        }
        for name, script in scripts.items():
            with self.subTest(name):
                steps = ("- env:\n    FULL_PROOFS: ${{ needs.select.outputs.full }}\n  run: |\n"
                         + textwrap.indent(script, "    "))
                repository = self.repo({".github/workflows/ci.yml": workflow(steps),
                                        "test_tool.py": SCRIPT})
                self.assertEqual(list(self.findings(repository.check())), ["test_tool.py"])
        inline = ('- run: |\n    if [ "${{ needs.select.outputs.full }}" = true ]; then\n'
                  "      python test_tool.py\n    fi")
        repository = self.repo({".github/workflows/ci.yml": workflow(inline), "test_tool.py": SCRIPT})
        self.assertEqual(repository.check()["status"], "pass")

    def test_failure_masking_and_short_circuit_lists_give_no_credit(self) -> None:
        # Review findings: `python test_a.py || true` masks a failure, and
        # `[ ... ] && python test_a.py` may skip the test while the step passes.
        cases = [
            "python test_a.py || true",
            "python test_a.py || echo failed || true",
            '[ "$RUN_TESTS" = true ] && python test_a.py\necho done',
            "python test_a.py &&\n  echo ran",
            "python test_a.py | tee log.txt",
            "! python test_a.py",
            "python test_a.py &",
        ]
        for script in cases:
            with self.subTest(script=script):
                steps = "- run: |\n" + textwrap.indent(script, "    ")
                repository = self.repo({".github/workflows/ci.yml": workflow(steps),
                                        "test_a.py": SCRIPT})
                self.assertEqual(list(self.findings(repository.check())), ["test_a.py"])

    def test_outcome_must_be_enforced_not_merely_mentioned(self) -> None:
        # Review finding: echoing the outcome does not make a failure blocking.
        enforcers = [
            '- if: always()\n  run: echo "${{ steps.t.outcome }}"',
            '- if: always()\n  run: test "${{ steps.t.outcome }}" = failure',
            '- if: always()\n  run: test "${{ steps.t.outcome }}" = success || true',
            '- if: always()\n  continue-on-error: true\n  run: test "${{ steps.t.outcome }}" = success',
            '- if: github.event_name == \'push\'\n  run: test "${{ steps.t.outcome }}" = success',
            '- name: "outcome ${{ steps.t.outcome }}"\n  run: echo checked',
        ]
        for enforcer in enforcers:
            with self.subTest(enforcer=enforcer):
                steps = f"- id: t\n  continue-on-error: true\n  run: python test_a.py\n{enforcer}"
                repository = self.repo({".github/workflows/ci.yml": workflow(steps),
                                        "test_a.py": SCRIPT})
                self.assertEqual(list(self.findings(repository.check())), ["test_a.py"])
        steps = ("- id: t\n  continue-on-error: true\n  run: python test_a.py\n"
                 "- if: always()\n  env:\n    OUTCOME: ${{ steps.t.outcome }}\n"
                 '  run: test "$OUTCOME" = success')
        repository = self.repo({".github/workflows/ci.yml": workflow(steps), "test_a.py": SCRIPT})
        self.assertEqual(repository.check()["status"], "pass")

    def test_package_load_tests_hook_withholds_discovery_credit(self) -> None:
        # Review finding: a package hook replaces discovery below it.
        repository = self.repo({
            ".github/workflows/ci.yml": workflow("- run: python -m unittest discover -s tests"),
            "tests/pkg/__init__.py": "def load_tests(loader, tests, pattern):\n    return tests\n",
            "tests/pkg/test_a.py": UNIT,
            "tests/plain/__init__.py": "",
            "tests/plain/test_b.py": UNIT,
        })
        findings = self.findings(repository.check())
        self.assertEqual(list(findings), ["tests/pkg/test_a.py"])
        self.assertIn("load_tests hook in tests/pkg/__init__.py", findings["tests/pkg/test_a.py"])

    def test_recognized_conditions_and_aggregated_continue_on_error_give_credit(self) -> None:
        steps = """\
            - if: ${{ !cancelled() && steps.deps.outcome == 'success' }}
              run: python test_a.py
            - id: b
              continue-on-error: true
              run: |
                if [[ -n "$X" ]]; then echo prepared; fi
                python test_b.py
            - if: always()
              run: test "${{ steps.b.outcome }}" = success
            """
        repository = self.repo({".github/workflows/ci.yml": workflow(steps),
                                "test_a.py": SCRIPT, "test_b.py": SCRIPT})
        report = repository.check()
        self.assertEqual(report["status"], "pass", report["findings"])
        self.assertEqual({row["level"] for row in report["tests"]}, {"always"})

    def test_job_skipped_with_a_selected_dependency_is_selected(self) -> None:
        text = workflow("- run: echo select") + textwrap.indent(textwrap.dedent("""\
              study:
                needs: check
                if: needs.check.outputs.full == 'true'
                runs-on: ubuntu-24.04
                steps:
                  - run: python test_study.py
              after:
                needs: study
                runs-on: ubuntu-24.04
                steps:
                  - run: python test_after.py
            """), "  ")
        repository = self.repo({".github/workflows/ci.yml": text,
                                "test_study.py": SCRIPT, "test_after.py": SCRIPT})
        levels = {row["path"]: row["level"] for row in repository.check()["tests"]}
        self.assertEqual(levels, {"test_after.py": "with full proofs",
                                  "test_study.py": "with full proofs"})

    def test_stale_reference_and_exemptions_are_reported(self) -> None:
        repository = self.repo({
            ".github/workflows/ci.yml": workflow("- run: python test_a.py\n- run: python test_gone.py"),
            "test_a.py": SCRIPT,
            "test_b.py": SCRIPT,
        })
        report = repository.check(exempt={"test_a.py": "x", "test_b.py": "y", "test_c.py": "z"})
        findings = self.findings(report)
        self.assertEqual(sorted(findings), ["test_a.py", "test_c.py", "test_gone.py"])
        self.assertIn("remove the exemption", findings["test_a.py"])
        self.assertIn("untracked", findings["test_c.py"])
        self.assertIn("not tracked", findings["test_gone.py"])

    def test_unreadable_runner_list_is_an_error(self) -> None:
        repository = self.repo({
            ".github/workflows/ci.yml": workflow("- run: python benchmark/test_evidence_tools.py"),
            "benchmark/test_evidence_tools.py": "commands = [name for name in ()]\n",
        })
        with self.assertRaises(guard.CoverageError):
            repository.check()


class RepositoryLayoutTests(unittest.TestCase):
    """The real workflows and runner, with test files added after the fact."""

    def setUp(self) -> None:
        listed = subprocess.run(["git", "-C", str(ROOT), "ls-files", "-z"], check=True,
                                stdout=subprocess.PIPE).stdout.decode().split("\0")
        keep = [p for p in listed if p and (
            p.startswith(".github/workflows/") or p == "benchmark/test_evidence_tools.py"
            or guard.TEST_FILE.fullmatch(Path(p).name))]
        self.repository = FixtureRepository(
            {p: (ROOT / p).read_text(encoding="utf-8") for p in keep})
        self.addCleanup(self.repository.tmp.cleanup)

    def test_current_layout_passes(self) -> None:
        report = guard.run_check(self.repository.root)
        self.assertEqual(report["status"], "pass", report["findings"])

    def test_new_study_test_is_reported_until_the_runner_lists_it(self) -> None:
        new = "benchmark/normal_tail_study/test_vba_kernel_source.py"
        self.repository.write({new: SCRIPT})
        report = guard.run_check(self.repository.root)
        self.assertEqual([item["path"] for item in report["findings"]], [new])
        runner = self.repository.root / "benchmark/test_evidence_tools.py"
        text = runner.read_text(encoding="utf-8").replace(
            '    ("normal_tail_study/test_kernel_model.py",),\n',
            '    ("normal_tail_study/test_kernel_model.py",),\n'
            '    ("normal_tail_study/test_vba_kernel_source.py",),\n')
        self.repository.write({"benchmark/test_evidence_tools.py": text})
        report = guard.run_check(self.repository.root)
        self.assertEqual(report["status"], "pass", report["findings"])

    def test_new_tests_module_runs_through_discovery(self) -> None:
        self.repository.write({"tests/test_pr_milestones_workflow.py": UNIT})
        self.assertEqual(guard.run_check(self.repository.root)["status"], "pass")
        self.repository.write({"tests/test_script_style.py": SCRIPT})
        report = guard.run_check(self.repository.root)
        self.assertEqual([item["path"] for item in report["findings"]], ["tests/test_script_style.py"])


class CommandLineTests(unittest.TestCase):
    def test_exit_codes(self) -> None:
        repository = FixtureRepository({".github/workflows/ci.yml": workflow("- run: echo"),
                                        "test_orphan.py": SCRIPT})
        self.addCleanup(repository.tmp.cleanup)
        with contextlib.redirect_stdout(io.StringIO()) as out:
            self.assertEqual(guard.main(["--root", str(repository.root)]), 1)
        self.assertIn("`test_orphan.py`", out.getvalue())
        repository.write({".github/workflows/ci.yml": "on: [pull_request\n"})
        with contextlib.redirect_stderr(io.StringIO()) as err:
            self.assertEqual(guard.main(["--root", str(repository.root)]), 2)
        self.assertIn("ERROR", err.getvalue())


if __name__ == "__main__":
    unittest.main()
