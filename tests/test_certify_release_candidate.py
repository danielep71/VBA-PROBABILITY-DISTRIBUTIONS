"""Tests for exact-commit release certification (#76).

`.github/scripts/certify_release_candidate.py` must certify a release commit
only from a green Excel run of that very commit. These tests run it against a
throwaway Git repository and an in-memory GitHub (runs and artifacts), with the
repository's own `benchmark/check_excel_certification.py`, and check that
`.github/workflows/release.yml` runs it on the tag and rehearsal paths.

Standard library and Git only. Run from the repository root:

    python -m unittest discover -s tests -p 'test_certify_release_candidate.py'
"""

from __future__ import annotations

import copy
import hashlib
import importlib.util
import io
import json
import re
import shutil
import subprocess
import sys
import tempfile
import unittest
from collections.abc import Sequence
from contextlib import redirect_stdout
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parents[1]
WORKFLOW = ROOT / ".github/workflows/release.yml"
SPEC = importlib.util.spec_from_file_location(
    "certify_release_candidate", ROOT / ".github/scripts/certify_release_candidate.py")
assert SPEC is not None and SPEC.loader is not None
certify = importlib.util.module_from_spec(SPEC)
sys.modules[SPEC.name] = certify
SPEC.loader.exec_module(certify)

REPOSITORY = "owner/repo"
EXCEL_WORKFLOW = ".github/workflows/excel.yml"
POLICY = {
    "schema_version": 1,
    "repository": REPOSITORY,
    "workflow": EXCEL_WORKFLOW,
    "entry_point": "RunAll_CI",
    "expected_assertions": 3,
    "regression_sources": ["src/A.bas", "tests/T.bas"],
    "grid_exporters": {"benchmark/main.csv": "benchmark/MainExport.bas"},
}
LOG = b"RESULT: ALL TESTS PASSED\r\n"


def git(root: Path, *args: str) -> str:
    return subprocess.run(["git", *args], cwd=root, capture_output=True, text=True,
                          check=True).stdout.strip()


def sha256(raw: bytes) -> str:
    return "sha256:" + hashlib.sha256(raw).hexdigest()


def record_for(root: Path, candidate: str, run_id: int) -> dict[str, Any]:
    """A green regression record, as the Excel workflow writes it for ``candidate``."""
    sources = [{"path": path, "sha256": sha256((root / path).read_bytes())}
               for path in sorted(POLICY["regression_sources"])]
    return {
        "schema_version": 1,
        "repository": REPOSITORY,
        "candidate_sha": candidate,
        "execution": "automated",
        "started_at": "2026-11-03T10:00:00Z",
        "finished_at": "2026-11-03T10:00:09Z",
        "runner": {"class": "self-hosted-excel", "identity": "PROBDIST",
                   "workflow": {"repository": REPOSITORY, "path": EXCEL_WORKFLOW,
                                "sha": candidate, "run_id": run_id, "run_attempt": 1}},
        "environment": {"excel_version": "16.0", "excel_build": "20228",
                        "office_bitness": "64-bit", "windows": "Windows 11 10.0.26200",
                        "os_architecture": "X64", "automation_security": "1 (isolated)",
                        "vba_project_access": "preconfigured runner"},
        "sources": sources,
        "stages": {name: {"status": "PASS", "detail": "ok"}
                   for name in ("import", "compile", "regression", "cleanup")},
        "harness": {"entry_point": "RunAll_CI", "assertions": 3, "passed": 3, "failed": 0},
        "log": {"path": "test-result.txt", "sha256": sha256(LOG)},
        "grids": {"benchmark/main.csv": {"exported": False, "sha256": None,
                                         "row_count": None, "exported_utc": None}},
    }


def run_entry(run_id: int, head_sha: str, **changes: Any) -> dict[str, Any]:
    entry: dict[str, Any] = {
        "id": run_id, "head_sha": head_sha, "path": EXCEL_WORKFLOW,
        "event": "workflow_dispatch", "status": "completed", "conclusion": "success",
        "head_repository": {"full_name": REPOSITORY},
    }
    entry.update(changes)
    return entry


class FakeGitHub:
    """The two gh calls the script makes: list runs, download one artifact."""

    def __init__(self) -> None:
        self.runs: list[dict[str, Any]] = []
        self.artifacts: dict[int, dict[str, bytes]] = {}
        self.calls: list[list[str]] = []

    def add(self, run: dict[str, Any], record: dict[str, Any] | None = None) -> None:
        self.runs.append(run)
        if record is not None:
            self.artifacts[run["id"]] = {
                "excel-certification.json":
                    b"\xef\xbb\xbf" + json.dumps(record, indent=2).encode(),
                "test-result.txt": LOG,
            }

    def __call__(self, args: Sequence[str]) -> subprocess.CompletedProcess[bytes]:
        args = list(args)
        self.calls.append(args)
        if args[0] == "api":
            match = re.search(r"head_sha=([0-9a-f]{40})", args[1])
            assert match, args
            # GitHub filters by head_sha; the script must not rely on it alone.
            body = {"total_count": len(self.runs), "workflow_runs": self.runs}
            return subprocess.CompletedProcess(args, 0, json.dumps(body).encode(), b"")
        assert args[:2] == ["run", "download"], args
        run_id = int(args[2])
        name = args[args.index("--name") + 1]
        files = self.artifacts.get(run_id)
        if files is None or name != f"excel-vba-regression-{run_id}":
            return subprocess.CompletedProcess(args, 1, b"",
                                               b"no valid artifacts found to download")
        directory = Path(args[args.index("--dir") + 1])
        for file_name, raw in files.items():
            (directory / file_name).write_bytes(raw)
        return subprocess.CompletedProcess(args, 0, b"", b"")

    def downloads(self) -> list[int]:
        return [int(call[2]) for call in self.calls if call[:2] == ["run", "download"]]


@unittest.skipIf(shutil.which("git") is None, "git is not installed")
class CertifyReleaseCandidateTest(unittest.TestCase):
    def setUp(self) -> None:
        directory = tempfile.TemporaryDirectory(prefix="release-certification-")
        self.addCleanup(directory.cleanup)
        self.tmp = Path(directory.name)
        self.root = self.tmp / "repo"
        self.root.mkdir()
        git(self.root, "init", "-q")
        git(self.root, "config", "user.email", "t@t")
        git(self.root, "config", "user.name", "t")
        git(self.root, "config", "commit.gpgsign", "false")
        files = {
            ".github/excel-evidence-policy.json": json.dumps(POLICY, indent=2) + "\n",
            "src/A.bas": 'Attribute VB_Name = "A"\n',
            "tests/T.bas": 'Attribute VB_Name = "T"\n',
            "benchmark/MainExport.bas": 'Attribute VB_Name = "MainExport"\n',
            "benchmark/main.csv": "function,value\nA,1\n",
            "README.md": "first\n",
        }
        for path, text in files.items():
            (self.root / path).parent.mkdir(parents=True, exist_ok=True)
            (self.root / path).write_text(text, encoding="utf-8", newline="\n")
        for name in ("check_excel_certification.py", "excel_certification.py"):
            shutil.copy2(ROOT / "benchmark" / name, self.root / "benchmark" / name)
        self.certified = self.commit("certified candidate")
        # A documentation-only follow-up: every numerical source byte is unchanged.
        (self.root / "README.md").write_text("second\n", encoding="utf-8")
        self.docs_only = self.commit("documentation only")
        self.assertEqual(git(self.root, "diff", "--name-only", self.certified,
                             self.docs_only), "README.md")
        git(self.root, "checkout", "-q", "--detach", self.certified)
        self.gh = FakeGitHub()

    def commit(self, message: str) -> str:
        git(self.root, "add", "-A")
        git(self.root, "commit", "-qm", message)
        return git(self.root, "rev-parse", "HEAD")

    def checkout(self, sha: str) -> None:
        git(self.root, "checkout", "-q", "--detach", sha)

    def main(self, candidate: str, mode: str = "tag") -> tuple[int, dict[str, str], str]:
        output = self.tmp / "github-output"
        summary = self.tmp / "summary.md"
        for path in (output, summary):
            path.write_text("", encoding="utf-8")
        args = ["--root", str(self.root), "--repository", REPOSITORY,
                "--candidate", candidate, "--mode", mode,
                "--work-dir", str(self.tmp / "runs"), "--keep-dir", str(self.tmp / "kept"),
                "--server-url", "https://github.example",
                "--github-output", str(output), "--summary", str(summary)]
        with redirect_stdout(io.StringIO()) as printed:
            code = certify.main(args, gh=self.gh)
        outputs = dict(line.split("=", 1)
                       for line in output.read_text(encoding="utf-8").splitlines())
        return code, outputs, summary.read_text(encoding="utf-8") + printed.getvalue()

    def checker(self, record: dict[str, Any], *extra: str) -> int:
        """The repository's checker on ``record``, as the old release step ran it."""
        path = self.tmp / "record.json"
        path.write_text(json.dumps(record), encoding="utf-8")
        return subprocess.run(
            [sys.executable, str(self.root / "benchmark/check_excel_certification.py"),
             "--record", str(path), "--require-pass", *extra],
            cwd=self.root, capture_output=True, check=False).returncode

    # Positive fixture: the matching candidate.

    def test_green_run_of_the_exact_commit_certifies_it(self) -> None:
        self.gh.add(run_entry(101, self.certified), record_for(self.root, self.certified, 101))
        code, outputs, text = self.main(self.certified)
        self.assertEqual(code, 0, text)
        self.assertEqual(outputs, {"certification": "CERTIFIED", "excel_run": "101"})
        self.assertIn(f"`{self.certified}`: **CERTIFIED**", text)
        self.assertIn("https://github.example/owner/repo/actions/runs/101", text)
        self.assertTrue((self.tmp / "kept/excel-certification.json").is_file())
        self.assertEqual((self.tmp / "kept/test-result.txt").read_bytes(), LOG)
        listing = self.gh.calls[0]
        self.assertEqual(listing[0], "api")
        self.assertIn(f"actions/workflows/excel.yml/runs?head_sha={self.certified}", listing[1])

    def test_rehearsal_of_the_exact_commit_is_also_certified(self) -> None:
        self.gh.add(run_entry(101, self.certified), record_for(self.root, self.certified, 101))
        code, outputs, _ = self.main(self.certified, mode="rehearsal")
        self.assertEqual((code, outputs["certification"]), (0, "CERTIFIED"))

    # Negative fixtures: a different commit with identical numerical source bytes.

    def test_later_commit_is_not_certified_by_the_earlier_run(self) -> None:
        self.gh.add(run_entry(101, self.certified), record_for(self.root, self.certified, 101))
        self.checkout(self.docs_only)
        code, outputs, text = self.main(self.docs_only)
        self.assertEqual(code, 1)
        self.assertEqual(outputs, {"certification": "UNCERTIFIED", "excel_run": ""})
        self.assertIn(f"has head commit {self.docs_only}", text)
        self.assertIn("The release stops here.", text)
        self.assertIn("::error::", text)
        self.assertEqual(self.gh.downloads(), [])

    def test_record_of_the_earlier_commit_is_rejected_for_the_later_one(self) -> None:
        record = record_for(self.root, self.certified, 102)
        # The bypass #76 closes: without --candidate-sha the earlier record
        # validates at the later commit, because its VBA bytes are identical.
        self.checkout(self.docs_only)
        self.assertEqual(self.checker(record, "--evidence-ref", self.docs_only), 0)
        self.assertEqual(self.checker(record, "--evidence-ref", self.docs_only,
                                      "--candidate-sha", self.docs_only), 1)
        # A run reported on the later commit whose record names the earlier one.
        self.gh.add(run_entry(102, self.docs_only), record)
        code, outputs, text = self.main(self.docs_only)
        self.assertEqual((code, outputs["certification"]), (1, "UNCERTIFIED"))
        self.assertIn(f"record candidate {self.certified} differs from required candidate "
                      f"{self.docs_only}", text)

    def test_rehearsal_reports_uncertified_without_failing(self) -> None:
        self.gh.add(run_entry(101, self.certified), record_for(self.root, self.certified, 101))
        self.checkout(self.docs_only)
        code, outputs, text = self.main(self.docs_only, mode="rehearsal")
        self.assertEqual(code, 0)
        self.assertEqual(outputs["certification"], "UNCERTIFIED")
        self.assertIn("**UNCERTIFIED**", text)
        self.assertIn("Rehearsal only: this run creates no attestation and no release", text)
        self.assertIn("::warning::", text)

    # Run selection and record binding.

    def test_runs_that_did_not_execute_the_commit_are_ignored(self) -> None:
        record = record_for(self.root, self.certified, 0)
        ignored = [
            run_entry(201, self.certified, event="pull_request"),
            run_entry(202, self.certified, head_repository={"full_name": "fork/repo"}),
            run_entry(203, self.certified, conclusion="failure"),
            run_entry(204, self.certified, status="in_progress", conclusion=None),
            run_entry(205, self.certified, path=".github/workflows/other.yml"),
            run_entry(206, self.docs_only),
        ]
        for run in ignored:
            changed = copy.deepcopy(record)
            changed["runner"]["workflow"]["run_id"] = run["id"]
            self.gh.add(run, changed)
        code, outputs, _ = self.main(self.certified)
        self.assertEqual((code, outputs["certification"]), (1, "UNCERTIFIED"))
        self.assertEqual(self.gh.downloads(), [])

    def test_expired_newer_artifact_falls_back_to_an_older_run(self) -> None:
        self.gh.add(run_entry(101, self.certified), record_for(self.root, self.certified, 101))
        self.gh.add(run_entry(105, self.certified))  # its artifact has expired
        code, outputs, _ = self.main(self.certified)
        self.assertEqual((code, outputs["excel_run"]), (0, "101"))
        self.assertEqual(self.gh.downloads(), [105, 101])

    def test_record_must_name_the_run_it_came_from(self) -> None:
        self.gh.add(run_entry(101, self.certified), record_for(self.root, self.certified, 999))
        code, outputs, text = self.main(self.certified)
        self.assertEqual((code, outputs["certification"]), (1, "UNCERTIFIED"))
        self.assertIn("run 999, not self-hosted-excel run 101", text)

    def test_red_record_does_not_certify(self) -> None:
        record = record_for(self.root, self.certified, 101)
        record["stages"]["cleanup"] = {"status": "FAIL", "detail": "workbook left open"}
        self.gh.add(run_entry(101, self.certified), record)
        code, outputs, text = self.main(self.certified)
        self.assertEqual((code, outputs["certification"]), (1, "UNCERTIFIED"))
        self.assertIn("green certification required; cleanup=FAIL", text)

    def test_changed_log_does_not_certify(self) -> None:
        self.gh.add(run_entry(101, self.certified), record_for(self.root, self.certified, 101))
        self.gh.artifacts[101]["test-result.txt"] = b"edited\r\n"
        code, _, text = self.main(self.certified)
        self.assertEqual(code, 1)
        self.assertIn("retained log digest does not match record", text)

    def test_candidate_must_be_the_checked_out_full_commit(self) -> None:
        self.gh.add(run_entry(101, self.certified), record_for(self.root, self.certified, 101))
        for candidate in (self.docs_only, self.certified[:12], "HEAD"):
            with self.subTest(candidate=candidate):
                code, outputs, _ = self.main(candidate)
                self.assertEqual((code, outputs), (2, {}))
        self.assertEqual(self.gh.calls, [])


def step(text: str, name: str) -> str:
    """The text of the workflow step named ``name``, up to the next step."""
    match = re.search(rf"^(\s*)- name: {re.escape(name)}\n(.*?)(?=^\1- name: |^\S|\Z)",
                      text, re.M | re.S)
    if not match:
        raise AssertionError(f"no step named {name!r}")
    return match.group(2)


def job(text: str, name: str) -> str:
    match = re.search(rf"^  {re.escape(name)}:\n(.*?)(?=^  \w[\w-]*:\n|\Z)", text, re.M | re.S)
    if not match:
        raise AssertionError(f"no job named {name!r}")
    return match.group(1)


class ReleaseWorkflowTest(unittest.TestCase):
    """release.yml binds every path to the checked-out commit."""

    text = WORKFLOW.read_text(encoding="utf-8")

    def test_identity_uses_the_peeled_tag_commit(self) -> None:
        identity = step(self.text, "Establish the release identity")
        self.assertIn("candidate=\"$(git rev-parse --verify 'HEAD^{commit}')\"", identity)
        self.assertIn('tagged="$(git rev-parse --verify "refs/tags/$TAG^{commit}")"', identity)
        self.assertIn('if [[ "$tagged" != "$candidate" ]]; then', identity)
        self.assertIn('git merge-base --is-ancestor "$candidate" upstream/main', identity)
        self.assertIn("candidate=%s", identity)

    def test_exact_candidate_certification_runs_on_every_path(self) -> None:
        verify = job(self.text, "verify")
        certification = step(verify, "Certify the exact candidate commit in Excel")
        self.assertNotIn("if:", certification)
        self.assertNotIn("continue-on-error", certification)
        self.assertIn("CANDIDATE: ${{ steps.identity.outputs.candidate }}", certification)
        self.assertIn('--candidate "$CANDIDATE"', certification)
        self.assertIn("mode=tag\n", certification)
        self.assertIn('if [[ "$REHEARSAL" == "true" ]]; then mode=rehearsal; fi',
                      certification)
        self.assertIn("certification: ${{ steps.certification.outputs.certification }}",
                      verify)
        self.assertRegex(verify, r"permissions:\n\s+contents: read\n(\s+#.*\n)*\s+actions: read\n")

    def test_retained_record_is_checked_at_the_candidate(self) -> None:
        retained = step(self.text, "Validate the retained Excel record against the candidate")
        self.assertIn('--evidence-ref "$CANDIDATE" --require-pass', retained)

    def test_build_and_publication_follow_the_certified_candidate(self) -> None:
        build = job(self.text, "build")
        self.assertIn("ref: ${{ needs.verify.outputs.candidate }}", build)
        self.assertIn('--ref "$CANDIDATE"', build)
        for name in ("attest", "draft"):
            with self.subTest(job=name):
                condition = re.search(r"^    if: >-\n((?:      .*\n)+)", job(self.text, name),
                                      re.M)
                assert condition, name
                self.assertEqual(condition.group(1).split(),
                                 ["needs.verify.outputs.rehearsal", "==", "'false'", "&&",
                                  "needs.verify.outputs.certification", "==", "'CERTIFIED'"])

    def test_no_job_checks_out_a_moving_ref_after_verification(self) -> None:
        refs = re.findall(r"^\s+ref: (.*)$", self.text, re.M)
        self.assertEqual(refs, ["${{ github.event_name == 'push' && github.ref || inputs.ref }}",
                                "${{ needs.verify.outputs.candidate }}"])


if __name__ == "__main__":
    unittest.main()
