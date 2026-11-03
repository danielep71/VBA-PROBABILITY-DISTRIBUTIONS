"""Tests for .github/workflows/pr-milestones.yml (Keep milestones issue-only).

Standard library only: the structure checks read the workflow text by
indentation, so they need no YAML parser, and the script tests run the
workflow's inline Python against an in-memory GitHub API. Run from the
repository root:

    python -m unittest discover -s tests -p 'test_pr_milestones_workflow.py'
"""

from __future__ import annotations

import io
import json
import os
import re
import tempfile
import textwrap
import unittest
import unittest.mock
from email.message import Message
from pathlib import Path
from typing import Any
from urllib.error import HTTPError

ROOT = Path(__file__).resolve().parents[1]
WORKFLOW = ROOT / ".github/workflows/pr-milestones.yml"

# The only permissions the job may hold.
ALLOWED_PERMISSIONS = {"issues": "read", "pull-requests": "write"}
# Events that run with the pull request's code or let its author drive the run.
FORBIDDEN_EVENTS = {"pull_request", "pull_request_review", "pull_request_review_comment",
                    "issue_comment", "workflow_run"}
EXPECTED_TYPES = ["opened", "reopened", "edited", "milestoned"]

Json = Any


def workflow_text() -> str:
    return WORKFLOW.read_text(encoding="utf-8")


def strip_comment(line: str) -> str:
    """The line without a trailing YAML comment (the workflow quotes no '#')."""
    return line.split(" #", 1)[0].rstrip() if not line.lstrip().startswith("#") else ""


def block(lines: list[str], start: int) -> list[str]:
    """The lines nested under ``lines[start]``, by indentation."""
    indent = len(lines[start]) - len(lines[start].lstrip())
    nested: list[str] = []
    for line in lines[start + 1:]:
        if line.strip() and len(line) - len(line.lstrip()) <= indent:
            break
        nested.append(line)
    return nested


def children(lines: list[str]) -> dict[str, list[str]]:
    """The direct child keys of an indented block and the lines under each."""
    content = [line for line in lines if strip_comment(line)]
    if not content:
        return {}
    indent = min(len(line) - len(line.lstrip()) for line in content)
    result: dict[str, list[str]] = {}
    for index, line in enumerate(lines):
        text = strip_comment(line)
        match = re.match(rf"^ {{{indent}}}([\w-]+):(.*)$", text)
        if match:
            result[match.group(1)] = [match.group(2).strip(), *block(lines, index)]
    return result


def run_script(text: str) -> str:
    """The inline Python of the job's step, dedented."""
    match = re.search(r"<<'PY'\n(.*?)\n\s*PY\n", text, re.S)
    if not match:
        raise ValueError("no inline Python heredoc found")
    return textwrap.dedent(match.group(1))


def structure_findings(text: str) -> list[str]:
    """Every way ``text`` departs from the pull-request-safe workflow shape."""
    lines = text.splitlines()
    top = children(lines)
    findings: list[str] = []

    events = children(top.get("on", [""])[1:])
    if "pull_request_target" not in events:
        findings.append("trigger: pull_request_target is missing")
    else:
        types = children(events["pull_request_target"][1:]).get("types", [""])[0]
        if [item.strip() for item in types.strip("[]").split(",")] != EXPECTED_TYPES:
            findings.append(f"trigger: pull_request_target types are {types!r}")
    for event in sorted(FORBIDDEN_EVENTS & events.keys()):
        findings.append(f"trigger: {event} runs pull request code or input")

    if top.get("permissions", [""])[0] != "{}":
        findings.append("permissions: the workflow default must be {}")

    jobs = children(top.get("jobs", [""])[1:])
    if list(jobs) != ["reconcile"]:
        findings.append(f"jobs: expected only 'reconcile', found {sorted(jobs)}")
    for name, job in jobs.items():
        permissions = children(job[1:]).get("permissions")
        if permissions is None or permissions[0]:
            findings.append(f"permissions: job {name} needs an explicit permission map")
            continue
        granted = {key: value[0] for key, value in children(permissions[1:]).items()}
        if granted != ALLOWED_PERMISSIONS:
            findings.append(f"permissions: job {name} holds {granted}")

    code = [strip_comment(line) for line in lines]
    if any(re.match(r"^\s*(-\s+)?uses:", line) for line in code):
        findings.append("steps: the job uses an action; it must run only its inline script")
    if any("actions/checkout" in line for line in code):
        findings.append("steps: the workflow checks out code")
    if any(re.search(r"github\.head_ref|pull_request\.head\.", line) for line in code):
        findings.append("steps: the workflow references the pull request head")
    try:
        if "${{" in run_script(text):
            findings.append("steps: the inline script interpolates an expression")
    except ValueError as error:
        findings.append(f"steps: {error}")
    return findings


class WorkflowStructureTests(unittest.TestCase):
    def test_the_workflow_has_the_safe_shape(self) -> None:
        self.assertEqual(structure_findings(workflow_text()), [])

    def mutated(self, old: str, new: str) -> list[str]:
        text = workflow_text()
        self.assertIn(old, text)
        return structure_findings(text.replace(old, new, 1))

    def assertRejected(self, findings: list[str], fragment: str) -> None:
        self.assertTrue(any(fragment in finding for finding in findings), findings)

    def test_a_pull_request_trigger_is_rejected(self) -> None:
        findings = self.mutated("  pull_request_target:\n", "  pull_request:\n")
        self.assertRejected(findings, "pull_request_target is missing")
        self.assertRejected(findings, "pull_request runs pull request code")

    def test_changed_event_types_are_rejected(self) -> None:
        self.assertRejected(self.mutated("edited, milestoned]", "edited, synchronize]"), "types are")

    def test_a_checkout_step_is_rejected(self) -> None:
        step = ("      - uses: actions/checkout@3d3c42e5aac5ba805825da76410c181273ba90b1 # v7.0.1\n"
                "        with:\n          ref: ${{ github.event.pull_request.head.sha }}\n")
        findings = self.mutated("    steps:\n", "    steps:\n" + step)
        self.assertRejected(findings, "uses an action")
        self.assertRejected(findings, "checks out code")
        self.assertRejected(findings, "pull request head")

    def test_a_broader_job_permission_is_rejected(self) -> None:
        self.assertRejected(self.mutated("      pull-requests: write\n",
                                         "      pull-requests: write\n      contents: write\n"), "holds")
        self.assertRejected(self.mutated("      issues: read\n", "      issues: write\n"), "holds")

    def test_a_job_permission_shorthand_is_rejected(self) -> None:
        text = workflow_text()
        start = text.index("    permissions:\n")
        end = text.index("    env:\n", start)
        findings = structure_findings(text[:start] + "    permissions: write-all\n" + text[end:])
        self.assertRejected(findings, "explicit permission map")

    def test_a_workflow_default_permission_is_rejected(self) -> None:
        self.assertRejected(self.mutated("permissions: {}\n", "permissions: write-all\n"), "must be {}")

    def test_a_second_job_is_rejected(self) -> None:
        extra = ("  build:\n    runs-on: ubuntu-24.04\n    permissions:\n      contents: read\n"
                 "    steps:\n      - run: echo\n")
        self.assertRejected(self.mutated("jobs:\n", "jobs:\n" + extra), "expected only 'reconcile'")

    def test_an_expression_in_the_script_is_rejected(self) -> None:
        findings = self.mutated('repository = os.environ["GITHUB_REPOSITORY"]',
                                'repository = "${{ github.event.pull_request.title }}"')
        self.assertRejected(findings, "interpolates an expression")


class FakeGitHub:
    """The issues and pulls endpoints the script calls, held in memory."""

    def __init__(self, pulls: list[dict[str, Json]], issues: tuple[dict[str, Json], ...] = (),
                 read_only: bool = False) -> None:
        self.items = {pull["number"]: dict(pull, pull_request={}) for pull in pulls}
        self.items.update({issue["number"]: dict(issue) for issue in issues})
        self.patches: list[int] = []
        self.read_only = read_only

    def __call__(self, method: str, path: str, data: Json = None) -> Json:
        if path.startswith("/repos/o/r/pulls?"):
            match = re.search(r"[?&]page=(\d+)", path)
            assert match is not None
            page = int(match.group(1))
            pulls = [item for item in self.items.values() if "pull_request" in item]
            return [dict(item) for item in pulls[(page - 1) * 100:page * 100]]
        number = int(path.rsplit("/", 1)[1])
        if method == "PATCH":
            if self.read_only:
                raise HTTPError(path, 403, "Resource not accessible by integration", Message(), None)
            self.patches.append(number)
            self.items[number]["milestone"] = data["milestone"]
        return dict(self.items[number])


def milestone(title: str = "v1.0.0") -> dict[str, str]:
    return {"title": title}


def load_script() -> dict[str, Any]:
    namespace: dict[str, Any] = {"__name__": "pr_milestones"}
    exec(run_script(workflow_text()), namespace)  # trusted: the script under test, from this repository
    return namespace


class ReconcileTests(unittest.TestCase):
    def setUp(self) -> None:
        self.script = load_script()
        quiet = unittest.mock.patch("sys.stdout", io.StringIO())
        quiet.start()
        self.addCleanup(quiet.stop)

    def test_a_pull_request_event_clears_only_that_pull_request(self) -> None:
        api = FakeGitHub([{"number": 1, "milestone": milestone()}, {"number": 2, "milestone": milestone()}])
        removed = self.script["reconcile"](api, "o/r", {"number": 1, "pull_request": {"number": 1}})
        self.assertEqual(removed, [(1, "v1.0.0")])
        self.assertEqual(api.patches, [1])
        self.assertIsNotNone(api.items[2]["milestone"])

    def test_a_pull_request_without_a_milestone_is_left_alone(self) -> None:
        api = FakeGitHub([{"number": 3, "milestone": None}])
        self.assertEqual(self.script["reconcile"](api, "o/r", {"pull_request": {"number": 3}}), [])
        self.assertEqual(api.patches, [])

    def test_a_scan_clears_every_pull_request_and_never_touches_issues(self) -> None:
        pulls = [{"number": n, "milestone": milestone() if n % 2 else None} for n in range(1, 151)]
        api = FakeGitHub(pulls, issues=({"number": 500, "milestone": milestone()},))
        removed = self.script["reconcile"](api, "o/r", {"schedule": "17 3 * * *"})
        self.assertEqual(api.patches, list(range(1, 151, 2)))
        self.assertEqual([number for number, _ in removed], list(range(1, 151, 2)))
        self.assertIsNotNone(api.items[500]["milestone"])

    def test_an_issue_number_is_refused(self) -> None:
        api = FakeGitHub([], issues=({"number": 7, "milestone": milestone()},))
        with self.assertRaisesRegex(ValueError, "Refusing to change issue #7"):
            self.script["reconcile"](api, "o/r", {"pull_request": {"number": 7}})
        self.assertEqual(api.patches, [])

    def test_a_milestone_that_survives_the_change_fails(self) -> None:
        api = FakeGitHub([{"number": 4, "milestone": milestone()}])

        def stuck(method: str, path: str, data: Json = None) -> Json:
            return dict(api.items[4]) if method == "PATCH" else api(method, path, data)

        with self.assertRaisesRegex(ValueError, "still has a milestone"):
            self.script["reconcile"](stuck, "o/r", {"pull_request": {"number": 4}})

    def test_the_summary_lists_each_removal(self) -> None:
        text = self.script["summary"]([(5, "v1.0.0"), (6, "a|b")], "every pull request")
        self.assertIn("| #5 | v1.0.0 |", text)
        self.assertIn("| #6 | a\\|b |", text)
        self.assertIn("Scope: every pull request.", text)
        self.assertIn("nothing was removed", self.script["summary"]([], "pull request #9"))


class MainTests(unittest.TestCase):
    """main() end to end: the event file, the API calls and the job summary."""

    def run_main(self, api: FakeGitHub, event: dict[str, Json]) -> tuple[str, str]:
        script = load_script()

        def urlopen(request: Any, timeout: float) -> io.BytesIO:
            self.assertTrue(request.full_url.startswith("https://api.example.test/"))
            self.assertEqual(request.get_header("Authorization"), "Bearer token")
            body = json.loads(request.data) if request.data else None
            path = request.full_url.removeprefix("https://api.example.test")
            return io.BytesIO(json.dumps(api(request.get_method(), path, body)).encode())

        script["urlopen"] = urlopen
        with tempfile.TemporaryDirectory() as folder:
            event_path = Path(folder, "event.json")
            event_path.write_text(json.dumps(event), encoding="utf-8")
            summary_path = Path(folder, "summary.md")
            environment = {"GITHUB_REPOSITORY": "o/r", "GITHUB_EVENT_PATH": str(event_path),
                           "GITHUB_STEP_SUMMARY": str(summary_path), "GH_TOKEN": "token",
                           "GITHUB_API_URL": "https://api.example.test"}
            output = io.StringIO()
            with unittest.mock.patch.dict(os.environ, environment), \
                    unittest.mock.patch("sys.stdout", output):
                script["main"]()
            summary = summary_path.read_text(encoding="utf-8") if summary_path.exists() else ""
        return summary, output.getvalue()

    def test_the_job_summary_records_the_removal(self) -> None:
        api = FakeGitHub([{"number": 8, "milestone": milestone("v1.0.1")}])
        summary, _ = self.run_main(api, {"number": 8, "pull_request": {"number": 8}})
        self.assertEqual(api.patches, [8])
        self.assertIn("Scope: pull request #8.", summary)
        self.assertIn("| #8 | v1.0.1 |", summary)

    def test_a_scan_with_nothing_to_remove_says_so(self) -> None:
        summary, _ = self.run_main(FakeGitHub([{"number": 1, "milestone": None}]), {"inputs": {}})
        self.assertIn("Scope: every pull request.", summary)
        self.assertIn("nothing was removed", summary)

    def test_a_read_only_token_on_a_pull_request_event_degrades_with_a_warning(self) -> None:
        api = FakeGitHub([{"number": 9, "milestone": milestone()}], read_only=True)
        summary, output = self.run_main(api, {"pull_request": {"number": 9}})
        self.assertIn("::warning title=Milestone not removed::", output)
        self.assertIn("cannot change pull request #9 (HTTP 403", summary)
        self.assertIsNotNone(api.items[9]["milestone"])

    def test_a_read_only_token_fails_a_scan(self) -> None:
        api = FakeGitHub([{"number": 9, "milestone": milestone()}], read_only=True)
        with self.assertRaises(HTTPError):
            self.run_main(api, {"schedule": "17 3 * * *"})


if __name__ == "__main__":
    unittest.main()
