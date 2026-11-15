"""Rules every workflow follows, each with a fixture that breaks it (#353).

The rules are read from the parsed YAML (tests/strict_yaml.py). They are the
ones the workflows already keep; the tests stop a new job or an edit from
dropping one unnoticed:

  pinned actions      every `uses:` of an action names a 40-hex commit;
  checkout            every actions/checkout sets persist-credentials: false,
                      except the jobs listed in PERSISTS_CREDENTIALS;
  permissions         every workflow declares top-level permissions with no
                      write, and only the jobs in WRITING_JOBS grant a write;
  timeout             every job declares timeout-minutes;
  concurrency         every workflow declares concurrency;
  pip                 every `pip install` requires hashes;
  toolchain           every hosted job names the exact runner image, and
                      every Python version is an exact x.y.z;
  Excel runner        a job on the self-hosted runner that a pull request
                      can reach refuses forks and Dependabot;
  duplicate keys      none in the workflows, Dependabot's configuration and
                      the issue forms.

Timeouts are also checked by tests/test_workflow_timeouts.py, which reads
the files as text and needs no parser.

The parser is PyYAML, installed hash-locked from
tools/requirements-lint-ci.txt in the Static checks workflow. Where it is
not installed these tests are skipped; test_the_parser_is_hash_locked does
not need it and keeps the lock in place.
"""
from __future__ import annotations

import re
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
WORKFLOWS = ROOT / ".github" / "workflows"
LOCK = ROOT / "tools" / "requirements-lint-ci.txt"

try:
    from tests import strict_yaml
except ImportError:                                   # run as a script, or PyYAML missing
    try:
        import strict_yaml  # type: ignore[no-redef]
    except ImportError:
        strict_yaml = None  # type: ignore[assignment]

RUNNER_IMAGE = "ubuntu-24.04"
SELF_HOSTED = ["self-hosted", "Windows", "X64", "excel"]
SHA = re.compile(r"@[0-9a-f]{40}$")
EXACT_PYTHON = re.compile(r"\d+\.\d+\.\d+")
PIP_INSTALL = re.compile(r"\bpip3?\s+install\b")

# (workflow file, job): the jobs that may keep the checkout's credentials.
# traffic pushes the daily snapshot to the traffic-history branch with the
# job's own token and runs no repository code; decided on 2026-11-15 to keep
# it so, rather than pass a token to each git command.
PERSISTS_CREDENTIALS = {("daily-traffic.yml", "traffic")}

# (workflow file, job) -> the scopes it may write. Anything else is read-only.
WRITING_JOBS = {
    ("codeql.yml", "analyze"): {"security-events"},
    ("daily-traffic.yml", "traffic"): {"contents"},
    ("daily-traffic.yml", "alert"): {"issues"},
    ("excel-runner-watch.yml", "retry"): {"actions"},
    ("excel-runner-watch.yml", "watch"): {"issues"},
    ("labels-sync.yml", "reconcile"): {"issues"},
    ("pr-milestones.yml", "reconcile"): {"pull-requests"},
    ("release.yml", "attest"): {"id-token", "attestations"},
    ("release.yml", "draft"): {"contents"},
    ("scorecard.yml", "published"): {"id-token", "security-events"},
}


FORK_GUARD = "github.event.pull_request.head.repo.full_name == github.repository"
DEPENDABOT_GUARD = "github.event.pull_request.user.login != 'dependabot[bot]'"


def pull_request_reaches(document: dict) -> bool:
    """Whether the workflow runs on pull_request."""
    triggers = strict_yaml.events(document)
    if isinstance(triggers, str):
        return triggers == "pull_request"
    return "pull_request" in (triggers or ())


def writes(permissions: object) -> set[str]:
    if permissions == "write-all":
        return {"*"}
    if isinstance(permissions, dict):
        return {str(scope) for scope, level in permissions.items() if level == "write"}
    return set()


def problems(name: str, document: dict) -> list[str]:
    """Every rule the workflow `name` breaks."""
    out: list[str] = []
    if "concurrency" not in document:
        out.append("no concurrency")
    top = document.get("permissions")
    if top is None:
        out.append("no top-level permissions")
    elif writes(top):
        out.append(f"top-level permissions grant write: {sorted(writes(top))}")
    env = document.get("env") or {}
    for key, value in env.items():
        if "PYTHON_VERSION" in str(key) and not EXACT_PYTHON.fullmatch(str(value)):
            out.append(f"{key} is not an exact version: {value!r}")
    for job, body in (document.get("jobs") or {}).items():
        where = f"job {job}"
        if "uses" in body:                            # a reusable workflow carries its own rules
            if not SHA.search(str(body["uses"])) and not str(body["uses"]).startswith("./"):
                out.append(f"{where}: reusable workflow not pinned to a commit")
            continue
        if "timeout-minutes" not in body:
            out.append(f"{where}: no timeout-minutes")
        runner = body.get("runs-on")
        if runner != RUNNER_IMAGE and runner != SELF_HOSTED:
            out.append(f"{where}: runs-on {runner!r} is neither {RUNNER_IMAGE} nor the Excel runner")
        if runner == SELF_HOSTED and pull_request_reaches(document):
            guard = " ".join(str(body.get("if", "")).split())
            if FORK_GUARD not in guard:
                out.append(f"{where}: a pull request reaches the Excel runner without the fork guard")
            if DEPENDABOT_GUARD not in guard:
                out.append(f"{where}: a pull request reaches the Excel runner without the Dependabot guard")
        granted = writes(body.get("permissions", top))
        allowed = WRITING_JOBS.get((name, job), set())
        if granted != allowed:
            out.append(f"{where}: writes {sorted(granted)}, listed {sorted(allowed)}")
        for step in body.get("steps") or []:
            uses = str(step.get("uses", ""))
            if uses and not uses.startswith("./") and not SHA.search(uses):
                out.append(f"{where}: action not pinned to a commit: {uses}")
            if uses.startswith("actions/checkout@"):
                persisted = (step.get("with") or {}).get("persist-credentials")
                if (name, job) in PERSISTS_CREDENTIALS:
                    if persisted is not True:
                        out.append(f"{where}: listed as persisting credentials but does not")
                elif persisted is not False:
                    out.append(f"{where}: checkout without persist-credentials: false")
            if uses.startswith("actions/setup-python@"):
                version = str((step.get("with") or {}).get("python-version", ""))
                if not (EXACT_PYTHON.fullmatch(version) or re.fullmatch(r"\$\{\{ env\.\w+ \}\}", version)):
                    out.append(f"{where}: python-version {version!r} is neither exact nor an env version")
            run = str(step.get("run", ""))
            if PIP_INSTALL.search(run) and "--require-hashes" not in run:
                out.append(f"{where}: pip install without --require-hashes")
    return out


GOOD = """
name: Fixture
on: push
permissions:
  contents: read
concurrency:
  group: fixture
env:
  PYTHON_VERSION: "3.12.16"
jobs:
  build:
    runs-on: ubuntu-24.04
    timeout-minutes: 5
    steps:
      - uses: actions/checkout@3d3c42e5aac5ba805825da76410c181273ba90b1
        with:
          persist-credentials: false
      - uses: actions/setup-python@6d4fd22b20465ccc0f7354476588e8df7c9ca0e0
        with:
          python-version: ${{ env.PYTHON_VERSION }}
      - run: python -m pip install --require-hashes -r tools/requirements-lint-ci.txt
"""

# A job on the Excel runner in a workflow that pull requests trigger.
GOOD_EXCEL = """
name: Fixture
on:
  pull_request:
  push:
permissions:
  contents: read
concurrency:
  group: fixture
jobs:
  excel:
    if: >-
      github.event_name != 'pull_request' ||
      (github.event.pull_request.head.repo.full_name == github.repository &&
      github.event.pull_request.user.login != 'dependabot[bot]')
    runs-on: [self-hosted, Windows, X64, excel]
    timeout-minutes: 5
    steps:
      - run: echo
"""

BROKEN_EXCEL = {
    "no Dependabot guard": (" &&\n      github.event.pull_request.user.login != 'dependabot[bot]')", ")",
                            "Dependabot guard"),
    "no fork guard": ("github.event.pull_request.head.repo.full_name == github.repository &&", "",
                      "fork guard"),
    "no condition": ("    if: >-\n      github.event_name != 'pull_request' ||\n"
                     "      (github.event.pull_request.head.repo.full_name == github.repository &&\n"
                     "      github.event.pull_request.user.login != 'dependabot[bot]')\n", "", "fork guard"),
}

# One edit of GOOD per rule, and a word of the finding it must produce.
BROKEN = {
    "unpinned action": ("actions/checkout@3d3c42e5aac5ba805825da76410c181273ba90b1", "actions/checkout@v7",
                        "not pinned"),
    "credentials kept": ("persist-credentials: false", "persist-credentials: true", "persist-credentials"),
    "credentials unset": ("        with:\n          persist-credentials: false\n", "", "persist-credentials"),
    "no timeout": ("    timeout-minutes: 5\n", "", "timeout-minutes"),
    "no concurrency": ("concurrency:\n  group: fixture\n", "", "concurrency"),
    "no permissions": ("permissions:\n  contents: read\n", "", "top-level permissions"),
    "top-level write": ("  contents: read\n", "  contents: write\n", "top-level permissions grant write"),
    "unlisted writing job": ("    timeout-minutes: 5\n",
                             "    timeout-minutes: 5\n    permissions:\n      issues: write\n", "writes"),
    "pip without hashes": ("--require-hashes -r", "-r", "require-hashes"),
    "floating image": ("runs-on: ubuntu-24.04", "runs-on: ubuntu-latest", "runs-on"),
    "floating python": ('PYTHON_VERSION: "3.12.16"', 'PYTHON_VERSION: "3.12"', "exact version"),
    "floating python in the step": ("python-version: ${{ env.PYTHON_VERSION }}", "python-version: 3.x",
                                    "python-version"),
}


class TheLock(unittest.TestCase):
    def test_the_parser_is_hash_locked(self) -> None:
        text = LOCK.read_text(encoding="utf-8")
        match = re.search(r"^pyyaml==\d+\.\d+\.\d+ \\\n((?:    --hash=sha256:[0-9a-f]{64}(?: \\)?\n)+)",
                          text, re.M | re.I)
        self.assertIsNotNone(match, "tools/requirements-lint-ci.txt does not pin PyYAML with hashes")


@unittest.skipIf(strict_yaml is None, "PyYAML is not installed; the Static checks workflow installs it")
class Rules(unittest.TestCase):
    def test_the_fixture_follows_every_rule(self) -> None:
        self.assertEqual(problems("fixture.yml", strict_yaml.loads(GOOD)), [])

    def test_each_rule_has_a_fixture_that_breaks_it(self) -> None:
        for name, (old, new, word) in BROKEN.items():
            with self.subTest(rule=name):
                self.assertIn(old, GOOD)
                found = problems("fixture.yml", strict_yaml.loads(GOOD.replace(old, new, 1)))
                self.assertTrue(any(word in text for text in found), f"{name}: {found}")

    def test_the_excel_runner_refuses_forks_and_dependabot(self) -> None:
        self.assertEqual(problems("fixture.yml", strict_yaml.loads(GOOD_EXCEL)), [])
        for label, (old, new, word) in BROKEN_EXCEL.items():
            with self.subTest(label):
                self.assertIn(old, GOOD_EXCEL)
                found = problems("fixture.yml", strict_yaml.loads(GOOD_EXCEL.replace(old, new)))
                self.assertTrue(any(word in finding for finding in found), found)
        # A workflow that pull requests do not trigger needs no guard.
        pushed = GOOD_EXCEL.replace("  pull_request:\n", "")
        unguarded = pushed.replace(BROKEN_EXCEL["no condition"][0], "")
        self.assertEqual(problems("fixture.yml", strict_yaml.loads(unguarded)), [])

    def test_a_listed_job_must_use_what_it_is_listed_for(self) -> None:
        document = strict_yaml.loads(GOOD)
        self.assertTrue(any("listed as persisting" in text
                            for text in problems("daily-traffic.yml", {**document, "jobs": {
                                "traffic": {**document["jobs"]["build"],
                                            "permissions": {"contents": "write"}}}})))

    def test_duplicate_keys_are_refused(self) -> None:
        with self.assertRaises(strict_yaml.DuplicateKeyError):
            strict_yaml.loads("a: 1\nb: 2\na: 3\n")
        with self.assertRaises(strict_yaml.DuplicateKeyError):
            strict_yaml.loads(GOOD.replace("    timeout-minutes: 5\n",
                                           "    timeout-minutes: 5\n    timeout-minutes: 360\n", 1))

    def test_every_workflow_follows_the_rules(self) -> None:
        files = sorted(WORKFLOWS.glob("*.y*ml"))
        self.assertGreaterEqual(len(files), 18)
        seen: set[tuple[str, str]] = set()
        for path in files:
            document = strict_yaml.loads(path.read_text(encoding="utf-8"))
            self.assertIsNotNone(strict_yaml.events(document), f"{path.name}: no trigger")
            self.assertEqual(problems(path.name, document), [], path.name)
            seen.update((path.name, job) for job in document["jobs"])
        self.assertEqual(sorted(set(WRITING_JOBS) - seen), [], "listed writing jobs that do not exist")
        self.assertEqual(sorted(PERSISTS_CREDENTIALS - seen), [], "listed jobs that do not exist")

    def test_no_duplicate_key_in_the_other_configuration(self) -> None:
        files = [ROOT / ".github" / "dependabot.yml",
                 *sorted((ROOT / ".github" / "ISSUE_TEMPLATE").glob("*.y*ml"))]
        self.assertGreaterEqual(len(files), 5)
        for path in files:
            with self.subTest(file=path.name):
                self.assertIsNotNone(strict_yaml.loads(path.read_text(encoding="utf-8")))


if __name__ == "__main__":
    unittest.main()
