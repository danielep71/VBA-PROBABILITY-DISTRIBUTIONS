"""Every workflow job declares timeout-minutes.

A job without one runs for up to six hours when a step hangs. The workflows
are read as text, by indentation, so the test needs no YAML parser: a job is a
key at two spaces under the top-level jobs: key, and its timeout is a key at
four spaces. A job that only calls a reusable workflow (uses: at four spaces)
cannot carry a timeout and is skipped.

Run: python3 -m unittest tests.test_workflow_timeouts
"""
import re
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
JOB = re.compile(r"^  ([A-Za-z_][A-Za-z0-9_-]*):[ \t]*(?:#.*)?$")


def jobs(text):
    """Map each job name to the lines of its body."""
    found, current, inside = {}, None, False
    for line in text.splitlines():
        if re.match(r"^jobs:[ \t]*(?:#.*)?$", line):
            inside = True
            continue
        if inside and line and not line.startswith((" ", "#")):
            break                                  # the next top-level key
        if not inside:
            continue
        match = JOB.match(line)
        if match:
            current = match.group(1)
            found[current] = []
        elif current is not None:
            found[current].append(line)
    return found


def missing_timeouts(text):
    out = []
    for name, body in jobs(text).items():
        if any(re.match(r"^    uses:", line) for line in body):
            continue
        if not any(re.match(r"^    timeout-minutes:[ \t]*\S", line) for line in body):
            out.append(name)
    return out


class WorkflowTimeouts(unittest.TestCase):
    def test_every_job_has_a_timeout(self):
        workflows = sorted((ROOT / ".github" / "workflows").glob("*.y*ml"))
        self.assertTrue(workflows)
        total = 0
        for path in workflows:
            text = path.read_text(encoding="utf-8")
            found = jobs(text)
            self.assertTrue(found, f"{path.name}: no job found")
            total += len(found)
            self.assertEqual(missing_timeouts(text), [], f"{path.name}: jobs without timeout-minutes")
        self.assertGreater(total, len(workflows) - 1)

    def test_a_job_without_a_timeout_is_reported(self):
        text = ("name: x\non:\n  push:\njobs:\n  first:\n    runs-on: ubuntu-24.04\n"
                "    timeout-minutes: 5\n    steps:\n      - run: true\n"
                "  second:\n    runs-on: ubuntu-24.04\n    steps:\n"
                "      - name: a step whose text says timeout-minutes: 5\n        run: true\n"
                "  reused:\n    uses: ./.github/workflows/other.yml\n")
        self.assertEqual(sorted(jobs(text)), ["first", "reused", "second"])
        self.assertEqual(missing_timeouts(text), ["second"])

    def test_keys_after_the_jobs_block_are_not_jobs(self):
        text = "jobs:\n  only:\n    timeout-minutes: 1\nenv:\n  NAME: value\n"
        self.assertEqual(list(jobs(text)), ["only"])


if __name__ == "__main__":
    unittest.main()
