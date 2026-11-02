"""Every workflow uses one exact CI Python, and the minimum Python is stated consistently.

Each workflow that sets up Python defines PYTHON_VERSION once, as an exact
patch release in its top-level env, and every setup-python step reads it (or,
in the minimum-python job, MINIMUM_PYTHON_VERSION). All workflows agree on the
version. The minimum Python is the same series in the minimum-python job,
mypy.ini, ruff.toml and tools/README.md.
"""

from __future__ import annotations

import re
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
WORKFLOWS = ROOT / ".github" / "workflows"
EXACT = re.compile(r"\d+\.\d+\.\d+")
VARIABLES = ("PYTHON_VERSION", "MINIMUM_PYTHON_VERSION")
ALLOWED = {f"${{{{ env.{name} }}}}" for name in VARIABLES}


def top_level_env(text: str) -> dict[str, list[str]]:
    """Values of each variable in the workflow-level env block."""
    values: dict[str, list[str]] = {}
    inside = False
    for line in text.splitlines():
        if re.match(r"^env:\s*$", line):
            inside = True
            continue
        if inside and line and not line.startswith((" ", "#")):
            inside = False
        match = re.match(r'^  ([A-Z_]+):\s*"?([^"#\s]*)"?', line) if inside else None
        if match:
            values.setdefault(match.group(1), []).append(match.group(2))
    return values


def workflow_findings(text: str) -> list[str]:
    """Every way one workflow departs from the single exact Python pin."""
    findings = []
    for line in text.splitlines():
        if re.match(r"^\s*python-version-file:", line):
            findings.append("python-version-file is not used; read env.PYTHON_VERSION")
        match = re.match(r"^\s*python-version:\s*(.+?)\s*$", line)
        if match and match.group(1) not in ALLOWED:
            findings.append(f"python-version must read env.PYTHON_VERSION: {match.group(1)}")
    if "actions/setup-python@" not in text:
        return findings
    env = top_level_env(text)
    for name in VARIABLES:
        everywhere = len(re.findall(rf"^\s+{name}:", text, flags=re.MULTILINE))
        defined = env.get(name, [])
        if name == "PYTHON_VERSION" and not defined:
            findings.append("PYTHON_VERSION is not defined in the top-level env")
        if everywhere > len(defined) or len(defined) > 1:
            findings.append(f"{name} must be defined once, in the top-level env")
        findings += [f"{name} is not an exact patch release: {value!r}"
                     for value in defined if not EXACT.fullmatch(value)]
    return findings


def series(version: str) -> str:
    return ".".join(version.split(".")[:2])


VALID = """\
name: Fixture
env:
  # The CI interpreter.
  PYTHON_VERSION: "3.12.16"
jobs:
  one:
    env:
      PYTHONUTF8: "1"
    steps:
      - uses: actions/setup-python@0000000000000000000000000000000000000000 # v0
        with:
          python-version: ${{ env.PYTHON_VERSION }}
"""


class WorkflowFixtureTest(unittest.TestCase):
    def test_valid_workflow(self) -> None:
        self.assertEqual(workflow_findings(VALID), [])
        self.assertEqual(top_level_env(VALID), {"PYTHON_VERSION": ["3.12.16"]})

    def assertRejected(self, text: str, expected: str) -> None:
        findings = workflow_findings(text)
        self.assertTrue(any(expected in finding for finding in findings), findings)

    def test_floating_or_literal_version_in_a_step(self) -> None:
        self.assertRejected(VALID.replace("${{ env.PYTHON_VERSION }}", '"3.12"'),
                            "must read env.PYTHON_VERSION")
        self.assertRejected(VALID.replace("${{ env.PYTHON_VERSION }}", '"3.12.16"'),
                            "must read env.PYTHON_VERSION")

    def test_floating_pin(self) -> None:
        self.assertRejected(VALID.replace('"3.12.16"', '"3.12"'), "not an exact patch release")
        self.assertRejected(VALID.replace('"3.12.16"', '"3.12.x"'), "not an exact patch release")

    def test_missing_or_repeated_definition(self) -> None:
        self.assertRejected(VALID.replace('  PYTHON_VERSION: "3.12.16"\n', ""), "not defined")
        self.assertRejected(VALID.replace('PYTHONUTF8: "1"', 'PYTHON_VERSION: "3.12.15"'),
                            "defined once")

    def test_version_file(self) -> None:
        self.assertRejected(VALID.replace("python-version: ${{ env.PYTHON_VERSION }}",
                                          "python-version-file: .python-version"),
                            "python-version-file")


class RepositoryTest(unittest.TestCase):
    def test_every_workflow_uses_one_exact_version(self) -> None:
        versions = {}
        for path in sorted(WORKFLOWS.glob("*.yml")):
            text = path.read_text(encoding="utf-8")
            self.assertEqual(workflow_findings(text), [], path.name)
            if "actions/setup-python@" in text:
                versions[path.name] = top_level_env(text)["PYTHON_VERSION"][0]
        self.assertIn("static-checks.yml", versions)
        self.assertEqual(len(set(versions.values())), 1, versions)

    def test_minimum_python_is_stated_consistently(self) -> None:
        static = (WORKFLOWS / "static-checks.yml").read_text(encoding="utf-8")
        env = top_level_env(static)
        minimum = series(env["MINIMUM_PYTHON_VERSION"][0])
        ci = series(env["PYTHON_VERSION"][0])
        self.assertLess(tuple(map(int, minimum.split("."))), tuple(map(int, ci.split("."))))
        self.assertIn(f"name: Minimum Python ({minimum}, not required)", static)
        self.assertIn("python-version: ${{ env.MINIMUM_PYTHON_VERSION }}", static)
        mypy = (ROOT / "mypy.ini").read_text(encoding="utf-8")
        self.assertRegex(mypy, rf"(?m)^python_version = {re.escape(minimum)}$")
        ruff = (ROOT / "ruff.toml").read_text(encoding="utf-8")
        self.assertIn(f'target-version = "py{minimum.replace(".", "")}"', ruff)
        readme = (ROOT / "tools" / "README.md").read_text(encoding="utf-8")
        self.assertIn(f"The minimum Python is **{minimum}**", readme)
        self.assertIn(f"`{env['PYTHON_VERSION'][0]}`", readme)


if __name__ == "__main__":
    unittest.main()
