"""Prove that the actionlint probe and mypy's strict configuration are active.

The probe tests run tools/probe_actionlint.py against stand-in actionlint
commands, so they need no download: a command that accepts everything, one
that rejects everything, one that fails for an unrelated reason and one that
cannot start must each fail the probe, and a faithful stand-in must pass it.
The real actionlint and mypy are exercised when they are installed (as in the
Static checks workflow) and skipped otherwise.
"""

from __future__ import annotations

import contextlib
import importlib.util
import io
import os
import shutil
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "tools"))

import probe_actionlint as P  # noqa: E402

FAITHFUL = """\
import sys
text = open(sys.argv[-1], encoding="utf-8").read()
if "bogus-key" in text:
    print('invalid.yml:12:9: unexpected key "bogus-key" for "step" section [syntax-check]')
    print('invalid.yml:11:24: undefined variable "no_such_context" [expression]')
    sys.exit(1)
"""


def stand_in(source: str) -> list[str]:
    return [sys.executable, "-c", source]


def actionlint_path() -> str | None:
    return shutil.which(os.environ.get("ACTIONLINT") or "actionlint")


class ActionlintProbeTest(unittest.TestCase):
    def run_probe(self, command: list[str]) -> tuple[int, str]:
        with contextlib.redirect_stdout(io.StringIO()) as printed, \
                contextlib.redirect_stderr(io.StringIO()):
            code = P.main(["--", *command])
        return code, printed.getvalue()

    def test_faithful_actionlint_passes(self) -> None:
        code, printed = self.run_probe(stand_in(FAITHFUL))
        self.assertEqual(code, 0, printed)
        self.assertIn("ACTIONLINT PROBE PASS", printed)

    def test_command_that_accepts_everything_fails(self) -> None:
        # The case the probe exists for: an invocation that checks nothing.
        code, printed = self.run_probe(stand_in("pass"))
        self.assertEqual(code, 1)
        self.assertIn("the invalid workflow must fail; exit code 0", printed)

    def test_command_that_rejects_everything_fails(self) -> None:
        source = ("print('unexpected key \"bogus-key\"; undefined variable "
                  "\"no_such_context\"'); raise SystemExit(1)")
        code, printed = self.run_probe(stand_in(source))
        self.assertEqual(code, 1)
        self.assertIn("the valid workflow must pass with no finding; exit code 1", printed)

    def test_failure_for_an_unrelated_reason_fails(self) -> None:
        source = ("import sys\nif 'bogus-key' in open(sys.argv[-1]).read():\n"
                  "    sys.exit('could not read config file')\n")
        code, printed = self.run_probe(stand_in(source))
        self.assertEqual(code, 1)
        self.assertIn('must report unexpected key "bogus-key"', printed)
        self.assertIn("could not read config file", printed)

    def test_command_that_cannot_start_fails(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            code, printed = self.run_probe([str(Path(directory) / "no-such-actionlint")])
        self.assertEqual(code, 1)
        self.assertIn("exit code None", printed)

    def test_no_command_is_a_usage_error(self) -> None:
        self.assertEqual(self.run_probe([])[0], 2)

    @unittest.skipIf(actionlint_path() is None, "actionlint is not installed")
    def test_real_actionlint_with_the_repository_configuration(self) -> None:
        command = [str(actionlint_path()), "-no-color", "-config-file", ".github/actionlint.yaml"]
        completed = subprocess.run([sys.executable, "tools/probe_actionlint.py", "--", *command],
                                   cwd=ROOT, capture_output=True, text=True, check=False)
        self.assertEqual(completed.returncode, 0, completed.stdout + completed.stderr)


@unittest.skipIf(importlib.util.find_spec("mypy") is None, "mypy is not installed: "
                 "python -m pip install -r tools/requirements-lint-ci.txt")
class MypyStrictTest(unittest.TestCase):
    def test_strict_configuration_rejects_an_unannotated_function(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            probe = Path(directory) / "strict_probe.py"
            command = [sys.executable, "-m", "mypy", "--config-file", str(ROOT / "mypy.ini"),
                       "--no-pretty", "--cache-dir", str(Path(directory) / "cache"), str(probe)]
            probe.write_text("def probe(value: int) -> int:\n    return value\n",
                             encoding="utf-8")
            ok = subprocess.run(command, cwd=ROOT, capture_output=True, text=True, check=False)
            probe.write_text("def probe(value):\n    return value\n", encoding="utf-8")
            bad = subprocess.run(command, cwd=ROOT, capture_output=True, text=True, check=False)
        self.assertEqual(ok.returncode, 0, ok.stdout + ok.stderr)
        self.assertNotEqual(bad.returncode, 0, bad.stdout + bad.stderr)
        self.assertIn("[no-untyped-def]", bad.stdout)


if __name__ == "__main__":
    unittest.main()
