#!/usr/bin/env python3
"""Prove that the actionlint invocation of the static checks rejects an invalid workflow.

A passing actionlint check means something only if the same command would fail
on a broken workflow: a wrong path, a configuration that disables the checks,
or a wrapper that exits 0 would otherwise pass silently. The probe runs the
given actionlint command on two workflow files in a temporary directory:

- a valid workflow, which must pass with no finding, and
- the same workflow with an unknown step key and an expression naming an
  undefined context, which must fail and report both findings.

    python3 tools/probe_actionlint.py -- actionlint -no-color \\
        -config-file .github/actionlint.yaml

Everything after `--` is the actionlint command; the probe appends the
workflow path. The exit code is 0 only when both expectations hold.
"""

from __future__ import annotations

import subprocess
import sys
import tempfile
from collections.abc import Sequence
from pathlib import Path

TIMEOUT_SECONDS = 120

VALID = """\
name: Probe
on: push
permissions: {}
jobs:
  probe:
    runs-on: ubuntu-24.04
    steps:
      - name: Greet
        run: echo "$GREETING"
        env:
          GREETING: ${{ github.sha }}
"""

# One unknown key and one undefined context: each must be reported.
INVALID = VALID.replace("${{ github.sha }}", "${{ no_such_context.value }}") + \
    "        bogus-key: true\n"
EXPECTED = ('unexpected key "bogus-key"', 'undefined variable "no_such_context"')


def lint(command: Sequence[str], directory: Path, name: str,
         text: str) -> tuple[int | None, str]:
    """Run the command on one workflow file; the exit code is None if it cannot run."""
    path = directory / name
    path.write_text(text, encoding="utf-8")
    try:
        completed = subprocess.run([*command, str(path)], capture_output=True, text=True,
                                   timeout=TIMEOUT_SECONDS, check=False)
    except (OSError, subprocess.TimeoutExpired) as error:
        return None, f"{type(error).__name__}: {error}"
    return completed.returncode, completed.stdout + completed.stderr


def shown(output: str) -> str:
    return output.rstrip() or "(no output)"


def probe(command: Sequence[str]) -> list[str]:
    """Every way the command fails to behave as a working actionlint."""
    problems: list[str] = []
    with tempfile.TemporaryDirectory(prefix="actionlint-probe-") as directory:
        root = Path(directory)
        code, output = lint(command, root, "valid.yml", VALID)
        if code != 0 or output.strip():
            problems.append(f"the valid workflow must pass with no finding; exit code {code}:"
                            f"\n{shown(output)}")
        code, output = lint(command, root, "invalid.yml", INVALID)
        if code is None or code == 0:
            problems.append(f"the invalid workflow must fail; exit code {code}")
        missing = [finding for finding in EXPECTED if finding not in output]
        if missing:
            problems.append("the invalid workflow must report "
                            + " and ".join(missing) + f"; output:\n{shown(output)}")
    return problems


def main(argv: Sequence[str] | None = None) -> int:
    arguments = list(sys.argv[1:] if argv is None else argv)
    if arguments[:1] == ["--"]:
        arguments = arguments[1:]
    if not arguments:
        print("usage: probe_actionlint.py -- ACTIONLINT [OPTION ...]", file=sys.stderr)
        return 2
    problems = probe(arguments)
    if problems:
        print("ACTIONLINT PROBE FAIL: the command does not reject an invalid workflow")
        for problem in problems:
            print("  - " + problem.replace("\n", "\n    "))
        return 1
    print("ACTIONLINT PROBE PASS: the valid workflow passes; the invalid one fails with "
          + " and ".join(EXPECTED))
    return 0


if __name__ == "__main__":
    sys.exit(main())
