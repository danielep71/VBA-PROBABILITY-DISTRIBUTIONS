#!/usr/bin/env python3
"""Reject a tracked file that names a machine (#491).

The repository is public, and a machine name identifies a device. Records and
documents name an Excel runner by its alias (docs/EXCEL_VBA_CI.md).

The check reads every tracked file from the Git index and fails when one holds:

- a default Windows machine name: DESKTOP-, LAPTOP- or WIN- followed by 7 to
  15 capitals and digits; or
- a line of the local deny list, compared without regard to case. The list
  is .git/info/machine-names, one name per line, or the file that
  PROBDIST_MACHINE_NAMES points to. It is never tracked, so a name that has no
  default form is caught only where the list exists.

A finding prints the path and the line number, never the name.

    python tools/check_machine_names.py             check the index
    python tools/check_machine_names.py --self-test  check the rules on made-up names
"""

from __future__ import annotations

import argparse
import os
import re
import subprocess
import sys
from collections.abc import Sequence
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
DEFAULT_NAME = re.compile(rb"(?<![A-Za-z0-9-])(?:DESKTOP|LAPTOP|WIN)-[A-Z0-9]{7,15}(?![A-Za-z0-9])")
DENY_FILE = ".git/info/machine-names"
DENY_ENV = "PROBDIST_MACHINE_NAMES"


def deny_list(root: Path) -> list[bytes]:
    """The local names, lower case; empty when there is no list."""
    path = Path(os.environ.get(DENY_ENV) or root / DENY_FILE)
    if not path.is_file():
        return []
    lines = path.read_text(encoding="utf-8").splitlines()
    return [line.strip().lower().encode() for line in lines if line.strip()]


def findings(data: bytes, denied: Sequence[bytes]) -> list[int]:
    """The 1-based numbers of the lines of data that name a machine."""
    out = []
    for number, line in enumerate(data.split(b"\n"), 1):
        lowered = line.lower()
        if DEFAULT_NAME.search(line) or any(name in lowered for name in denied):
            out.append(number)
    return out


def tracked(root: Path) -> list[str]:
    done = subprocess.run(["git", "ls-files", "-z"], cwd=root, capture_output=True, check=True)
    return [name for name in done.stdout.decode("utf-8", "surrogateescape").split("\0") if name]


def blob(root: Path, name: str) -> bytes:
    return subprocess.run(["git", "show", f":{name}"], cwd=root, capture_output=True,
                          check=True).stdout


def check(root: Path) -> int:
    denied = deny_list(root)
    failures = 0
    for name in tracked(root):
        if not (root / name).is_file():
            continue
        for number in findings(blob(root, name), denied):
            print(f"FAIL: {name}:{number} names a machine")
            failures += 1
    if failures:
        print(f"machine names: {failures} finding(s); use the runner's alias")
        return 1
    print(f"machine names: PASS ({len(denied)} local name(s) on the deny list)")
    return 0


def self_test() -> int:
    made_up = "DESKTOP-" + "A1B2C3D"
    cases = [
        (f"runner {made_up}, build 1".encode(), [], [1]),
        (b"ok\n`LAPTOP-" + b"ZZ99ZZ9" + b"`", [], [2]),
        (b"excel-runner-b and DESKTOP-short", [], []),
        (b"MYDESKTOP-" + b"A1B2C3D", [], []),
        (b"asset Tag-0042 here", [b"tag-0042"], [1]),
        (b"asset Tag-0042 here", [], []),
    ]
    bad = [data for data, denied, want in cases if findings(data, denied) != want]
    if bad:
        print(f"self-test: FAIL ({len(bad)} of {len(cases)} cases)")
        return 1
    print(f"self-test: PASS ({len(cases)} cases)")
    return 0


def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=(__doc__ or "").split("\n\n")[0])
    parser.add_argument("--root", default=str(ROOT))
    parser.add_argument("--self-test", action="store_true")
    args = parser.parse_args(argv)
    return self_test() if args.self_test else check(Path(args.root).resolve())


if __name__ == "__main__":
    sys.exit(main())
