#!/usr/bin/env python3
"""Negative controls for render_contract_table.py's generated README block."""
from __future__ import annotations

import tempfile
from pathlib import Path

import render_contract_table as R

failures: list[str] = []
BEGIN = "<!-- BEGIN generated: accuracy_contracts.csv via render_contract_table.py. Do not hand-edit. -->"
END = "<!-- END generated -->"

with tempfile.TemporaryDirectory() as td:
    root = Path(td)
    readme = root / "README.md"

    # A stale hand-edited block must be deterministically replaced.
    readme.write_text("before\n" + BEGIN + "\n\nSTALE CONTENT\n\n" + END + "\nafter\n",
                      encoding="utf-8")
    before = readme.read_text(encoding="utf-8")
    R.write_into_readme(str(readme))
    after = readme.read_text(encoding="utf-8")
    if before == after:
        failures.append("stale generated block was not replaced")
    if "STALE CONTENT" in after:
        failures.append("stale content survived regeneration")
    if R.render() not in after:
        failures.append("canonical rendered table was not inserted")
    if not after.startswith("before\n") or not after.endswith("after\n"):
        failures.append("text outside generated markers was changed")

    # Missing markers must fail rather than overwriting an arbitrary README.
    missing = root / "missing.md"
    missing.write_text("no generated markers here\n", encoding="utf-8")
    try:
        R.write_into_readme(str(missing))
    except SystemExit:
        pass
    else:
        failures.append("missing generated markers did not fail closed")

if failures:
    print("FAIL: generated contract-table fixtures")
    for failure in failures:
        print("  - " + failure)
    raise SystemExit(1)
print("PASS: generated contract-table fixtures (stale block replaced, marker loss fails, surrounding README preserved)")
