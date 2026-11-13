#!/usr/bin/env python3
"""Fail when a document restates a count that a record already holds (#328).

One home per number: assertion counts, grid row counts, contract counts and
the number of open release blockers live in their records and are generated
into the README assurance block. A document links to them instead of copying
them, because a copy goes stale silently: the audit of 2026-11-11 found the
same assertion count wrong in four places.

Scope is the top-level guides, docs/*.md. Dated research records under
docs/research/ and the CHANGELOG keep historical numbers on purpose.

A finding is either:
  * a current count from its record, written as a standalone number:
      - expected_assertions in .github/excel-evidence-policy.json;
      - expected_assertions in .github/excel-study-policy.json;
      - the data rows of the main grid and of the holdout grid;
  * a restating phrase whatever its value, which also catches a stale copy:
    "currently <number>", "<number> open blockers", "(<number> open)",
    "<number>-row".

The evidence-rules table of #328 must also be identical in CONTRIBUTING.md
and docs/EXCEL_CERTIFICATION.md, between the markers
<!-- BEGIN evidence rules --> and <!-- END evidence rules -->.

Usage: python3 tools/check_restated_counts.py [--root .]
"""

from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
GRIDS = {
    "main-grid rows": "benchmark/probability_accuracy_grid.csv",
    "holdout rows": "benchmark/holdout/holdout_grid.csv",
}
POLICIES = {
    "regression assertions": ".github/excel-evidence-policy.json",
    "study assertions": ".github/excel-study-policy.json",
}
PHRASES = (
    re.compile(r"\bcurrently\s+\d[\d,]*(?![\d.Ee])"),
    re.compile(r"\b\d+\s+(?:registered\s+)?open\s+blockers\b"),
    re.compile(r"\(\d+\s+open\)"),
    re.compile(r"(?<![\w.#])\d[\d,]*-row\b"),
)
TABLE_BEGIN = "<!-- BEGIN evidence rules -->"
TABLE_END = "<!-- END evidence rules -->"
TABLE_HOMES = ("CONTRIBUTING.md", "docs/EXCEL_CERTIFICATION.md")


def known_counts(root: Path) -> dict[str, int]:
    counts = {}
    for label, rel in POLICIES.items():
        counts[label] = int(json.loads((root / rel).read_text(encoding="utf-8"))
                            ["expected_assertions"])
    for label, rel in GRIDS.items():
        lines = (root / rel).read_text(encoding="utf-8").splitlines()
        counts[label] = sum(1 for line in lines[1:] if line.strip())
    return counts


def number_pattern(value: int) -> re.Pattern[str]:
    plain, grouped = str(value), f"{value:,}"
    alternatives = "|".join(re.escape(t) for t in dict.fromkeys((grouped, plain)))
    # Not part of a longer number, an issue reference or a decimal. A count in
    # an inline code span still counts; a commit reference such as `5987abc`
    # does not, because a word character follows the digits.
    return re.compile(rf"(?<![\w.#,])(?:{alternatives})(?!\w|[.,]\d)")


def table(text: str) -> str | None:
    if TABLE_BEGIN not in text or TABLE_END not in text:
        return None
    return text.split(TABLE_BEGIN, 1)[1].split(TABLE_END, 1)[0].strip()


def findings(root: Path) -> list[str]:
    found = []
    counts = known_counts(root)
    patterns = {label: number_pattern(v) for label, v in counts.items() if v >= 100}
    for path in sorted((root / "docs").glob("*.md")):
        rel = path.relative_to(root).as_posix()
        for lineno, line in enumerate(path.read_text(encoding="utf-8").splitlines(), 1):
            for label, pattern in patterns.items():
                for m in pattern.finditer(line):
                    found.append(f"{rel}:{lineno}: restates {label} ({m.group(0)}); "
                                 "link the record instead")
            for pattern in PHRASES:
                for m in pattern.finditer(line):
                    found.append(f"{rel}:{lineno}: restating phrase {m.group(0)!r}; "
                                 "link the record instead")
    tables = {}
    for rel in TABLE_HOMES:
        path = root / rel
        body = table(path.read_text(encoding="utf-8")) if path.exists() else None
        if body is None:
            found.append(f"{rel}: evidence-rules table markers missing")
        tables[rel] = body
    bodies = {b for b in tables.values() if b is not None}
    if len(bodies) > 1:
        found.append("the evidence-rules tables in " + " and ".join(TABLE_HOMES)
                     + " differ")
    return found


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    ap.add_argument("--root", type=Path, default=ROOT)
    a = ap.parse_args(argv)
    found = findings(a.root.resolve())
    for f in found:
        print("FAIL: " + f)
    if found:
        return 1
    print("PASS: no record-held count is restated in docs/*.md; the evidence-rules "
          "tables agree")
    return 0


if __name__ == "__main__":
    sys.exit(main())
