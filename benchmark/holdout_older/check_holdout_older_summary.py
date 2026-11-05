"""
Fail when the committed holdout_older_summary.md is stale (#164).

Renders the summary from the committed grid and contract registry with the
analyzer's own render() and compares it with the committed file. Any
difference fails, with a unified diff. Nothing is written.

This checks that the record is current, not that the contracts hold: a summary
that records FAIL verdicts honestly passes, and one that reports PASS against a
registry it no longer matches fails. The verdicts themselves are in the
summary, and the analyzer still exits non-zero while any contract fails.

Before this check existed the summary drifted twice without notice: 43d2415
retightened two Gamma contracts and 1fac483 retired LogGamma.all.output, and
the committed file kept the verdicts of 2026-07-21.

Usage:
    python check_holdout_older_summary.py
    python check_holdout_older_summary.py --summary S --grid G --contracts C
Exit 0 when current, 1 when stale or unreadable.
"""
import argparse
import difflib
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import analyze_holdout_older as A  # noqa: E402


def check(summary, grid, contracts):
    """Return (ok, message)."""
    try:
        expected, results, _failing = A.render(*A.load(grid, contracts))
    except (OSError, KeyError, ValueError, ArithmeticError) as exc:
        return False, f"cannot render the summary: {exc!r}"
    try:
        with open(summary, encoding="utf-8", newline="") as f:
            committed = f.read()
    except OSError as exc:
        return False, f"cannot read {summary}: {exc}"
    nfail = sum(1 for _, _, _, v in results if v == "FAIL")
    if committed == expected:
        return True, (f"{os.path.basename(summary)} matches a re-run "
                      f"({len(results)} contracts, {nfail} FAIL recorded)")
    diff = "".join(difflib.unified_diff(
        committed.splitlines(keepends=True), expected.splitlines(keepends=True),
        "committed", "re-run", n=0))
    return False, (f"{os.path.basename(summary)} is stale against the current grid and "
                   f"registry; run analyze_holdout_older.py and commit the result\n{diff}")


def main():
    ap = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    ap.add_argument("--summary", default=A.SUMMARY)
    ap.add_argument("--grid", default=A.GRID)
    ap.add_argument("--contracts", default=A.CONTRACTS)
    a = ap.parse_args()
    ok, message = check(a.summary, a.grid, a.contracts)
    print(("PASS: " if ok else "FAIL: ") + message)
    raise SystemExit(0 if ok else 1)


if __name__ == "__main__":
    main()
