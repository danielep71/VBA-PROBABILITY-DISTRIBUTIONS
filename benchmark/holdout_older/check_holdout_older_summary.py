"""
Fail when the committed holdout_older_summary.md is stale (#164), or when a
grid row is neither scored nor marked characterization only (#171).

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

A current summary still fails when a grid row has no disposition: no contract
with a threshold scores it, and no contract with status `characterization_only`
covers its function and regime. 1fac483 (2026-08-13) left three LogGamma rows
in that state; after #164 the summary listed them, and nothing failed (#171).

Usage:
    python check_holdout_older_summary.py
    python check_holdout_older_summary.py --summary S --grid G --contracts C
Exit 0 when current with every row disposed, 1 when stale, unreadable or a
row has no disposition.
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
        expected, results, _failing, unclaimed, _above = A.render(*A.load(grid, contracts))
    except (OSError, KeyError, ValueError, ArithmeticError) as exc:
        return False, f"cannot render the summary: {exc!r}"
    try:
        with open(summary, encoding="utf-8", newline="") as f:
            committed = f.read()
    except OSError as exc:
        return False, f"cannot read {summary}: {exc}"
    nfail = sum(1 for _, _, _, v in results if v == "FAIL")
    if committed == expected:
        loose = A.undisposed(unclaimed)
        if loose:
            listed = "\n".join(f"  {r['function']}({A.args_text(r)}) in regime {r['regime']}"
                                for r in loose)
            return False, (f"{len(loose)} grid row(s) are neither scored by a contract nor "
                           "marked characterization only; move each into the regime of a "
                           "contract that covers it, or cover its function and regime with "
                           f"a characterization_only contract\n{listed}")
        marked = len(unclaimed)
        return True, (f"{os.path.basename(summary)} matches a re-run "
                      f"({len(results)} contracts, {nfail} FAIL recorded; "
                      f"{marked} row(s) marked characterization only)")
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
