"""
Duplicate-identity guard for the committed accuracy grids (#17).

Every committed grid row must have a unique canonical identity: function, the
four arguments as IEEE-754 binary64 bit patterns, regime and evidence_set. A
duplicate is ambiguous evidence. Two copies of one row can carry different
references or observations, nothing says which is authoritative, and a
contract fed both counts one point twice.

The writers refuse duplicates on their own paths (generate_reference_values.py,
promote_grid_rows.py, migrate_references.py), but a grid can also change by a
hand edit or a fresh Excel export. This check runs on the committed files
themselves so that no route can land a duplicate unnoticed.

It also requires the canonical column set, so a grid rewritten by a tool with a
different schema fails here rather than further down the gate.

Run: python3 check_grid_keys.py   (exit 0 = pass, nonzero = fail)
"""
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import _grid_safety as gs  # noqa: E402


def check(path):
    """Problems with one grid file; an empty list means it passes."""
    if not os.path.exists(path):
        return [f"{path} is missing"]
    rows = gs.read_grid(path)
    problems = []
    with open(path, newline="", encoding="utf-8") as fh:
        header = fh.readline().rstrip("\r\n").split(",")
    if header != gs.GRID_FIELDS:
        problems.append(f"{path}: columns {header} differ from the canonical "
                        f"schema {gs.GRID_FIELDS}")
        return problems
    dup = gs.duplicate_keys(rows)
    for k, n in sorted(dup.items()):
        problems.append(f"{path}: x{n}  {gs.describe_key(k)}")
    return problems


def main(paths=None):
    paths = paths or [gs.COMMITTED_GRID, gs.COMMITTED_HOLDOUT]
    problems, total = [], 0
    for p in paths:
        problems += check(p)
        if os.path.exists(p):
            total += len(gs.read_grid(p))
    if problems:
        print("FAIL: grid canonical-key guard")
        for p in problems:
            print("  - " + p)
        return 1
    print(f"PASS: grid canonical keys unique ({len(paths)} grid(s), {total} rows)")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:] or None))
