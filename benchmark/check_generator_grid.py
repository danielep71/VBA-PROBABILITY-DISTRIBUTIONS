"""
The reference generator and the committed main grid describe the same rows.

Every row generate_reference_values.py emits must be in
probability_accuracy_grid.csv under the same canonical key (function, four
binary64 argument bit patterns, regime, evidence_set) with the same reference,
claim, metric and expected_error. A row the generator emits and the grid lacks,
or holds differently, fails (#349).

The reverse is not required. Many committed rows were promoted from study
folders the generator knows nothing about, so a grid row the generator does not
emit is counted and kept.

Read only. A difference is settled by deciding which side is right:
correct the generator, or change the grid through promote_grid_rows.py with a
counted acknowledgement and a reason.

Run: python3 check_generator_grid.py   (exit 0 = agree, 1 = differ)
"""
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import _grid_safety as gs  # noqa: E402

COMPARED = ("reference", "claim", "metric", "expected_error")


def compare(gen_rows, grid_rows):
    """(missing, different, grid_only): generator rows absent from the grid,
    (key, fields) for rows held differently, and the count of grid rows the
    generator does not emit."""
    grid = {gs.canonical_key(r): r for r in grid_rows}
    seen = set()
    missing, different = [], []
    for row in gen_rows:
        k = gs.canonical_key(row)
        seen.add(k)
        cur = grid.get(k)
        if cur is None:
            missing.append(k)
            continue
        fields = [c for c in COMPARED if cur.get(c, "") != row.get(c, "")]
        if fields:
            different.append((k, fields))
    return missing, different, sum(1 for k in grid if k not in seen)


def main():
    from promote_grid_rows import load_generator
    gen = load_generator(os.path.join(HERE, "generate_reference_values.py"))
    grid = gs.read_grid(gs.COMMITTED_GRID)
    for name, rows in (("generator", gen), ("grid", grid)):
        dup = gs.duplicate_keys(rows)
        if dup:
            print(f"FAIL: {len(dup)} duplicate canonical key(s) in the {name}")
            for k in sorted(dup)[:10]:
                print(f"   {gs.describe_key(k)}")
            return 1
    missing, different, grid_only = compare(gen, grid)
    print(f"generator rows {len(gen)}; identical in the grid "
          f"{len(gen) - len(missing) - len(different)}; absent from the grid "
          f"{len(missing)}; held differently {len(different)}; grid rows the "
          f"generator does not emit {grid_only} (kept)")
    for k in missing[:20]:
        print(f"   absent     {gs.describe_key(k)}")
    for k, fields in different[:20]:
        print(f"   differs    {gs.describe_key(k)}  fields={fields}")
    if missing or different:
        print("FAIL: the generator and the committed grid disagree. Correct the "
              "generator, or change the grid with promote_grid_rows.py and a "
              "stated reason.")
        return 1
    print("PASS: every generator row is in the committed grid, unchanged")
    return 0


if __name__ == "__main__":
    sys.exit(main())
