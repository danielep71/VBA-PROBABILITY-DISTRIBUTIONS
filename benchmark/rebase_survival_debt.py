"""
Re-base the survival references registered as debt in survival_reference_debt.json
(#159), once the bindings that blocked them are rebuilt (#81).

Every debt entry names one cell: the grid, the function, the argument tokens and
the committed reference text. For each one this tool takes the new reference
from the source that owns the row, never from the audit's own oracle:

  probability_accuracy_grid.csv
      the study grid the row is a copy of (beta_f_unbalanced, cdf_large_shape),
      whose fixed generators re-based it in #165; otherwise the main generator,
      generate_reference_values.py;
  holdout/holdout_grid.csv
      holdout/generate_holdout.py;
  holdout_older/holdout_older_grid.csv
      holdout_older/gen_holdout_older.py;
  survival_boundary/survival_boundary_grid.csv
      survival_boundary/generate_survival_boundary.py.

A row is matched to its source by function and the binary64 values of its
arguments. Then the new value must agree with the independent direct-route
oracle (_survival_oracle.py) to the 25-digit floor that check_survival_references.py
enforces. The tool refuses, writing nothing, when any entry has no source row,
has two, or its new value misses the floor.

With --write it changes the reference cell of each entry and nothing else. It
then checks that every other byte of each grid is unchanged, appends one line per
change to study_reference_rebase.csv, and deletes the entries from the registry.
Report-only without --write.

A grid whose bytes a manifest binds (the main grid, the holdout) must then be
re-exported from Excel and rebound in the same change, as PROVENANCE.md requires.

    python rebase_survival_debt.py            # report
    python rebase_survival_debt.py --write    # patch, record, clear the debt
"""
import argparse
import csv
import importlib.util
import io
import json
import os
import sys

import mpmath as mp

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import _survival_oracle as oracle  # noqa: E402
from _study_merge import RECORD, RECORD_FIELDS  # noqa: E402

DEBT = os.path.join(HERE, "survival_reference_debt.json")
MAIN = "probability_accuracy_grid.csv"
COPIES = ("beta_f_unbalanced/beta_f_unbalanced_grid.csv", "cdf_large_shape/cdf_large_shape_grid.csv")


class RebaseError(RuntimeError):
    pass


def key(row):
    fn = row.get("function") or row.get("quantity")
    return fn, tuple(float(str(row[c])) for c in ("arg1", "arg2", "arg3", "arg4")
                     if str(row.get(c) or "").strip())


def tokens(row):
    return [row[c].strip() for c in ("arg1", "arg2", "arg3", "arg4") if (row.get(c) or "").strip()]


def _module(path, name):
    spec = importlib.util.spec_from_file_location(name, os.path.join(HERE, path))
    module = importlib.util.module_from_spec(spec)
    sys.modules[name] = module
    saved = sys.argv
    sys.argv = [path]
    try:
        spec.loader.exec_module(module)
    finally:
        sys.argv = saved
    return module


def _read(path):
    with open(os.path.join(HERE, path), newline="", encoding="utf-8") as f:
        text = f.read()
    return text, list(csv.DictReader(io.StringIO(text)))


def source_rows(grid):
    """(source name, rows) pairs that own the rows of a grid, in order of authority."""
    if grid == MAIN:
        out = [(path, _read(path)[1]) for path in COPIES]
        generator = _module("generate_reference_values.py", "gen_main")
        mp.mp.dps = 50
        out.append(("generate_reference_values.py", generator.build_rows()))
        return out
    if grid == "holdout/holdout_grid.csv":
        return [("holdout/generate_holdout.py", _module("holdout/generate_holdout.py", "gen_holdout").build())]
    if grid == "holdout_older/holdout_older_grid.csv":
        return [("holdout_older/gen_holdout_older.py", _module("holdout_older/gen_holdout_older.py", "gen_older").R)]
    if grid == "survival_boundary/survival_boundary_grid.csv":
        return [("survival_boundary/generate_survival_boundary.py",
                 _module("survival_boundary/generate_survival_boundary.py", "gen_boundary").build())]
    raise RebaseError(f"no source is known for {grid}")


# Main-grid rows whose arguments the generator no longer emits bit for bit: the
# four tiny_unbalanced F_Survival rows carry arguments one ulp from what
# build_tiny_unbalanced_rows computes today. Their observations were taken at the
# committed arguments, so the reference is the generator's own function, at 50
# digits, evaluated at those arguments' binary64 values, and written as the
# generator writes it.
AT_GRID_ARGUMENTS = {("F_Survival", "tiny_unbalanced"): "_tu_f_sf"}


def at_grid_arguments(function, row):
    generator = sys.modules.get("gen_main") or _module("generate_reference_values.py", "gen_main")
    evaluate = getattr(generator, AT_GRID_ARGUMENTS[(function, row["regime"])])
    with mp.workdps(50):
        value = evaluate(*[mp.mpf(float(t)) for t in tokens(row)])
        return ("generate_reference_values.py " + AT_GRID_ARGUMENTS[(function, row["regime"])]
                + " at the committed arguments", mp.nstr(value, 25))


def plan(entries):
    """Every change, or RebaseError naming every entry that cannot be re-based."""
    changes, problems = [], []
    by_grid = {}
    for e in entries:
        by_grid.setdefault(e["grid"], []).append(e)
    for grid, mine in sorted(by_grid.items()):
        _, rows = _read(grid)
        sources = source_rows(grid)
        for e in mine:
            hits = [(i, r) for i, r in enumerate(rows)
                    if r["function"] == e["function"] and tokens(r) == e["arguments"]
                    and r["reference"].strip() == e["reference"]]
            if len(hits) != 1:
                problems.append(f"{grid} {e['function']} {e['arguments']}: {len(hits)} committed rows match")
                continue
            index, row = hits[0]
            k = key(row)
            found = []
            for name, srows in sources:
                found += [(name, str(s["reference"]).strip()) for s in srows if key(s) == k]
                if found:
                    break                       # the first source that owns the row decides
            if not found and grid == MAIN and (e["function"], row.get("regime")) in AT_GRID_ARGUMENTS:
                found = [at_grid_arguments(e["function"], row)]
            if len(found) != 1:
                problems.append(f"{grid} {e['function']} {e['arguments']}: {len(found)} source rows")
                continue
            name, new = found[0]
            value, _, route = oracle.certify(e["function"], e["arguments"])
            if not oracle.agrees(new, value):
                problems.append(f"{grid} {e['function']} {e['arguments']}: the {name} value {new} "
                                f"misses the 25-digit floor against the oracle")
                continue
            with mp.workdps(80):
                change = abs(mp.mpf(new) - mp.mpf(e["reference"])) / abs(mp.mpf(new))
            changes.append({"grid": grid, "row": index + 2, "function": e["function"],
                            "arguments": e["arguments"], "regime": row.get("regime", ""),
                            "old": e["reference"], "new": new, "source": name,
                            "relative_change": float(change)})
    if problems:
        raise RebaseError("refusing; nothing written:\n  - " + "\n  - ".join(problems))
    return changes


def apply(changes):
    for grid in sorted({c["grid"] for c in changes}):
        path = os.path.join(HERE, grid)
        text, rows = _read(grid)
        lines = text.split("\n")
        header = next(csv.reader([lines[0]]))
        column = header.index("reference")
        for c in (c for c in changes if c["grid"] == grid):
            cells = next(csv.reader([lines[c["row"] - 1]]))
            if cells[column] != c["old"]:
                raise RebaseError(f"{grid} line {c['row']}: reference is {cells[column]}, not {c['old']}")
            old_line = lines[c["row"] - 1]
            if old_line.count(c["old"]) != 1:
                raise RebaseError(f"{grid} line {c['row']}: the reference text is not unique on the line")
            lines[c["row"] - 1] = old_line.replace(c["old"], c["new"])
        patched = "\n".join(lines)
        after = list(csv.DictReader(io.StringIO(patched)))
        moved = {c["row"] - 2 for c in changes if c["grid"] == grid}
        for i, (a, b) in enumerate(zip(rows, after)):
            other = {k: v for k, v in a.items() if not (i in moved and k == "reference")}
            if any(b[k] != v for k, v in other.items()):
                raise RebaseError(f"{grid} row {i + 2}: a cell other than the re-based reference changed")
        if len(after) != len(rows):
            raise RebaseError(f"{grid}: the row count changed")
        with open(path, "w", newline="", encoding="utf-8") as f:
            f.write(patched)
    exists = os.path.exists(RECORD)
    with open(RECORD, "a", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=RECORD_FIELDS, lineterminator="\n")
        if not exists:
            writer.writeheader()
        for c in changes:
            writer.writerow({"grid": c["grid"], "row": c["row"], "function": c["function"],
                             "arguments": "|".join(c["arguments"]), "regime": c["regime"],
                             "old_reference": c["old"], "new_reference": c["new"],
                             "relative_change": f"{c['relative_change']:.3e}", "issue": "#159"})


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    ap.add_argument("--write", action="store_true", help="patch the grids, record, clear the debt")
    a = ap.parse_args(argv)
    with open(DEBT, encoding="utf-8") as f:
        debt = json.load(f)
    try:
        changes = plan(debt["entries"])
    except (RebaseError, oracle.OracleError) as exc:
        print(f"REFUSED: {exc}")
        return 1
    for c in changes:
        print(f"  {c['grid']}:{c['row']} {c['function']}({', '.join(c['arguments'])}) "
              f"{c['old']} -> {c['new']}  [{c['relative_change']:.2e}, {c['source']}]")
    print(f"{len(changes)} reference(s) re-based from their sources; every one meets the floor")
    if not a.write:
        print("report only; re-run with --write")
        return 0
    apply(changes)
    debt["entries"] = []
    with open(DEBT, "w", encoding="utf-8", newline="\n") as f:
        json.dump(debt, f, indent=2)
        f.write("\n")
    print(f"wrote the grids, appended {len(changes)} line(s) to {os.path.basename(RECORD)}, "
          f"and cleared {os.path.basename(DEBT)}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
