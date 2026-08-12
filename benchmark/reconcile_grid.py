"""
Reconcile probability_accuracy_grid.csv against its declared generator.

REPORT ONLY. This script never writes the grid. It exists because the grid holds
2030 Excel observations that cannot be regenerated, and because
generate_reference_values.py no longer reproduces it in either direction (issue
#17). Nothing may write that file until the divergence is understood.

Key. `function` + the four argument columns does NOT identify a row: 11 keys are
duplicated. The correct key adds `regime` and `evidence_set`, which resolves 8 of
them legitimately - the same argument promoted into a study evidence set
alongside its main-grid appearance, or two outputs discriminated by regime as
Lognormal_ParametersFromMeanStdDev already does with param_meanlog /
param_stddevlog. The residual duplicates are reported separately.

Classification:
  MATCH                       present both sides, all compared fields equal
  REFERENCE_DIFFERENCE        present both sides, reference differs
  METADATA_DIFFERENCE         present both sides, claim/metric/expected_error differs
  GRID_ONLY_EXPECTED          in the grid, declares a non-main evidence_set
  GRID_ONLY_UNEXPLAINED       in the grid, tagged `main grid`, generator omits it
  GENERATOR_ONLY_UNEXPLAINED  generator emits it, grid omits it

Where a GRID_ONLY_UNEXPLAINED row's arguments appear in a study folder's own
grid, that folder is reported as a probable origin. It is a lead, not a
declaration: study grids use a different schema, so the match is numeric only.
"""
import argparse, csv, glob, importlib.util, os, sys
from collections import Counter, defaultdict

import mpmath as mp

HERE = os.path.dirname(os.path.abspath(__file__))
COMPARE = ("claim", "metric", "expected_error")


def key(r):
    return (r["function"], r["arg1"], r["arg2"], r["arg3"], r["arg4"],
            r["regime"], r.get("evidence_set", ""))


def load_generator(path):
    spec = importlib.util.spec_from_file_location("gen", path)
    m = importlib.util.module_from_spec(spec)
    sys.modules["gen"] = m
    spec.loader.exec_module(m)
    mp.mp.dps = 50
    rows = m.build_rows()
    from _contract_eval import predicted_expected_error
    for r in rows:
        r["expected_error"] = "1" if predicted_expected_error(
            r.get("function", ""), r.get("arg2", ""), r.get("arg3", "")) else ""
    return rows


# Top-level benchmark CSVs are not studies; indexing them would attribute every
# row to the very file being reconciled.
_NOT_A_STUDY = {"benchmark", ""}


def study_arg_index(root):
    """Numeric argument tuples appearing in each study folder's own grid."""
    idx = defaultdict(set)
    for path in glob.glob(os.path.join(root, "*", "*.csv")):
        folder = os.path.basename(os.path.dirname(path))
        if folder in _NOT_A_STUDY or folder.startswith("__"):
            continue
        try:
            with open(path, newline="") as f:
                for row in csv.DictReader(f):
                    vals = []
                    for k in ("arg1", "arg2", "arg3", "x", "z"):
                        v = (row.get(k) or "").strip()
                        if v:
                            try: vals.append(float(v))
                            except ValueError: pass
                    if vals:
                        idx[folder].add(tuple(vals))
        except Exception:
            continue
    return idx


def probable_origin(r, idx):
    vals = []
    for k in ("arg1", "arg2", "arg3"):
        v = (r.get(k) or "").strip()
        if v:
            try: vals.append(float(v))
            except ValueError: pass
    if not vals:
        return ""
    hits = [f for f, s in idx.items() if tuple(vals) in s]
    return ";".join(sorted(hits))


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--grid", default=os.path.join(HERE, "probability_accuracy_grid.csv"))
    ap.add_argument("--generator", default=os.path.join(HERE, "generate_reference_values.py"))
    # The study folders are benchmark/<study>/, i.e. directly under HERE.
    # Defaulting to dirname(HERE) made the glob match benchmark/*.csv - the main
    # grid itself - and every row then "matched" that pseudo-folder.
    ap.add_argument("--study-root", default=HERE)
    ap.add_argument("--out", default=os.path.join(HERE, "grid_reconciliation.csv"))
    a = ap.parse_args()

    grid = list(csv.DictReader(open(a.grid)))
    gen = load_generator(a.generator)

    gk, ggk = Counter(key(r) for r in grid), Counter(key(r) for r in gen)
    dup_grid = {k: v for k, v in gk.items() if v > 1}
    dup_gen = {k: v for k, v in ggk.items() if v > 1}
    print(f"grid rows {len(grid)}   generator rows {len(gen)}")
    print(f"duplicate keys - grid {len(dup_grid)}, generator {len(dup_gen)}")
    for k, v in sorted(dup_grid.items()):
        print(f"   x{v}  {k[0]}  args=({','.join(x for x in k[1:5] if x)})  regime={k[5]}  set={k[6]}")

    G = {key(r): r for r in grid}
    N = {key(r): r for r in gen}
    idx = study_arg_index(a.study_root)
    # Print what was actually indexed. A wrong --study-root silently produces a
    # plausible-looking attribution rather than an error: pointing it one level
    # too high makes the glob match benchmark/*.csv, so the main grid indexes
    # itself as a pseudo-folder and every row "matches" it.
    print(f"\nstudy folders indexed under {a.study_root}: {len(idx)}")
    print("  " + (", ".join(sorted(idx)) if idx else "NONE - check --study-root"))

    out, counts = [], Counter()
    for k in sorted(set(G) | set(N)):
        g, n = G.get(k), N.get(k)
        origin = ""
        if g is not None and n is not None:
            if g["reference"] != n["reference"]:
                cls = "REFERENCE_DIFFERENCE"
            elif any(g.get(c, "") != n.get(c, "") for c in COMPARE):
                cls = "METADATA_DIFFERENCE"
            else:
                cls = "MATCH"
            origin = "generate_reference_values"
        elif g is not None:
            if g.get("evidence_set", "") not in ("", "main grid"):
                cls = "GRID_ONLY_EXPECTED"; origin = g["evidence_set"]
            else:
                cls = "GRID_ONLY_UNEXPLAINED"; origin = probable_origin(g, idx)
        else:
            cls = "GENERATOR_ONLY_UNEXPLAINED"; origin = "generate_reference_values"
        counts[cls] += 1
        src = g or n
        out.append({
            "key": "|".join(k), "function": k[0],
            "arguments": ";".join(x for x in k[1:5] if x),
            "regime": k[5], "evidence_set": k[6],
            "generator_state": "present" if n is not None else "absent",
            "grid_state": "present" if g is not None else "absent",
            "observed": "filled" if (g and (g.get("observed_vba") or "").strip()) else "",
            "classification": cls, "origin": origin,
        })

    with open(a.out, "w", newline="") as f:
        w = csv.DictWriter(f, fieldnames=list(out[0].keys()), lineterminator="\n")
        w.writeheader(); w.writerows(out)

    print(f"\nwrote {a.out}: {len(out)} keys\n")
    for c, n in counts.most_common():
        print(f"  {c:30s} {n:5d}")
    print("\nGRID_ONLY_UNEXPLAINED by function:")
    un = [r for r in out if r["classification"] == "GRID_ONLY_UNEXPLAINED"]
    for fn, n in Counter(r["function"] for r in un).most_common(12):
        withorigin = sum(1 for r in un if r["function"] == fn and r["origin"])
        print(f"   {n:4d}  {fn:36s} probable origin found for {withorigin}")
    # Attribution strength, stated honestly: a numeric argument tuple matching a
    # study folder's own grid is a lead, not a declaration. With 18 folders in
    # play, common arguments collide. If nothing is uniquely attributed, origin
    # cannot be reconstructed after the fact and must instead be WRITTEN by
    # whatever promotes a row.
    uniq = sum(1 for r in un if r["origin"] and ";" not in r["origin"])
    multi = Counter(len(r["origin"].split(";")) for r in un if r["origin"])
    obs = sum(1 for r in un if r["observed"])
    print(f"\n  {len(un)} unexplained rows, {obs} carrying an observation.")
    print(f"  uniquely attributed to one study folder: {uniq}")
    for k in sorted(multi):
        if k > 1:
            print(f"    {multi[k]:4d} row(s) matched {k} folders - coincidence, not evidence")
    if uniq == 0 and un:
        print("  => origin CANNOT be inferred retrospectively. It must be written")
        print("     by whatever promotes a row. See issue #17 finding F.")


if __name__ == "__main__":
    main()
