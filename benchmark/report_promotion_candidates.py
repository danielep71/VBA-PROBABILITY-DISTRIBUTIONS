"""
Candidate row report for the #12/#15 promotion. Report only; writes nothing.

Answers, before promote_grid_rows.py --allow-add is run:
  * which rows would be added, by contract regime;
  * whether every candidate reference matches the Step-2-audited study
    reference at the same binary64 key;
  * whether any candidate falls outside its contract's stated domain;
  * whether any candidate key is duplicated, or already in the main grid.
"""
import csv, importlib.util, struct, sys
from collections import Counter, defaultdict

import mpmath as mp

NRM = 3.854839696505424e-308            # PROB_MIN_NORMAL / EulerGamma
SEAM = 0.25

DOMAIN = {
    ("LogGamma1p", "small"): (NRM, SEAM),
    ("LogGamma1p", "series_seam"): (0.24999999999999994, 0.25000000000000006),
}

STUDIES = {
    "LogGamma": ("loggamma_regimes_study/loggamma_regimes_grid.csv", "LogGamma"),
    "LogGamma1p": ("loggamma1p_study/loggamma1p_grid.csv", "LogGamma1p"),
}


def bits(t):
    t = (t or "").strip()
    if t == "":
        return ""
    try:
        return struct.pack(">d", float(t)).hex()
    except (ValueError, OverflowError):
        return "?" + t


def main():
    spec = importlib.util.spec_from_file_location("gen", "generate_reference_values.py")
    g = importlib.util.module_from_spec(spec); sys.modules["gen"] = g
    spec.loader.exec_module(g)
    mp.mp.dps = 50
    gen = g.build_rows()
    grid = list(csv.DictReader(open("probability_accuracy_grid.csv")))
    in_grid = {(r["function"], bits(r["arg1"]), r["regime"]) for r in grid}

    study = {}
    for fn, (path, qty) in STUDIES.items():
        for r in csv.DictReader(open(path)):
            if r["quantity"] == qty:
                study[(fn, bits(r["arg1"]))] = r["reference"]

    cand = [r for r in gen
            if (r["function"] == "LogGamma" and r["regime"] != "all")
            or r["function"] == "LogGamma1p"]
    groups = defaultdict(list)
    for r in cand:
        groups[(r["function"], r["regime"])].append(r)

    print(f"{'contract / regime':40s} {'rows':>5s} {'min arg1':>14s} {'max arg1':>14s} "
          f"{'in grid':>8s} {'append':>7s}")
    problems = []
    for k in sorted(groups):
        rows = groups[k]
        xs = sorted(float(r["arg1"]) for r in rows)
        already = sum(1 for r in rows
                      if (r["function"], bits(r["arg1"]), r["regime"]) in in_grid)
        print(f"{k[0] + '.' + k[1]:40s} {len(rows):5d} {xs[0]:14.4e} {xs[-1]:14.4e} "
              f"{already:8d} {len(rows) - already:7d}")
        lo_hi = DOMAIN.get(k)
        if lo_hi:
            out = [x for x in xs if not (lo_hi[0] <= x <= lo_hi[1])]
            if out:
                problems.append(f"{k}: {len(out)} row(s) outside domain {lo_hi}: {out[:4]}")

    print("\n--- reference provenance ---")
    matched = mismatched = absent = 0
    worst = None
    for r in cand:
        sk = (r["function"], bits(r["arg1"]))
        sref = study.get(sk)
        if sref is None:
            absent += 1
            problems.append(f"no study row for {r['function']} arg1={r['arg1']}")
            continue
        # The study serialises 34 significant digits, the grid 25. Equality of
        # the decoded values is therefore the wrong test: it would flag pure
        # truncation as a difference. The candidate must agree with the audited
        # study value to within the GRID's own serialisation, half an ulp at 25
        # digits, and anything beyond that is a real numerical difference.
        S, G = mp.mpf(sref), mp.mpf(r["reference"])
        d = abs(S - G)
        rel = d / abs(S) if S != 0 else d
        if rel <= mp.mpf(10) ** -24:
            matched += 1
        else:
            if worst is None or rel > worst[0]:
                worst = (rel, r["function"], r["arg1"], sref, r["reference"])
            mismatched += 1
    print(f"  matching the audited study value to 25-digit serialisation: {matched}")
    print(f"  differing: {mismatched}    no study row: {absent}")
    if worst:
        print(f"  largest difference {float(worst[0]):.2e} at {worst[1]} arg1={worst[2]}")
        print(f"     study     {worst[3]}")
        print(f"     generator {worst[4]}")

    print("\n--- key hygiene ---")
    kc = Counter((r["function"], bits(r["arg1"]), r["regime"]) for r in cand)
    dup = {k for k, v in kc.items() if v > 1}
    print(f"  duplicate candidate keys: {len(dup)}")
    if dup:
        problems.append(f"{len(dup)} duplicate candidate key(s)")

    print("\n--- excluded from promotion (deliberate) ---")
    ex = Counter()
    for r in gen:
        if r["function"] == "LogGamma" and r["regime"] == "all":
            ex["LogGamma.all (existing main-grid rows)"] += 1
    for path, qty, keep in (("loggamma1p_study/loggamma1p_grid.csv", "LogGamma1p",
                             {"small", "series_seam"}),
                            ("loggamma_regimes_study/loggamma_regimes_grid.csv", "LogGamma",
                             {"small_positive", "near_zero", "exact_zero", "general"})):
        for r in csv.DictReader(open(path)):
            if r["quantity"] == qty and r["regime"] not in keep:
                ex[f"{qty}.{r['regime']} (study only)"] += 1
    for k, v in sorted(ex.items()):
        print(f"  {k:52s} {v:4d}")

    print()
    if problems:
        print("PROBLEMS:")
        for p in problems:
            print("  -", p)
    else:
        print("No problems. Candidate set is ready for --allow-add review.")


if __name__ == "__main__":
    main()
