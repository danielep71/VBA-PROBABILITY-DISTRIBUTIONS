"""
Analyze the independent holdout: do the provisional thresholds hold on fresh data?

Joins holdout observations to the regime-specific contracts by (function, regime),
computes each contract's measure, and reports worst error vs threshold with margin.
If every regime-specific contract passes on this unseen data, the thresholds
generalise and may be frozen (measured provisional -> validated and frozen).
"""
import argparse, csv, os
from decimal import Decimal, getcontext
import mpmath as mp
import os as _os, sys as _sys
# Single-sourced reference helper: benchmark/_ibeta.py is the only copy.
_sys.path.insert(0, _os.path.dirname(_os.path.dirname(_os.path.abspath(__file__))))
from _ibeta import ibeta, f_cdf
# Single-sourced evaluation primitives shared with the release gate
# (compute_errors.py) so the two analyzers cannot diverge on metric arithmetic,
# parsing, or tail-residual normalisation.
from _contract_eval import (parse_observed, parse_reference, calculate_error,
                            normalize_tail_residual, dispositions)
getcontext().prec = 50
mp.mp.dps = 50

HERE = os.path.dirname(os.path.abspath(__file__))

def load_contracts(p):
    return list(csv.DictReader(open(p)))

def worst_for(measure, metric, rows, fn):
    # `metric` (absolute|relative) chooses the arithmetic; `measure` only selects
    # the tail_probability_residual path. Rows in the F envelope-reject region
    # (expected_error) carry no accuracy claim and are excluded; a blank/ERROR row
    # inside the envelope is unexpected and is reported so main() can BLOCK it
    # rather than silently drop it.
    to_measure, n_expected, n_missing, n_error, n_violation = dispositions(rows)
    w = Decimal(-1); at = ""; n = 0
    for r in to_measure:
        o = parse_observed(r["observed_vba"])
        if o is None:                      # defensive: dispositions already excluded these
            continue
        if measure == "tail_probability_residual":
            target = mp.mpf(r["arg1"]); a2 = mp.mpf(r["arg2"]); a3 = mp.mpf(r["arg3"]); x = mp.mpf(str(o))
            rec = ibeta(x, a2, a3) if fn == "Beta_InverseCumulative" else f_cdf(x, a2, a3)
            e = normalize_tail_residual(rec, target)
        else:
            ref = parse_reference(r["reference"])
            if ref is None:
                continue
            e = calculate_error(o, ref, metric)
        n += 1
        if e > w:
            w = e; at = ", ".join(z for z in (r["arg1"], r["arg2"], r["arg3"]) if z)
    worst = w if n else None
    return worst, at, n, n_missing, n_error, n_violation

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--grid", default=os.path.join(HERE, "holdout_grid.csv"))
    ap.add_argument("--contracts", default=os.path.join(HERE, "..", "accuracy_contracts.csv"))
    a = ap.parse_args()
    rows = list(csv.DictReader(open(a.grid)))
    contracts = load_contracts(a.contracts)

    grid_by = {}
    for r in rows:
        grid_by.setdefault((r["function"], r["regime"]), []).append(r)

    # Select contracts by PRESENCE in the holdout grid (function + regime), not by
    # provenance. The holdout must stay rerunnable after freezing, so both
    # "measured provisional" and "validated and frozen" contracts are re-checked.
    print("Independent holdout - contract thresholds vs fresh data\n")
    header = f"{'Contract':<48}{'metric':<10}{'threshold':>10}{'holdout worst':>15}{'pts':>5}{'margin':>9}  verdict"
    print(header)
    results = []
    all_hold = True; tested = 0; incomplete = 0
    for c in sorted(contracts, key=lambda c: c["contract_id"]):
        matched = grid_by.get((c["function"], c["regime"]), [])
        if not matched:
            continue                       # this contract's regime is not in the holdout
        if not c["threshold"].strip():
            continue                       # no numeric threshold to test (e.g. characterized)
        w, at, n, n_missing, n_error, n_violation = worst_for(c["measure"], c["metric"], matched, c["function"])
        if n_missing or n_error or n_violation:
            # Unexpected blank/ERROR evidence INSIDE the envelope, or an envelope-reject
            # row that failed to return #NUM!, blocks - it is not silently excluded.
            incomplete += 1
            bits = []
            if n_missing:
                bits.append(f"{n_missing} unobserved")
            if n_error:
                bits.append(f"{n_error} unexpected ERROR")
            if n_violation:
                bits.append(f"{n_violation} envelope-reject not #NUM!")
            print(f"{c['contract_id']:<48}{c['metric']:<10}{c['threshold']:>10}"
                  f"{'INCOMPLETE':>15}{n:>5}{'':>9}  {'; '.join(bits)}")
            results.append((c, None, None, n, "INCOMPLETE"))
            continue
        if w is None or w < 0:
            print(f"{c['contract_id']:<48}{c['metric']:<10}{c['threshold']:>10}{'(no obs)':>15}")
            results.append((c, None, None, 0, "NO OBS"))
            continue
        thr = Decimal(c["threshold"]); ok = w <= thr; tested += 1
        margin = float(thr / w) if w > 0 else float("inf")
        all_hold = all_hold and ok
        verdict = "PASS" if ok else "FAIL"
        print(f"{c['contract_id']:<48}{c['metric']:<10}{c['threshold']:>10}{float(w):>15.2e}{n:>5}{margin:>8.1f}x  {verdict}")
        results.append((c, w, margin, n, verdict))
    print()
    if incomplete:
        print(f"{incomplete} contract(s) have unexpected missing/ERROR evidence in-envelope - "
              "these BLOCK rather than being silently excluded.")
    if tested and all_hold and not incomplete:
        print(f"ALL {tested} contract(s) present in the holdout hold on fresh data.")
    elif tested and not all_hold:
        print("At least one contract exceeded its threshold on the holdout - investigate before (re)freezing.")
    elif not tested and not incomplete:
        print("No contracts matched the holdout grid - check that the grid's function/regime "
              "tags line up with the contract file.")

    # Write a committed summary artifact so the freezing evidence is reproducible.
    out = os.path.join(HERE, "holdout_summary.md")
    md = ["# Holdout summary", "",
          "Regenerated by `analyze_holdout.py`. Every contract whose (function, regime) "
          "appears in `holdout_grid.csv` is re-checked against fresh, unseen data - the "
          "evidence that justified freezing, reproducible on demand.", "",
          "| Contract | Metric | Threshold | Holdout worst | Points | Margin | Provenance | Verdict |",
          "|---|---|---|---:|---:|---:|---|---|"]
    for c, w, margin, n, verdict in results:
        wtxt = f"{float(w):.2e}" if w is not None else "-"
        mtxt = f"{margin:.1f}x" if margin is not None else "-"
        md.append(f"| {c['contract_id']} | {c['metric']} | {c['threshold']} | {wtxt} | "
                  f"{n} | {mtxt} | {c['provenance']} | {verdict} |")
    npass = sum(1 for _, _, _, _, v in results if v == "PASS")
    nfail = sum(1 for _, _, _, _, v in results if v == "FAIL")
    ninc = sum(1 for _, _, _, _, v in results if v == "INCOMPLETE")
    tail = f", {ninc} incomplete" if ninc else ""
    md += ["", f"> {npass} pass, {nfail} fail{tail} across {len(results)} contract(s) present "
           "in the holdout. Margin = threshold / holdout worst; higher is more headroom."]
    with open(out, "w", encoding="utf-8") as f:
        f.write("\n".join(md) + "\n")
    print(f"\nwrote {out}")
    if (tested and not all_hold) or incomplete:
        raise SystemExit(1)

if __name__ == "__main__":
    main()
