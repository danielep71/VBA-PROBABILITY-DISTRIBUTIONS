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
from _ibeta import ibeta, f_cdf
getcontext().prec = 50
mp.mp.dps = 50

HERE = os.path.dirname(os.path.abspath(__file__))

def parse(s):
    s = s.strip()
    return None if (not s or s.upper() == "ERROR") else sum(Decimal(p) for p in s.split(";"))

def load_contracts(p):
    return list(csv.DictReader(open(p)))

def worst_for(measure, rows, fn):
    w = Decimal(-1); at = ""; n = 0
    for r in rows:
        o = parse(r["observed_vba"])
        if o is None:
            continue
        if measure == "tail_probability_residual":
            p = mp.mpf(r["arg1"]); a2 = mp.mpf(r["arg2"]); a3 = mp.mpf(r["arg3"]); x = mp.mpf(str(o))
            rec = ibeta(x, a2, a3) if fn == "Beta_InverseCumulative" else f_cdf(x, a2, a3)
            e = Decimal(str(abs(rec - p) / min(p, 1 - p)))
        else:
            ref = Decimal(r["reference"])
            ae = abs(o - ref)
            e = ae if measure == "log_absolute_error" else (ae / abs(ref) if ref != 0 else Decimal(0))
        n += 1
        if e > w:
            w = e; at = ", ".join(z for z in (r["arg1"], r["arg2"], r["arg3"]) if z)
    return (w, at, n) if n else (None, "", 0)

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

    print("Independent holdout — provisional thresholds vs fresh data\n")
    print(f"{'Contract':<48}{'threshold':>10}{'holdout worst':>15}{'margin':>9}  verdict")
    all_hold = True; tested = 0
    for c in sorted(contracts, key=lambda c: c["contract_id"]):
        if c["provenance"] != "measured provisional":
            continue
        matched = grid_by.get((c["function"], c["regime"]), [])
        if not matched:
            print(f"{c['contract_id']:<48}{c['threshold']:>10}{'(no holdout pts)':>15}")
            continue
        w, at, n = worst_for(c["measure"], matched, c["function"])
        if w is None:
            print(f"{c['contract_id']:<48}{c['threshold']:>10}{'(no obs)':>15}")
            continue
        thr = Decimal(c["threshold"]); ok = w <= thr; tested += 1
        margin = float(thr / w) if w > 0 else float("inf")
        all_hold = all_hold and ok
        print(f"{c['contract_id']:<48}{c['threshold']:>10}{float(w):>15.2e}{margin:>8.1f}x  "
              f"{'PASS' if ok else 'FAIL'}")
    print()
    if tested and all_hold:
        print(f"ALL {tested} provisional contracts hold on the holdout -> ready to freeze")
        print("(flip provenance: measured provisional -> validated and frozen)")
    else:
        print("Some contract exceeded its threshold on the holdout -> do NOT freeze; adjust that threshold")

if __name__ == "__main__":
    main()
