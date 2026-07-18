"""
Analyze the LogBeta switch study: relative error of PROB_LogBeta vs ratio.

Reads logbeta_switch_grid.csv after the VBA export macro has filled observed_vba
(as hi;lo two-part sums). Prints, per Small value, the relative error at each
Small/Large ratio, flags where it crosses the 5E-15 Beta claim, and marks where
the asymptotic branch currently fires (ratio <= 1E-15).
"""
import argparse, csv
from decimal import Decimal, getcontext
getcontext().prec = 60

SWITCH = Decimal("1E-15")     # current PROB_EPS branch threshold
CLAIM  = Decimal("5E-15")

def parse_observed(s):
    s = s.strip()
    if not s or s.upper() == "ERROR":
        return None
    return sum(Decimal(p) for p in s.split(";"))

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--grid", default="logbeta_switch_grid.csv")
    args = ap.parse_args()
    rows = [r for r in csv.DictReader(open(args.grid)) if r["function"] == "LogBeta"]

    by_small = {}
    for r in rows:
        by_small.setdefault(r["arg2"], []).append(r)

    print("PROB_LogBeta relative error vs Small/Large ratio")
    print("(branch fires at ratio <= 1E-15; general identity used above it)\n")
    for small in sorted(by_small, key=lambda x: Decimal(x)):
        print(f"Small = {small}")
        print(f"  {'ratio':>10} {'rel_err':>12} {'>claim?':>8} {'path':>9}")
        grp = sorted(by_small[small], key=lambda r: Decimal(r["arg1"]))
        for r in grp:
            large = Decimal(r["arg1"]); sm = Decimal(r["arg2"])
            ratio = sm / large
            ref = Decimal(r["reference"])
            obs = parse_observed(r["observed_vba"])
            if obs is None:
                print(f"  {float(ratio):>10.0e} {'(no obs)':>12}")
                continue
            rel = abs(obs - ref) / abs(ref) if ref != 0 else Decimal(0)
            over = "FAIL" if rel > CLAIM else "ok"
            path = "branch" if ratio <= SWITCH else "general"
            print(f"  {float(ratio):>10.0e} {float(rel):>12.2e} {over:>8} {path:>9}")
        print()

if __name__ == "__main__":
    main()
