"""
Step-6 analysis: function-level relative error of the public Beta/F functions at
strongly disparate shapes / degrees of freedom.

Per the metric correction, each PUBLIC function is judged by its OWN relative
error (not the PROB_LogBeta proxy), because the incomplete-beta continued
fraction, tail selection and inverse solver damp or amplify the LogBeta
normalization error differently. This produces the measured per-function
worst-case that freezes the unbalanced contract threshold.
"""
import argparse, csv
from collections import defaultdict
from decimal import Decimal, getcontext
getcontext().prec = 50

def parse(s):
    s = s.strip()
    if not s or s.upper() == "ERROR":
        return None
    return sum(Decimal(p) for p in s.split(";"))

def suggest_threshold(worst):
    """Provisional frozen threshold = worst measured, rounded up with headroom."""
    if worst is None or worst == 0:
        return None
    # round up to 1 significant figure, then double for headroom
    import math
    exp = math.floor(math.log10(float(worst)))
    lead = float(worst) / (10 ** exp)
    rounded = math.ceil(lead) * (10 ** exp)
    return 2 * rounded

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--grid", default="beta_f_unbalanced_grid.csv")
    a = ap.parse_args()
    rows = list(csv.DictReader(open(a.grid)))

    per_fn = defaultdict(list)     # fn -> list of (rel_err, args, |value|)
    errors = defaultdict(int)
    for r in rows:
        fn = r["function"]
        obs = parse(r["observed_vba"])
        ref = Decimal(r["reference"]) if r["reference"].strip() not in ("", "0") else Decimal(0)
        if obs is None:
            errors[fn] += 1
            continue
        if ref == 0:
            rel = Decimal(0) if obs == 0 else Decimal("Infinity")
        else:
            rel = abs(obs - ref) / abs(ref)
        args = ", ".join(x for x in (r["arg1"], r["arg2"], r["arg3"]) if x)
        per_fn[fn].append((rel, args, abs(ref)))

    print("Public Beta/F function-level relative error at unbalanced arguments")
    print("(each function judged by its OWN relative error)\n")
    print(f"{'Function':<18} {'points':>7} {'ERROR':>6} {'worst rel err':>14} {'at (X, a, b / df1, df2)':<40}")
    frozen = {}
    for fn in sorted(per_fn):
        pts = per_fn[fn]
        worst, at, _ = max(pts, key=lambda t: t[0])
        frozen[fn] = suggest_threshold(worst)
        n_err = errors.get(fn, 0)
        print(f"{fn:<18} {len(pts):>7} {n_err:>6} {float(worst):>14.2e}  {at:<40}")

    print("\nSuggested frozen unbalanced contract thresholds (worst measured x headroom):")
    for fn in sorted(frozen):
        t = frozen[fn]
        print(f"  {fn:<18} relative <= {t:.0e}" if t else f"  {fn:<18} (no data)")
    print("\nFreeze these ONLY after the final VBA implementation; measure each function")
    print("separately (do not infer one common threshold). Retain a known_limitation")
    print("only where a measured function still exceeds its revised contract.")

if __name__ == "__main__":
    main()
