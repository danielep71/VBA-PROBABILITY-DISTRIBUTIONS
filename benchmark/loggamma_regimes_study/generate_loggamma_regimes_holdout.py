"""
Independent holdout for the PROB_LogGamma regime contracts.

Validates LogGamma.small_positive.log_abs, LogGamma.near_zero.log_abs and
LogGamma.general.output_rel on data that was NOT used to set any threshold. If
the provisional thresholds hold here they generalise and can be frozen
(`measured provisional` -> `validated and frozen`).

DISJOINTNESS

Every point here is checked against loggamma_regimes_grid.csv at generation
time; the script refuses to write a grid that shares a single Double with the
fitting set. Disjointness is not eyeballed, it is asserted.

The points are also chosen to be structurally unlike the fitting set, not merely
different from it. The main grid uses decade landmarks (1E-320, 1E-315, 1E-8),
exact binary landmarks (2^-1030, 2^-1074) and the arguments the committed Beta/F
rows happen to reach. A holdout built from neighbouring decades would test the
same shape of number. So this grid uses:

  - irrational-mantissa points, so no argument is a round decimal or a power of
    two, and the subnormal quantisation lands differently;
  - odd binary exponents between the fitting set's even landmarks, giving
    retained-bit counts the fitting set never measured;
  - asymmetric offsets around Z = 1 and Z = 2, since the fitting set approaches
    both zeros symmetrically;
  - points inside the reflection interval, which the fitting set samples at only
    five places.

The reference metric per regime matches the contract: absolute error in the
logarithm everywhere except `general`.
"""
import argparse
import csv
import math
import os
import sys

import mpmath as mp

SERIES_MAX = 0.25
FITTING_SET = "loggamma_regimes_grid.csv"

# irrational-ish multipliers: nothing here is a round decimal or a power of two
PHI = 1.6180339887498949
SQ2 = 1.4142135623730951
SQ3 = 1.7320508075688772
E_1 = 2.718281828459045


def reference(z):
    dps = 60 + max(0, int(-math.log10(z)) + 10) if z < 1 else 60
    with mp.workdps(dps):
        return +mp.loggamma(mp.mpf(z))


def regime(z):
    if z in (1.0, 2.0):
        return "exact_zero"
    if 0.9 <= z <= 1.1 or 1.9 <= z <= 2.1:
        return "near_zero"
    if z <= SERIES_MAX:
        return "small_positive"
    if z < 0.5:
        return "reflection"
    return "general"


def build_points():
    pts = []
    # subnormal band, odd binary exponents the fitting set never lands on,
    # multiplied off any power of two
    for e in (1073, 1067, 1051, 1039, 1027, 1023):
        pts.append(math.ldexp(PHI, -e))
    for e in (1061, 1045, 1033):
        pts.append(math.ldexp(SQ3, -e))
    # normal-result small positive, irrational mantissas between the decades
    pts += [SQ2 * 1e-307, PHI * 1e-290, SQ3 * 1e-250, E_1 * 1e-170,
            PHI * 1e-77, SQ2 * 1e-33, SQ3 * 1e-12, E_1 * 1e-7,
            PHI * 1e-5, SQ2 * 1e-3, SQ3 * 1e-2]
    # series interior, off the fitting set's round steps
    pts += [0.0173, 0.0619, 0.1037, 0.1414, 0.1732, 0.2071, 0.2361, 0.2489]
    # reflection interval, which the fitting set samples only five times
    pts += [0.2618, 0.2887, 0.3183, 0.3536, 0.3820, 0.4142, 0.4472, 0.4796]
    # asymmetric approach to the zeros
    pts += [0.9128, 0.9511, 0.9819, 0.9973, 1.0027, 1.0181, 1.0489, 1.0872]
    pts += [1.9128, 1.9511, 1.9819, 1.9973, 2.0027, 2.0181, 2.0489, 2.0872]
    # general regime, non-integer and spread over the validated range
    pts += [2.7183, 3.1416, 6.2832, 17.32, 141.42, 2.7183e4, 6.1803e9,
            1.4142e18, 3.1416e33, 8.6603e47]
    out = []
    for p in pts:
        if p > 0 and p not in out:
            out.append(p)
    return out


FIELDS = ["quantity", "regime", "metric", "arg1", "bits", "reference", "observed_vba"]
QUANTITIES = ("EchoZ", "LogGamma")


def build(fitting):
    rows = []
    for z in build_points():
        if z in fitting:
            raise SystemExit(f"holdout point {z!r} is in the fitting set; refusing to write")
        ref = reference(z)
        reg = regime(z)
        for q in QUANTITIES:
            rows.append({
                "quantity": q,
                "regime": reg,
                "metric": "absolute" if (q == "EchoZ" or reg != "general") else "relative",
                "arg1": repr(z),
                "bits": str(1074 + math.floor(math.log2(z)) + 1)
                        if z < 2.2250738585072014e-308 else "53",
                "reference": mp.nstr(mp.mpf(z) if q == "EchoZ" else ref, 34),
                "observed_vba": "",
            })
    return rows


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default="loggamma_regimes_holdout.csv")
    ap.add_argument("--fitting", default=FITTING_SET)
    a = ap.parse_args()

    here = os.path.dirname(os.path.abspath(__file__))
    fpath = a.fitting if os.path.isabs(a.fitting) else os.path.join(here, a.fitting)
    if not os.path.exists(fpath):
        sys.exit(f"fitting set {fpath} not found; disjointness cannot be verified")
    fitting = {float(r["arg1"]) for r in csv.DictReader(open(fpath))}

    rows = build(fitting)
    with open(a.out, "w", newline="") as f:
        w = csv.DictWriter(f, fieldnames=FIELDS, lineterminator="\n")
        w.writeheader()
        w.writerows(rows)
    n = len(rows) // len(QUANTITIES)
    print(f"wrote {a.out}: {len(rows)} rows ({n} points x {len(QUANTITIES)} quantities)")
    print(f"disjoint from {a.fitting} ({len(fitting)} fitting points): verified")


if __name__ == "__main__":
    main()
