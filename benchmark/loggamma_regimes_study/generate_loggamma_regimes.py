"""
Regime study for PROB_LogGamma.

The module has advertised a single global claim, "relative error below 6.1E-14
across Z in [1E-8, 1E+50]". That claim is wrong in two independent ways, and
this grid is the evidence for replacing it with regime-aware contracts:

  A. The reflection path forms Sin(PROB_PI * Z) after PROB_PI * Z has entered
     the subnormal range and lost significand bits. Log(Sin(...)) is then off by
     that relative error ABSOLUTELY, on a result of magnitude ~700. A genuine
     defect, corrected by the LogGamma1p(Z) - Log(Z) branch.

  B. Log(Gamma(Z)) is zero at Z = 1 and Z = 2, so a global RELATIVE contract is
     ill-conditioned by construction. Near those zeros a few E-15 of absolute
     error reads as a large relative error and says nothing about the kernel.

Metric. Absolute error in the logarithm is the primary metric everywhere except
the general regime, because that is the quantity downstream callers propagate:
the relative error of Exp(v) is approximately the absolute error of v. The two
exact zeros are evaluated for absolute error ONLY; they cannot participate in a
relative contract at all.

Baseline. Export this grid ONCE with the pre-Phase-1 module and keep it as
loggamma_regimes_baseline.csv, then again after the edit. Without a baseline the
subnormal improvement has nothing to be measured against: no committed
observation covers Z below 1E-8.
"""
import argparse
import csv
import math
import struct

import mpmath as mp

SERIES_MAX = 0.25


def next_up(x):
    return math.nextafter(x, math.inf)


def next_down(x):
    return math.nextafter(x, -math.inf)


def reference(z):
    """Log(Gamma(Z)) at precision sufficient for arbitrarily small positive Z."""
    dps = 60 + max(0, int(-math.log10(z)) + 10) if z < 1 else 60
    with mp.workdps(dps):
        return +mp.loggamma(mp.mpf(z))


def regime(z):
    """Which contract bucket the point belongs to."""
    if z in (1.0, 2.0):
        return "exact_zero"                      # absolute error only, never relative
    if abs(z - 1.0) < 1e-9 or abs(z - 2.0) < 1e-9:
        return "near_zero"
    if 0.9 <= z <= 1.1 or 1.9 <= z <= 2.1:
        return "near_zero"
    if z <= SERIES_MAX:
        return "small_positive"                  # the branch Phase 1 replaces
    if z < 0.5:
        return "reflection"                      # unchanged branch; must not move
    return "general"


def build_points():
    pts = []
    # subnormal band where the reflection route actually degrades
    pts += [5e-324, 1e-322, 1e-320, 1e-318, 1e-315, 1e-312, 1e-310, 1e-308]
    # binary landmarks: retained significand bits are legible directly
    pts += [2.0 ** -1022, 2.0 ** -1030, 2.0 ** -1040, 2.0 ** -1050,
            2.0 ** -1060, 2.0 ** -1070, 2.0 ** -1074]
    # decades through the small-positive branch
    pts += [1e-300, 1e-250, 1e-200, 1e-100, 1e-50, 1e-20, 1e-13, 1e-10,
            1e-8, 1e-6, 1e-4, 1e-3, 1e-2]
    # arguments the committed Beta/F rows actually reach
    pts += [9.9e-14, 2.7e-12, 3e-11, 0.02, 0.04, 0.05, 0.1, 0.125, 0.2, 0.24]
    # series seam
    pts += [next_down(SERIES_MAX), SERIES_MAX, next_up(SERIES_MAX)]
    # reflection interval and its upper boundary
    pts += [0.3, 0.4, 0.45, next_down(0.5), 0.5, next_up(0.5)]
    # approach to the first zero
    pts += [0.6, 0.75, 0.9, 0.99, next_down(1.0)]
    # the zeros themselves and their neighbourhoods
    pts += [1.0, next_up(1.0), 1.01, 1.25, 1.5, 1.75, 1.99,
            next_down(2.0), 2.0, next_up(2.0), 2.01]
    # general regime, where relative error remains meaningful
    pts += [2.5, 3.0, 5.0, 10.0, 100.0, 1e6, 1e15, 1e30, 1e50]
    out = []
    for p in pts:
        if p not in out:
            out.append(p)
    return out


QUANTITIES = ("EchoZ", "LogGamma")


def build():
    rows = []
    for z in build_points():
        ref = reference(z)
        reg = regime(z)
        for q in QUANTITIES:
            r = mp.mpf(z) if q == "EchoZ" else ref
            rows.append({
                "quantity": q,
                "regime": reg,
                "metric": "absolute" if (q == "EchoZ" or reg != "general") else "relative",
                "arg1": repr(z),
                "bits": str(1074 + math.floor(math.log2(z)) + 1) if z < 2.2250738585072014e-308 else "53",
                "reference": mp.nstr(r, 34),
                "observed_vba": "",
            })
    return rows


FIELDS = ["quantity", "regime", "metric", "arg1", "bits", "reference", "observed_vba"]


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default="loggamma_regimes_grid.csv")
    a = ap.parse_args()
    rows = build()
    with open(a.out, "w", newline="") as f:
        w = csv.DictWriter(f, fieldnames=FIELDS, lineterminator="\n")
        w.writeheader()
        w.writerows(rows)
    n = len(rows) // len(QUANTITIES)
    print(f"wrote {a.out}: {len(rows)} rows ({n} points x {len(QUANTITIES)} quantities)")


if __name__ == "__main__":
    main()
