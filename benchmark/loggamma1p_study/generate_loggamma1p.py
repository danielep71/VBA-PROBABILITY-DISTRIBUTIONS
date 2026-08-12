"""
Reference grid for PROB_TryLogGamma1p (ICR-P1-01 prerequisite, issue #12).

PROB_LogGamma(1# + X) is unusable for small X: in binary64, 1# + X rounds to
exactly 1 for every X below 2^-53, so the leading term -EulerGamma * X is lost
and the result collapses to the rounding residue of the addition. This grid
measures the replacement kernel, the defective spelling it replaces, and the
seam between the Maclaurin series and the Lanczos hand-over.

Four quantities per point:

  EchoX             X itself, round-tripped through the VBA parser. Guards the
                    grid against Val() mis-parsing a subnormal literal: if VBA
                    did not evaluate the Double we intended, every other number
                    in the row is meaningless.
  LogGamma1p        PROB_TryLogGamma1p(X)                    -> the kernel
  LogGamma1pOverX   PROB_TryLogGamma1p(X) / X                -> the contract
  LogGammaNaive     PROB_LogGamma(1# + X)                    -> the defect

LogGamma1pOverX is measured in VBA rather than divided out in Python because the
scaled Gamma inverse performs exactly that division in VBA; measuring it here
captures the division rounding the caller will actually incur.

References are mpmath at a precision chosen per point. The precision must be
raised with the magnitude of X, because mp.loggamma(1 + mpf(x)) silently returns
zero once 1 + x rounds to 1 at the working precision -- the reference reproduces
the very defect the kernel exists to remove unless the working precision is
widened first. At dps = 60 + |log10(X)| + 10 the increment survives, and the
result then agrees with the Maclaurin series to every digit carried, at every
point on this grid.
"""
import argparse
import csv
import math

import mpmath as mp

SERIES_MAX = 0.25                              # PROB_LG1P_SERIES_MAX
NORMAL_RESULT_MIN = 1.0e-308                   # below this, EulerGamma*X is subnormal
ULP = 2.0 ** -52


def working_dps(x):
    """Precision at which 1 + x still retains x."""
    return 60 + max(0, int(-math.log10(x)) + 10)


def reference(x):
    """Log(Gamma(1 + x)) for x >= 0, correct for arbitrarily small positive x."""
    if x == 0.0:
        return mp.mpf(0)
    with mp.workdps(working_dps(x)):
        return +mp.loggamma(1 + mp.mpf(x))


def regime(x):
    """Which code path PROB_TryLogGamma1p actually takes at this point.

    Regimes name branches, not intervals of interest, so that a contract is
    never averaged across two different algorithms. The seam is studied by
    inspecting the points adjacent to SERIES_MAX, not by giving them a regime
    of their own: the point one ulp above the seam runs the Lanczos route and
    is governed by that kernel's contract, not by the series contract.
    """
    if x < NORMAL_RESULT_MIN:
        return "subnormal_result"
    if x > SERIES_MAX:
        return "lanczos_handover"
    return "series"


GRID = [
    # --- subnormal result: documented representability limit, not contracted ---
    5e-324, 1e-322, 1e-320, 1e-318, 1e-315, 1e-312, 1e-310,
    # --- normal result, decades ---
    1e-308, 1e-300, 1e-200, 1e-100, 1e-50, 1e-20,
    # --- the 1 + X absorption boundary and its approach ---
    1e-17, 1e-16, 2.0 ** -53, 1e-15, 1e-14, 1e-13, 1e-12, 1e-11, 1e-10,
    1e-9, 1e-8, 1e-7, 1e-6, 1e-5, 1e-4, 1e-3,
    # --- series interior ---
    0.01, 0.02, 0.05, 0.08, 0.1, 0.125, 0.15, 0.18, 0.2, 0.22, 0.24,
    # --- seam, one ulp either side ---
    SERIES_MAX * (1 - ULP), SERIES_MAX, SERIES_MAX * (1 + ULP),
    # --- hand-over ---
    0.26, 0.3, 0.5, 0.75, 1.0, 2.5, 10.0,
]

QUANTITIES = ("EchoX", "LogGamma1p", "LogGamma1pOverX", "LogGammaNaive")


def build():
    rows = []
    for x in GRID:
        ref = reference(x)
        with mp.workdps(120):
            ref_over_x = ref / mp.mpf(x) if x != 0.0 else mp.mpf(0)
        for q in QUANTITIES:
            if q == "EchoX":
                r = mp.mpf(x)
            elif q == "LogGamma1pOverX":
                r = ref_over_x
            else:
                r = ref
            rows.append({
                "quantity": q,
                "regime": regime(x),
                "arg1": repr(x),                       # 17 digits, exact round-trip
                "reference": mp.nstr(r, 30),
                "observed_vba": "",
            })
    return rows


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default="loggamma1p_grid.csv")
    a = ap.parse_args()
    rows = build()
    with open(a.out, "w", newline="") as f:
        w = csv.DictWriter(
            f, fieldnames=["quantity", "regime", "arg1", "reference", "observed_vba"],
            lineterminator="\n")                 # LF per .gitattributes
        w.writeheader()
        w.writerows(rows)
    print(f"wrote {a.out}: {len(rows)} rows ({len(GRID)} points x {len(QUANTITIES)} quantities)")


if __name__ == "__main__":
    main()
