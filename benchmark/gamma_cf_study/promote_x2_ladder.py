"""Reference-only #26 promotion: the preregistered X = 2 survival ladder.

The ladder is the twelve x2_* rows of
benchmark/positive_ratio_subnormal/gamma_q_complement.csv, selected by their
frozen role (X = 2), never by a measured error. Each shape A gives two main-grid
rows with the same reference Q(A, 2):

    Gamma_Survival(2, A, 1)
    ChiSquare_Survival(4, 2A)      (df = 2A and x = 4 are exact doublings)

Use with benchmark/promote_grid_rows.py; never copy study observations.

Both rows claim rel<=5E-15, the Gamma_Survival.all.output contract. The
registry's ChiSquare_Survival.all.output contract is broader (2.6E-10), so
meets_frozen_bound() holds the Chi-square rows to 5E-15 independently, as
#23 does for its promoted Chi-square row.

Run: python3 promote_x2_ladder.py   (prints the rows and the reference check)
"""
from decimal import Decimal, DecimalException, localcontext
from pathlib import Path
import csv
import sys

import mpmath as mp

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE.parent))
from _contract_eval import calculate_error, parse_observed, parse_reference  # noqa: E402

LADDER = HERE.parent / "positive_ratio_subnormal" / "gamma_q_complement.csv"
X = 2.0
CLAIM = "rel<=5E-15"
BOUND = Decimal("5E-15")


def ladder_shapes():
    """The binary64 shapes of the frozen x2_* rows, in file order."""
    with open(LADDER, newline="", encoding="utf-8") as f:
        rows = [r for r in csv.DictReader(f) if r["point_id"].startswith("x2_")]
    shapes = []
    for r in rows:
        hi, lo = r["shape"].split(";")
        x_hi, x_lo = r["x"].split(";")
        if float(lo) != 0.0 or float(x_lo) != 0.0 or float(x_hi) != X:
            raise ValueError(f"{r['point_id']}: not an exact X = 2 ladder row")
        shapes.append(float(hi))
    return shapes


def upper_q(shape, dps):
    with mp.workdps(dps):
        return mp.gammainc(mp.mpf(shape), mp.mpf(X), mp.inf, regularized=True)


def reference(shape):
    """Q(A, 2) at the binary64 shape, required stable between 80 and 120 digits."""
    low, high = upper_q(shape, 80), upper_q(shape, 120)
    with mp.workdps(120):
        if abs(low - high) > mp.mpf("1E-60") * abs(high):
            raise ValueError(f"reference for shape {shape!r} unstable at 80/120 digits")
    with mp.workdps(120):
        return mp.nstr(high, 60)


def build_rows():
    rows = []
    for shape in ladder_shapes():
        ref = reference(shape)
        for function, a1, a2, a3 in (("Gamma_Survival", X, shape, 1.0),
                                     ("ChiSquare_Survival", 2.0 * X, 2.0 * shape, None)):
            rows.append({"function": function, "vba_kernel": "K_STATS_" + function,
                         "claim": CLAIM, "metric": "rel", "arg1": repr(a1), "arg2": repr(a2),
                         "arg3": "" if a3 is None else repr(a3), "arg4": "",
                         "reference": ref, "observed_vba": "", "regime": "all",
                         "evidence_set": "main grid", "expected_error": ""})
    return rows


def meets_frozen_bound(row):
    """Hold a promoted observation to 5E-15 even where the registry is broader."""
    try:
        with localcontext() as context:
            context.prec = 80
            observed = parse_observed(row.get("observed_vba", ""))
            ref = parse_reference(row.get("reference", ""))
            if (observed is None or ref is None or not observed.is_finite()
                    or not ref.is_finite() or ref <= 0):
                return False
            return calculate_error(observed, ref, "relative") <= BOUND
    except (DecimalException, ValueError):
        return False


if __name__ == "__main__":
    for r in build_rows():
        print(r["function"], r["arg1"], r["arg2"], r["arg3"], r["reference"][:24])
