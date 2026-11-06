"""
Binary64 model of the production Student-t large-df tail route (#34), on the
repaired normal kernel of #72.

PRODUCTION_PREREGISTRATION.md freezes the route this module models. It is the
order-5 expansion the study adopted (PREREGISTRATION.md, results.json),

    S(t; df) = Q(t) + phi(t) * sum_{k=1..5} g_k(t) / df^k,   df >= 1E5, |t| <= 8,

written the way PROB_TryStudentTTail will evaluate it in VBA:

  * every g_k is odd, g_k(t) = t * P_k(t^2), so U = t * t and each P_k is
    evaluated by Horner in U;
  * the sum over k is Horner in W = 1 / df;
  * Q(t) is PROB_NormalUpperTail, phi(t) is the module's 1/sqrt(2 pi) literal
    pair times PROB_ExpNegHalfSquare, both modelled by
    normal_tail_study/kernel_model.py, which reproduces the committed Excel
    observations of the repaired kernel within 4 ulps;
  * every coefficient is the Double nearest the exact rational of
    coefficients.json, written as a hi + lo pair of 15-significant-digit
    literals whose Double sum is exactly that Double (the VBE keeps 15 digits).

Each statement rounds to Double, as a VBA assignment to a Double variable does.

    python production_route_model.py      print the coefficient literals
"""
import json
import os
import sys
from fractions import Fraction

HERE = os.path.dirname(os.path.abspath(__file__))
BENCH = os.path.dirname(HERE)
sys.path.insert(0, os.path.join(BENCH, "normal_tail_study"))
import kernel_model as km  # noqa: E402

ORDER = 5                   # the adopted K*
DF_MIN = 1e5                # the adopted DF_MIN
T_MAX = 8.0                 # the adopted T_MAX
BOUND = 2.0 ** -50          # the study's B


def vba_literal_pair(x):
    """(hi, lo) 15-significant-digit literals whose Double sum is exactly x."""
    hi = format(x, ".15g")
    remainder = Fraction(x) - Fraction(float(hi))
    lo = "0" if remainder == 0 else format(float(remainder), ".15g")
    return hi, lo


def _coefficients():
    """P_k coefficients by ascending power of U, as (Fraction, Double, hi, lo)."""
    with open(os.path.join(HERE, "coefficients.json"), encoding="utf-8") as f:
        g = json.load(f)["g"]
    table = []
    for k in range(1, ORDER + 1):
        terms = {power: Fraction(num, den) for power, num, den in g[str(k)]["terms"]}
        if any(power % 2 == 0 for power in terms):
            raise ValueError(f"g_{k} is not odd")
        degree = (max(terms) - 1) // 2
        row = []
        for j in range(degree + 1):
            exact = terms.get(2 * j + 1, Fraction(0))
            double = float(exact)
            hi, lo = vba_literal_pair(double)
            if float(hi) + float(lo) != double:
                raise ValueError(f"g_{k} U^{j}: the literal pair does not rebuild its Double")
            row.append((exact, double, hi, lo))
        table.append(row)
    return table


COEFFICIENTS = _coefficients()
INV_SQRT_TWO_PI = km.COEFFICIENTS["inv_sqrt_two_pi"]


def series(t, df):
    """sum_k P_k(U) W^k, statement by statement."""
    u = t * t
    w = 1.0 / df
    total = 0.0
    for row in reversed(COEFFICIENTS):          # k = 5 down to 1
        p = row[-1][1]
        for _, c, _, _ in reversed(row[:-1]):
            p = p * u + c
        total = total * w + p
    return total * w


def in_region(abs_t, df):
    return df >= DF_MIN and abs_t <= T_MAX


def route_tail(abs_t, df):
    """The production upper tail S(|t|; df) inside the region."""
    q = km.candidate_upper_tail(abs_t)
    phi = INV_SQRT_TWO_PI * km.exp_neg_half_square(abs_t)
    correction = abs_t * series(abs_t, df)
    correction = phi * correction
    return q + correction


def survival(t, df):
    """K_STATS_StudentT_Survival inside the region."""
    tail = route_tail(abs(t), df)
    return tail if t >= 0.0 else 1.0 - tail


def cumulative(t, df):
    """K_STATS_StudentT_Cumulative inside the region."""
    tail = route_tail(abs(t), df)
    return tail if t < 0.0 else 1.0 - tail


def main():
    for k, row in enumerate(COEFFICIENTS, start=1):
        for j, (exact, double, hi, lo) in enumerate(row):
            print(f"g_{k}  U^{j:<2} {str(exact):>18}  {hi:>24} + {lo}")


if __name__ == "__main__":
    main()
