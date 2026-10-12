"""
Binary64 model of the candidate production route for #34 (production
feasibility, finding F1).

The study in this folder validated the large-df tail expansion

    S(t; df) = Q(t) + phi(t) * sum_{k=1..K} g_k(t) / df^k

with every term evaluated at high precision. Production VBA would evaluate it
in binary64 and would obtain Q(t) from a kernel. This module models that
production arithmetic so its accuracy can be measured before any VBA is
written:

  * vba_upper_tail(z) is a line-for-line binary64 port of
    PROB_NormalUpperTailPositive in src/M_STATS_PROBDIST_NORMALFAMILY.bas
    (Hart rational approximation below 7.07106781186547, Laplace continued
    fraction above), with the module's constants built the way the VBE builds
    them;
  * route(t, df, q) evaluates the order-5 expansion in binary64, the
    coefficients rounded to Double, each g_k by Horner in t and the sum by
    Horner in 1/df, with the normal tail supplied by q.

Fidelity: test_binary64_route_model.py checks the port against every committed
Excel observation of NormalStandard_Survival and NormalStandard_Cumulative.
The remaining differences are within the last-bit difference between VBA Exp
and the C library's exp.

    python binary64_route_model.py     print the feasibility table
"""
import json
import math
import os

import mpmath as mp

HERE = os.path.dirname(os.path.abspath(__file__))
ORDER = 5                       # adopted K* (results.json "adopted")
DF_MIN = 1e5                    # adopted DF_MIN
T_MAX = 8.0                     # adopted T_MAX
BOUND = 2.0 ** -50              # preregistered accuracy bound B

# Constants exactly as M_STATS_PROBDIST_NORMALFAMILY.bas declares them; the
# split expression is evaluated in Double, as the VBE does.
INV_SQRT_TWO_PI = 0.398942280401432 + 7.21644966006352E-16
CDF_SPLIT = 7.07106781186547
TAIL_UNDERFLOW_Z = 38.5
NORMAL_TAIL_CF_TERMS = 16


def vba_upper_tail(z):
    """Binary64 port of PROB_NormalUpperTailPositive (Q(z) for z >= 0)."""
    if z >= TAIL_UNDERFLOW_Z:
        return 0.0
    exponential = math.exp(-0.5 * z * z)
    if z < CDF_SPLIT:
        a = 3.52624965998911E-02 * z + 0.700383064443688
        a = a * z + 6.37396220353165
        a = a * z + 33.912866078383
        a = a * z + 112.079291497871
        a = a * z + 221.213596169931
        a = a * z + 220.206867912376
        b = 8.83883476483184E-02 * z + 1.75566716318264
        b = b * z + 16.064177579207
        b = b * z + 86.7807322029461
        b = b * z + 296.564248779674
        b = b * z + 637.333633378831
        b = b * z + 793.826512519948
        b = b * z + 440.413735824752
        return exponential * a / b
    denominator = z
    for k in range(NORMAL_TAIL_CF_TERMS, 0, -1):
        denominator = z + float(k) / denominator
    return exponential * INV_SQRT_TWO_PI / denominator


def vba_survival(z):
    """PROB_NormalSurvival: the positive tail directly, else its complement."""
    return vba_upper_tail(z) if z >= 0.0 else 1.0 - vba_upper_tail(-z)


def vba_cdf(z):
    """PROB_NormalCDF: the negative tail directly, else its complement."""
    return vba_upper_tail(-z) if z < 0.0 else 1.0 - vba_upper_tail(z)


def exact_upper_tail(z):
    """Q(z) correctly rounded to Double: the best any normal kernel can give."""
    with mp.workdps(60):
        return float(mp.erfc(mp.mpf(z) / mp.sqrt(2)) / 2)


def _coefficients():
    with open(os.path.join(HERE, "coefficients.json"), encoding="utf-8") as f:
        g = json.load(f)["g"]
    table = []
    for k in range(1, ORDER + 1):
        terms = {power: num / den for power, num, den in g[str(k)]["terms"]}
        degree = max(terms)
        table.append([terms.get(p, 0.0) for p in range(degree, -1, -1)])
    return table


COEFFICIENTS = _coefficients()


def route(t, df, upper_tail=vba_upper_tail):
    """Order-5 large-df tail in binary64, with Q from `upper_tail`."""
    series = 0.0
    for k in range(ORDER, 0, -1):
        g = 0.0
        for c in COEFFICIENTS[k - 1]:
            g = g * t + c
        series = series / df + g
    series = series / df
    phi = INV_SQRT_TWO_PI * math.exp(-0.5 * t * t)
    return upper_tail(t) + phi * series


def oracle_tail(t, df):
    """Primary oracle of the study: the upper tail directly from I_z(df/2, 1/2)."""
    with mp.workdps(60):
        t, df = mp.mpf(t), mp.mpf(df)
        return mp.betainc(df / 2, mp.mpf(1) / 2, 0, df / (df + t * t), regularized=True) / 2


def relative_error(value, reference):
    with mp.workdps(60):
        return float(abs((mp.mpf(value) - reference) / reference))


def recorded_cf_errors():
    """C(df, t): the study's measured error of the current binary64 CF route."""
    with open(os.path.join(HERE, "results.json"), encoding="utf-8") as f:
        fitting = json.load(f)["forward_fitting"]
    return {(p["t"], p["df"]): float(p["C"]) for p in fitting}


def feasibility():
    """Every forward fitting point inside the adopted region, three routes."""
    current = recorded_cf_errors()
    rows = []
    for (t, df), cf in sorted(current.items()):
        if df < DF_MIN or t > T_MAX:
            continue
        reference = oracle_tail(t, df)
        rows.append({
            "t": t, "df": df,
            "current_cf": cf,
            "route_vba_kernel": relative_error(route(t, df, vba_upper_tail), reference),
            "route_exact_q": relative_error(route(t, df, exact_upper_tail), reference),
        })
    return rows


def main():
    rows = feasibility()
    print(f"{'t':>4} {'df':>7} {'current CF':>11} {'route+VBA Q':>12} {'route+exact Q':>14}")
    for r in rows:
        flag = "  WORSE" if r["route_vba_kernel"] > max(r["current_cf"], BOUND) else ""
        print(f"{r['t']:4} {r['df']:7.0e} {r['current_cf']:11.2e} {r['route_vba_kernel']:12.2e} "
              f"{r['route_exact_q']:14.2e}{flag}")


if __name__ == "__main__":
    main()
