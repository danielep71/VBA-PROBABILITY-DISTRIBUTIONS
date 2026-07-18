#!/usr/bin/env python3
"""
generate_reference_values.py
============================================================================
Phase 1 of the reproducible accuracy harness for VBA-PROBABILITY-DISTRIBUTIONS.

Emits a grid of inputs and high-precision REFERENCE values for every function
that publishes a measured-accuracy claim in the VBA source:

  SPECIALFUNCS kernels
    PROB_LogGamma           rel err < 6.1E-14  for Z in [1E-8, 1E+50]
    PROB_LogGammaHalfDiff   rel err <= 2.1E-15 for Z > 0
    PROB_StirlingError      abs err <= 3E-17   for N >= 0.5
    PROB_LogChoose          rel err <= 3.2E-16 for N in [2, 2^53], all K

  TFAMILY public UDFs
    StudentT density / cumulative / survival / quantile
    ChiSquare cumulative / survival / quantile
    F cumulative / survival / quantile

The reference column is computed with mpmath at 50 decimal digits. The
`observed_vba` column is left EMPTY: it is filled by the companion VBA export
macro (M_STATS_PROBDIST_ACCURACY_EXPORT.bas) running inside Excel, because the
library under test is VBA and cannot be executed from Python. compute_errors.py
then joins the two and produces accuracy_summary.md.

Usage:
    python generate_reference_values.py            # writes probability_accuracy_grid.csv
    python generate_reference_values.py --digits 60
"""
import argparse
import csv
import datetime as _dt

import mpmath as mp


def _loggamma(z):
    return mp.log(mp.gamma(z))


def _loggamma_halfdiff(z):
    # Log(Gamma(Z + 1/2)) - Log(Gamma(Z)), no cancellation in the reference
    return mp.log(mp.gamma(z + mp.mpf(1) / 2)) - mp.log(mp.gamma(z))


def _stirling_error(n):
    # delta(N) = logGamma(N+1) - [ N*log(N) - N + 0.5*log(2*pi*N) ]
    n = mp.mpf(n)
    return mp.log(mp.gamma(n + 1)) - (
        n * mp.log(n) - n + mp.mpf("0.5") * mp.log(2 * mp.pi * n)
    )


def _logchoose(n, k):
    # Log(C(N,K)) = logGamma(N+1) - logGamma(K+1) - logGamma(N-K+1)
    n, k = mp.mpf(n), mp.mpf(k)
    return mp.log(mp.gamma(n + 1)) - mp.log(mp.gamma(k + 1)) - mp.log(mp.gamma(n - k + 1))


def _student_t_pdf(x, df):
    x, df = mp.mpf(x), mp.mpf(df)
    c = mp.gamma((df + 1) / 2) / (mp.sqrt(df * mp.pi) * mp.gamma(df / 2))
    return c * (1 + x * x / df) ** (-(df + 1) / 2)


def _student_t_cdf(x, df):
    x, df = mp.mpf(x), mp.mpf(df)
    # CDF via regularized incomplete beta; stable both tails
    xt = df / (df + x * x)
    ib = mp.betainc(df / 2, mp.mpf("0.5"), 0, xt, regularized=True) / 2
    return 1 - ib if x > 0 else ib


def _student_t_sf(x, df):
    # Direct upper tail; never 1 - cdf (which cancels for large x)
    x, df = mp.mpf(x), mp.mpf(df)
    xt = df / (df + x * x)
    half_ib = mp.betainc(df / 2, mp.mpf("0.5"), 0, xt, regularized=True) / 2
    return half_ib if x >= 0 else 1 - half_ib


def _bisect(f, a, b, tol=mp.mpf("1e-40"), maxit=400):
    a, b = mp.mpf(a), mp.mpf(b)
    fa, fb = f(a), f(b)
    if fa == 0:
        return a
    if fb == 0:
        return b
    if mp.sign(fa) == mp.sign(fb):
        raise ValueError("root not bracketed")
    for _ in range(maxit):
        m = (a + b) / 2
        fm = f(m)
        if fm == 0 or (b - a) / 2 < tol * (1 + abs(m)):
            return m
        if mp.sign(fm) == mp.sign(fa):
            a, fa = m, fm
        else:
            b, fb = m, fm
    return (a + b) / 2


def _student_t_ppf(p, df):
    p, df = mp.mpf(p), mp.mpf(df)
    if p == mp.mpf("0.5"):
        return mp.mpf(0)
    return _bisect(lambda t: _student_t_cdf(t, df) - p, mp.mpf("-1e7"), mp.mpf("1e7"))


def _chi2_cdf(x, df):
    x, df = mp.mpf(x), mp.mpf(df)
    return mp.gammainc(df / 2, 0, x / 2, regularized=True)


def _chi2_sf(x, df):
    x, df = mp.mpf(x), mp.mpf(df)
    return mp.gammainc(df / 2, x / 2, mp.inf, regularized=True)


def _chi2_ppf(p, df):
    p, df = mp.mpf(p), mp.mpf(df)
    hi = df * 100 + mp.mpf(1000)
    return _bisect(lambda x: _chi2_cdf(x, df) - p, mp.mpf("1e-30"), hi)


def _f_cdf(x, d1, d2):
    x, d1, d2 = mp.mpf(x), mp.mpf(d1), mp.mpf(d2)
    xt = d1 * x / (d1 * x + d2)
    return mp.betainc(d1 / 2, d2 / 2, 0, xt, regularized=True)


def _f_sf(x, d1, d2):
    # Direct upper tail via the beta symmetry 1 - I_x(a,b) = I_{1-x}(b,a)
    x, d1, d2 = mp.mpf(x), mp.mpf(d1), mp.mpf(d2)
    comp = d2 / (d1 * x + d2)  # = 1 - xt
    return mp.betainc(d2 / 2, d1 / 2, 0, comp, regularized=True)


def _f_ppf(p, d1, d2):
    p, d1, d2 = mp.mpf(p), mp.mpf(d1), mp.mpf(d2)
    return _bisect(lambda x: _f_cdf(x, d1, d2) - p, mp.mpf("1e-30"), mp.mpf("1e7"))


def logspace(lo, hi, n):
    lo, hi = mp.mpf(lo), mp.mpf(hi)
    return [mp.e ** (mp.log(lo) + (mp.log(hi) - mp.log(lo)) * i / (n - 1)) for i in range(n)]


def build_rows():
    rows = []

    def add(func, vba_kernel, claim, metric, args, ref):
        rows.append(
            {
                "function": func,
                "vba_kernel": vba_kernel,
                "claim": claim,
                "metric": metric,
                "arg1": mp.nstr(args[0], 17) if len(args) > 0 else "",
                "arg2": mp.nstr(args[1], 17) if len(args) > 1 else "",
                "arg3": mp.nstr(args[2], 17) if len(args) > 2 else "",
                "reference": mp.nstr(ref, 25),
                "observed_vba": "",
            }
        )

    # --- PROB_LogGamma : Z in [1E-8, 1E+50], rel < 6.1E-14 ---
    for z in logspace("1e-8", "1e50", 40):
        add("LogGamma", "PROB_LogGamma", "rel<6.1E-14", "rel", (z,), _loggamma(z))

    # --- PROB_LogGammaHalfDiff : Z > 0, rel <= 2.1E-15 ---
    for z in logspace("1e-6", "1e12", 30):
        add("LogGammaHalfDiff", "PROB_LogGammaHalfDiff", "rel<=2.1E-15", "rel", (z,), _loggamma_halfdiff(z))

    # --- PROB_StirlingError : N >= 0.5, abs <= 3E-17 (include N=501 hot spot) ---
    ns = [mp.mpf("0.5"), mp.mpf(1), mp.mpf(2), mp.mpf(3), mp.mpf(5), mp.mpf(10),
          mp.mpf(50), mp.mpf(100), mp.mpf(500), mp.mpf(501), mp.mpf(1000), mp.mpf(1e6)]
    for n in ns:
        add("StirlingError", "PROB_StirlingError", "abs<=3E-17", "abs", (n,), _stirling_error(n))

    # --- PROB_LogChoose : N in [2, 2^53], all K, rel <= 3.2E-16 ---
    for n in [mp.mpf(2), mp.mpf(10), mp.mpf(100), mp.mpf(1030), mp.mpf(1e6), mp.mpf(2) ** 53]:
        for frac in [mp.mpf("0.0"), mp.mpf("0.01"), mp.mpf("0.5"), mp.mpf("0.99"), mp.mpf("1.0")]:
            k = mp.floor(n * frac)
            add("LogChoose", "PROB_LogChoose", "rel<=3.2E-16", "rel", (n, k), _logchoose(n, k))

    # --- Student t ---
    for df in [mp.mpf(1), mp.mpf(2), mp.mpf(5), mp.mpf(30), mp.mpf(1000)]:
        for x in [mp.mpf("0.1"), mp.mpf(1), mp.mpf(2), mp.mpf(5), mp.mpf(20)]:
            add("StudentT_Density", "K_STATS_StudentT_Density", "rel<=8.4E-15", "rel", (x, df), _student_t_pdf(x, df))
            add("StudentT_Cumulative", "K_STATS_StudentT_Cumulative", "rel<=1.3E-12", "rel", (x, df), _student_t_cdf(x, df))
            add("StudentT_Survival", "K_STATS_StudentT_Survival", "rel<=1.3E-12", "rel", (x, df), _student_t_sf(x, df))
        for p in [mp.mpf("0.001"), mp.mpf("0.05"), mp.mpf("0.5"), mp.mpf("0.95"), mp.mpf("0.999")]:
            add("StudentT_InverseCumulative", "K_STATS_StudentT_InverseCumulative", "rel<=3.0E-12", "rel", (p, df), _student_t_ppf(p, df))

    # --- Chi-square ---
    for df in [mp.mpf(1), mp.mpf(2), mp.mpf(5), mp.mpf(30), mp.mpf(100)]:
        for x in [mp.mpf("0.5"), mp.mpf(1), mp.mpf(5), mp.mpf(20), mp.mpf(80)]:
            add("ChiSquare_Cumulative", "K_STATS_ChiSquare_Cumulative", "rel<=2.6E-10", "rel", (x, df), _chi2_cdf(x, df))
            add("ChiSquare_Survival", "K_STATS_ChiSquare_Survival", "rel<=2.6E-10", "rel", (x, df), _chi2_sf(x, df))
        for p in [mp.mpf("0.001"), mp.mpf("0.05"), mp.mpf("0.5"), mp.mpf("0.95"), mp.mpf("0.999")]:
            add("ChiSquare_InverseCumulative", "K_STATS_ChiSquare_InverseCumulative", "rel<=4.7E-12", "rel", (p, df), _chi2_ppf(p, df))

    # --- F ---
    for d1, d2 in [(mp.mpf(1), mp.mpf(1)), (mp.mpf(5), mp.mpf(2)), (mp.mpf(10), mp.mpf(30)), (mp.mpf(100), mp.mpf(100))]:
        for x in [mp.mpf("0.25"), mp.mpf(1), mp.mpf(2), mp.mpf(10)]:
            add("F_Cumulative", "K_STATS_F_Cumulative", "rel<=1.1E-10", "rel", (x, d1, d2), _f_cdf(x, d1, d2))
            add("F_Survival", "K_STATS_F_Survival", "rel<=1.1E-10", "rel", (x, d1, d2), _f_sf(x, d1, d2))
        for p in [mp.mpf("0.05"), mp.mpf("0.5"), mp.mpf("0.95")]:
            add("F_InverseCumulative", "K_STATS_F_InverseCumulative", "rel<=5.9E-13", "rel", (p, d1, d2), _f_ppf(p, d1, d2))

    return rows


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--digits", type=int, default=50)
    ap.add_argument("--out", default="probability_accuracy_grid.csv")
    args = ap.parse_args()

    mp.mp.dps = args.digits
    rows = build_rows()

    fields = ["function", "vba_kernel", "claim", "metric",
              "arg1", "arg2", "arg3", "reference", "observed_vba"]
    with open(args.out, "w", newline="") as f:
        w = csv.DictWriter(f, fieldnames=fields)
        w.writeheader()
        w.writerows(rows)

    print(f"wrote {args.out}: {len(rows)} reference rows at {args.digits} digits")
    print(f"generated {_dt.date.today().isoformat()}")


if __name__ == "__main__":
    main()
