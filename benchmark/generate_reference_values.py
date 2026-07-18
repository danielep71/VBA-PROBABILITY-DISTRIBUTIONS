#!/usr/bin/env python3
"""
generate_reference_values.py
============================================================================
Phase 1 of the reproducible accuracy harness for VBA-PROBABILITY-DISTRIBUTIONS.

Emits a grid of inputs and high-precision REFERENCE values for every function
that publishes a measured-accuracy claim in the VBA source:

  SPECIALFUNCS kernels
    PROB_LogGamma           rel err < 6.1E-14  for Z in [1E-8, 1E+50]
    PROB_LogGammaHalfDiff   rel err <= 2E-14 for Z > 0 (tested range)
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



def _phi(z):
    z = mp.mpf(z)
    return mp.e ** (-z * z / 2) / mp.sqrt(2 * mp.pi)


def _Phi(z):
    return mp.ncdf(mp.mpf(z))


def _Phi_sf(z):
    return mp.ncdf(-mp.mpf(z))


def _Phi_inv(p):
    p = mp.mpf(p)
    return _bisect(lambda z: _Phi(z) - p, mp.mpf("-40"), mp.mpf("40"))


def _lognorm_params(mean, sd):
    mean, sd = mp.mpf(mean), mp.mpf(sd)
    varlog = mp.log(1 + (sd / mean) ** 2)
    return mp.log(mean) - varlog / 2, mp.sqrt(varlog)



def _gamma_pdf(x, k, th):
    x, k, th = mp.mpf(x), mp.mpf(k), mp.mpf(th)
    return x ** (k - 1) * mp.e ** (-x / th) / (th ** k * mp.gamma(k))


def _gamma_cdf(x, k, th):
    return mp.gammainc(mp.mpf(k), 0, mp.mpf(x) / mp.mpf(th), regularized=True)


def _gamma_sf(x, k, th):
    return mp.gammainc(mp.mpf(k), mp.mpf(x) / mp.mpf(th), mp.inf, regularized=True)


def _gamma_ppf(pq, k, th):
    k, th = mp.mpf(k), mp.mpf(th)
    return _bisect(lambda x: _gamma_cdf(x, k, th) - mp.mpf(pq), mp.mpf("1e-30"), k * th * 100 + 1000) 


def _beta_pdf(x, a, b):
    x, a, b = mp.mpf(x), mp.mpf(a), mp.mpf(b)
    return x ** (a - 1) * (1 - x) ** (b - 1) / mp.beta(a, b)


def _beta_cdf(x, a, b):
    return mp.betainc(mp.mpf(a), mp.mpf(b), 0, mp.mpf(x), regularized=True)


def _beta_sf(x, a, b):
    return mp.betainc(mp.mpf(a), mp.mpf(b), mp.mpf(x), 1, regularized=True)


def _beta_ppf(pq, a, b):
    a, b = mp.mpf(a), mp.mpf(b)
    return _bisect(lambda x: _beta_cdf(x, a, b) - mp.mpf(pq), mp.mpf("1e-30"), mp.mpf(1) - mp.mpf("1e-30"))


def _weibull_mean(k, lam):
    k, lam = mp.mpf(k), mp.mpf(lam)
    return lam * mp.gamma(1 + 1 / k)


def _weibull_var(k, lam):
    k, lam = mp.mpf(k), mp.mpf(lam)
    return lam ** 2 * (mp.gamma(1 + 2 / k) - mp.gamma(1 + 1 / k) ** 2)


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

    # --- PROB_LogGammaHalfDiff : Z > 0, rel <= 2E-14 (tested range) ---
    for z in logspace("1e-6", "1e12", 30):
        add("LogGammaHalfDiff", "PROB_LogGammaHalfDiff", "rel<=2E-14", "rel", (z,), _loggamma_halfdiff(z))

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
            add("StudentT_Density", "K_STATS_StudentT_Density", "rel<=2E-14", "rel", (x, df), _student_t_pdf(x, df))
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


    # ===================== NORMAL FAMILY =====================
    FIVE_E15 = "rel<=5E-15"

    # --- Standard Normal ---
    for z in [mp.mpf("-2"), mp.mpf("-0.5"), mp.mpf("0.5"), mp.mpf(1), mp.mpf("1.96"), mp.mpf(3)]:
        add("NormalStandard_Density", "K_STATS_NormalStandard_Density", FIVE_E15, "rel", (z,), _phi(z))
        add("NormalStandard_Cumulative", "K_STATS_NormalStandard_Cumulative", FIVE_E15, "rel", (z,), _Phi(z))
        add("NormalStandard_Survival", "K_STATS_NormalStandard_Survival", FIVE_E15, "rel", (z,), _Phi_sf(z))
    for pq in [mp.mpf("0.01"), mp.mpf("0.25"), mp.mpf("0.5"), mp.mpf("0.975"), mp.mpf("0.999")]:
        add("NormalStandard_InverseCumulative", "K_STATS_NormalStandard_InverseCumulative", FIVE_E15, "rel", (pq,), _Phi_inv(pq))
        add("NormalStandard_InverseSurvival", "K_STATS_NormalStandard_InverseSurvival", FIVE_E15, "rel", (pq,), -_Phi_inv(pq))
        add("NormalStandard_InverseCumulativeFast", "K_STATS_NormalStandard_InverseCumulativeFast", "rel<=5E-9", "rel", (pq,), _Phi_inv(pq))
    for lo, up in [(mp.mpf("-1.96"), mp.mpf("1.96")), (mp.mpf("-1"), mp.mpf(2)), (mp.mpf("0"), mp.mpf(3))]:
        add("NormalStandard_IntervalProbability", "K_STATS_NormalStandard_IntervalProbability", FIVE_E15, "rel", (lo, up), _Phi(up) - _Phi(lo))

    # --- General Normal ---
    for (x, m, sd) in [(mp.mpf("1.96"), mp.mpf(0), mp.mpf(1)), (mp.mpf(110), mp.mpf(100), mp.mpf(15)),
                       (mp.mpf(3), mp.mpf(5), mp.mpf(2))]:
        z = (x - m) / sd
        add("Normal_Density", "K_STATS_Normal_Density", FIVE_E15, "rel", (x, m, sd), _phi(z) / sd)
        add("Normal_Cumulative", "K_STATS_Normal_Cumulative", FIVE_E15, "rel", (x, m, sd), _Phi(z))
        add("Normal_Survival", "K_STATS_Normal_Survival", FIVE_E15, "rel", (x, m, sd), _Phi_sf(z))
        add("Normal_ZScore", "K_STATS_Normal_ZScore", FIVE_E15, "rel", (x, m, sd), z)
    for (pq, m, sd) in [(mp.mpf("0.99"), mp.mpf(100), mp.mpf(15)), (mp.mpf("0.025"), mp.mpf(10), mp.mpf(2))]:
        add("Normal_InverseCumulative", "K_STATS_Normal_InverseCumulative", FIVE_E15, "rel", (pq, m, sd), m + sd * _Phi_inv(pq))
        add("Normal_InverseSurvival", "K_STATS_Normal_InverseSurvival", FIVE_E15, "rel", (pq, m, sd), m - sd * _Phi_inv(pq))

    # --- Lognormal ---
    for (x, ml, sl) in [(mp.mpf(1), mp.mpf(0), mp.mpf(1)), (mp.mpf(2), mp.mpf("0.5"), mp.mpf("0.25")),
                        (mp.mpf("0.5"), mp.mpf(0), mp.mpf(1))]:
        zz = (mp.log(x) - ml) / sl
        add("Lognormal_Density", "K_STATS_Lognormal_Density", FIVE_E15, "rel", (x, ml, sl), _phi(zz) / (x * sl))
        add("Lognormal_Cumulative", "K_STATS_Lognormal_Cumulative", FIVE_E15, "rel", (x, ml, sl), _Phi(zz))
        add("Lognormal_Survival", "K_STATS_Lognormal_Survival", FIVE_E15, "rel", (x, ml, sl), _Phi_sf(zz))
    for (pq, ml, sl) in [(mp.mpf("0.5"), mp.mpf(0), mp.mpf(1)), (mp.mpf("0.025"), mp.mpf(0), mp.mpf(1))]:
        add("Lognormal_InverseCumulative", "K_STATS_Lognormal_InverseCumulative", FIVE_E15, "rel", (pq, ml, sl), mp.e ** (ml + sl * _Phi_inv(pq)))
        add("Lognormal_InverseSurvival", "K_STATS_Lognormal_InverseSurvival", FIVE_E15, "rel", (pq, ml, sl), mp.e ** (ml - sl * _Phi_inv(pq)))
    for (ml, sl) in [(mp.mpf(0), mp.mpf(1)), (mp.mpf("0.5"), mp.mpf("0.25"))]:
        add("Lognormal_Mean", "K_STATS_Lognormal_Mean", FIVE_E15, "rel", (ml, sl), mp.e ** (ml + sl * sl / 2))
        add("Lognormal_Variance", "K_STATS_Lognormal_Variance", FIVE_E15, "rel", (ml, sl), (mp.e ** (sl * sl) - 1) * mp.e ** (2 * ml + sl * sl))
        add("Lognormal_StdDev", "K_STATS_Lognormal_StdDev", FIVE_E15, "rel", (ml, sl), mp.sqrt((mp.e ** (sl * sl) - 1) * mp.e ** (2 * ml + sl * sl)))
    # ParametersFromMeanStdDev returns a 1x2 array; test each output separately
    for (mean, sd) in [(mp.mpf(2), mp.mpf("0.5")), (mp.mpf(10), mp.mpf(3))]:
        mlref, slref = _lognorm_params(mean, sd)
        add("Lognormal_ParamMeanLog", "K_STATS_Lognormal_ParametersFromMeanStdDev", FIVE_E15, "rel", (mean, sd), mlref)
        add("Lognormal_ParamStdDevLog", "K_STATS_Lognormal_ParametersFromMeanStdDev", FIVE_E15, "rel", (mean, sd), slref)


    # ===================== CONTINUOUS FAMILY =====================
    # Bounds set from measured worst-case error over the tested grid (5E-15 for
    # near-machine-epsilon functions, 2E-14 where a digit is lost). Exponential is
    # parameterized by RATE (Lambda), not scale.
    PROV = "rel<=1E-8"

    # --- Gamma(X, Shape k, ScaleParam theta) ---
    for (x, k, th) in [(mp.mpf(2), mp.mpf(2), mp.mpf(1)), (mp.mpf(5), mp.mpf(3), mp.mpf(2)),
                       (mp.mpf("0.5"), mp.mpf("1.5"), mp.mpf(1))]:
        add("Gamma_Density", "K_STATS_Gamma_Density", "rel<=2E-14", "rel", (x, k, th), _gamma_pdf(x, k, th))
        add("Gamma_Cumulative", "K_STATS_Gamma_Cumulative", "rel<=2E-14", "rel", (x, k, th), _gamma_cdf(x, k, th))
        add("Gamma_Survival", "K_STATS_Gamma_Survival", "rel<=2E-14", "rel", (x, k, th), _gamma_sf(x, k, th))
    for (pq, k, th) in [(mp.mpf("0.5"), mp.mpf(2), mp.mpf(1)), (mp.mpf("0.95"), mp.mpf(3), mp.mpf(2))]:
        add("Gamma_InverseCumulative", "K_STATS_Gamma_InverseCumulative", "rel<=2E-14", "rel", (pq, k, th), _gamma_ppf(pq, k, th))
    for (k, th) in [(mp.mpf(2), mp.mpf(3)), (mp.mpf(5), mp.mpf(2))]:
        add("Gamma_Mean", "K_STATS_Gamma_Mean", "rel<=5E-15", "rel", (k, th), mp.mpf(k) * mp.mpf(th))
        add("Gamma_Variance", "K_STATS_Gamma_Variance", "rel<=5E-15", "rel", (k, th), mp.mpf(k) * mp.mpf(th) ** 2)
        add("Gamma_StdDev", "K_STATS_Gamma_StdDev", "rel<=5E-15", "rel", (k, th), mp.sqrt(mp.mpf(k)) * mp.mpf(th))

    # --- Beta(X, Alpha a, Beta b) ---
    for (x, a, b) in [(mp.mpf("0.5"), mp.mpf(2), mp.mpf(2)), (mp.mpf("0.3"), mp.mpf(2), mp.mpf(5)),
                      (mp.mpf("0.8"), mp.mpf(5), mp.mpf(1))]:
        add("Beta_Density", "K_STATS_Beta_Density", "rel<=5E-15", "rel", (x, a, b), _beta_pdf(x, a, b))
        add("Beta_Cumulative", "K_STATS_Beta_Cumulative", "rel<=2E-14", "rel", (x, a, b), _beta_cdf(x, a, b))
        add("Beta_Survival", "K_STATS_Beta_Survival", "rel<=5E-15", "rel", (x, a, b), _beta_sf(x, a, b))
    for (pq, a, b) in [(mp.mpf("0.5"), mp.mpf(2), mp.mpf(2)), (mp.mpf("0.95"), mp.mpf(2), mp.mpf(5))]:
        add("Beta_InverseCumulative", "K_STATS_Beta_InverseCumulative", "rel<=5E-15", "rel", (pq, a, b), _beta_ppf(pq, a, b))
    for (a, b) in [(mp.mpf(2), mp.mpf(3)), (mp.mpf(5), mp.mpf(2))]:
        add("Beta_Mean", "K_STATS_Beta_Mean", "rel<=5E-15", "rel", (a, b), mp.mpf(a) / (mp.mpf(a) + mp.mpf(b)))
        add("Beta_Variance", "K_STATS_Beta_Variance", "rel<=5E-15", "rel", (a, b), mp.mpf(a) * mp.mpf(b) / ((mp.mpf(a) + mp.mpf(b)) ** 2 * (mp.mpf(a) + mp.mpf(b) + 1)))
        add("Beta_StdDev", "K_STATS_Beta_StdDev", "rel<=5E-15", "rel", (a, b), mp.sqrt(mp.mpf(a) * mp.mpf(b) / ((mp.mpf(a) + mp.mpf(b)) ** 2 * (mp.mpf(a) + mp.mpf(b) + 1))))

    # --- Exponential(X, Lambda=rate) ---
    for (x, lam) in [(mp.mpf(1), mp.mpf(1)), (mp.mpf("0.5"), mp.mpf(2)), (mp.mpf(3), mp.mpf("0.5"))]:
        lam = mp.mpf(lam)
        add("Exponential_Density", "K_STATS_Exponential_Density", "rel<=5E-15", "rel", (x, lam), lam * mp.e ** (-lam * mp.mpf(x)))
        add("Exponential_Cumulative", "K_STATS_Exponential_Cumulative", "rel<=5E-15", "rel", (x, lam), 1 - mp.e ** (-lam * mp.mpf(x)))
        add("Exponential_Survival", "K_STATS_Exponential_Survival", "rel<=5E-15", "rel", (x, lam), mp.e ** (-lam * mp.mpf(x)))
    for (pq, lam) in [(mp.mpf("0.5"), mp.mpf(1)), (mp.mpf("0.95"), mp.mpf(2))]:
        add("Exponential_InverseCumulative", "K_STATS_Exponential_InverseCumulative", "rel<=5E-15", "rel", (pq, lam), -mp.log(1 - mp.mpf(pq)) / mp.mpf(lam))

    # --- Weibull(X, Shape k, ScaleParam lam) ---
    for (x, k, lam) in [(mp.mpf(1), mp.mpf("1.5"), mp.mpf(1)), (mp.mpf(2), mp.mpf(2), mp.mpf(2)),
                        (mp.mpf("0.5"), mp.mpf(3), mp.mpf(1))]:
        k2, lam2 = mp.mpf(k), mp.mpf(lam)
        add("Weibull_Density", "K_STATS_Weibull_Density", "rel<=5E-15", "rel", (x, k, lam), (k2 / lam2) * (mp.mpf(x) / lam2) ** (k2 - 1) * mp.e ** (-(mp.mpf(x) / lam2) ** k2))
        add("Weibull_Cumulative", "K_STATS_Weibull_Cumulative", "rel<=5E-15", "rel", (x, k, lam), 1 - mp.e ** (-(mp.mpf(x) / lam2) ** k2))
        add("Weibull_Survival", "K_STATS_Weibull_Survival", "rel<=5E-15", "rel", (x, k, lam), mp.e ** (-(mp.mpf(x) / lam2) ** k2))
    for (pq, k, lam) in [(mp.mpf("0.5"), mp.mpf("1.5"), mp.mpf(1)), (mp.mpf("0.95"), mp.mpf(2), mp.mpf(2))]:
        add("Weibull_InverseCumulative", "K_STATS_Weibull_InverseCumulative", "rel<=5E-15", "rel", (pq, k, lam), mp.mpf(lam) * (-mp.log(1 - mp.mpf(pq))) ** (1 / mp.mpf(k)))
    for (k, lam) in [(mp.mpf("1.5"), mp.mpf(1)), (mp.mpf(2), mp.mpf(2))]:
        add("Weibull_Mean", "K_STATS_Weibull_Mean", "rel<=2E-14", "rel", (k, lam), _weibull_mean(k, lam))
        add("Weibull_Variance", "K_STATS_Weibull_Variance", "rel<=2E-14", "rel", (k, lam), _weibull_var(k, lam))
        add("Weibull_StdDev", "K_STATS_Weibull_StdDev", "rel<=5E-15", "rel", (k, lam), mp.sqrt(_weibull_var(k, lam)))

    # --- Uniform(X, LowerBound a, UpperBound b) ---
    for (x, a, b) in [(mp.mpf(3), mp.mpf(0), mp.mpf(10)), (mp.mpf("2.5"), mp.mpf(1), mp.mpf(4))]:
        a2, b2 = mp.mpf(a), mp.mpf(b)
        add("Uniform_Density", "K_STATS_Uniform_Density", "rel<=5E-15", "rel", (x, a, b), 1 / (b2 - a2))
        add("Uniform_Cumulative", "K_STATS_Uniform_Cumulative", "rel<=5E-15", "rel", (x, a, b), (mp.mpf(x) - a2) / (b2 - a2))
        add("Uniform_Survival", "K_STATS_Uniform_Survival", "rel<=5E-15", "rel", (x, a, b), (b2 - mp.mpf(x)) / (b2 - a2))
    for (pq, a, b) in [(mp.mpf("0.5"), mp.mpf(0), mp.mpf(10)), (mp.mpf("0.9"), mp.mpf(1), mp.mpf(4))]:
        add("Uniform_InverseCumulative", "K_STATS_Uniform_InverseCumulative", "rel<=5E-15", "rel", (pq, a, b), mp.mpf(a) + mp.mpf(pq) * (mp.mpf(b) - mp.mpf(a)))

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
