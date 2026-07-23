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


_BETA_BAL = {"Beta_Density", "Beta_Cumulative", "Beta_Survival", "Beta_InverseCumulative"}
_F_VAL = {"F_Cumulative", "F_Survival", "F_InverseCumulative"}


def _regime_for(func):
    if func in _BETA_BAL:
        return "balanced"
    if func in _F_VAL:
        return "validated"
    return "all"


def _load_contracts(path=None):
    """Load the regime-aware contract, keyed by (function, regime)."""
    import csv, os
    if path is None:
        path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "accuracy_contracts.csv")
    contracts = {}
    with open(path, newline="") as f:
        for row in csv.DictReader(f):
            m = "rel" if row["metric"].strip().lower().startswith("rel") else "abs"
            contracts[(row["function"], row["regime"])] = {
                "metric": m, "claim": f'{m}<={row["threshold"].strip()}'}
    return contracts


_CONTRACTS = _load_contracts()



# ===========================================================================
# DISCRETE FAMILY (Binomial / Poisson / Geometric)
#   Self-contained so it can be generated independently of the contract-backed
#   families. CDF/survival go through the same regularized incomplete beta /
#   gamma the VBA kernels use, so large n and large mean are exercised. Rows
#   carry a provisional claim and empty observed_vba until the Excel export and
#   holdout freeze (Phase 2); no discrete contract is active yet, so the strict
#   accuracy gate does not evaluate them.
# ===========================================================================
def _betacf(a, b, x):
    tiny = mp.mpf("1e-300"); qab = a + b; qap = a + 1; qam = a - 1
    c = mp.mpf(1); d = 1 - qab * x / qap
    if abs(d) < tiny: d = tiny
    d = 1 / d; h = d
    for m in range(1, 20000):
        m2 = 2 * m
        aa = m * (b - m) * x / ((qam + m2) * (a + m2))
        d = 1 + aa * d
        if abs(d) < tiny: d = tiny
        c = 1 + aa / c
        if abs(c) < tiny: c = tiny
        d = 1 / d; h *= d * c
        aa = -(a + m) * (qab + m) * x / ((a + m2) * (qap + m2))
        d = 1 + aa * d
        if abs(d) < tiny: d = tiny
        c = 1 + aa / c
        if abs(c) < tiny: c = tiny
        d = 1 / d; de = d * c; h *= de
        if abs(de - 1) < mp.mpf("1e-45"): break
    return h


def _ibeta(x, a, b):
    x, a, b = mp.mpf(x), mp.mpf(a), mp.mpf(b)
    if x <= 0: return mp.mpf(0)
    if x >= 1: return mp.mpf(1)
    lbt = mp.loggamma(a + b) - mp.loggamma(a) - mp.loggamma(b) + a * mp.log(x) + b * mp.log(1 - x)
    bt = mp.e ** lbt
    if x < (a + 1) / (a + b + 2): return bt * _betacf(a, b, x) / a
    return 1 - bt * _betacf(b, a, 1 - x) / b


def _binom_pmf(k, n, pr):
    k, n, pr = mp.mpf(k), mp.mpf(n), mp.mpf(pr)
    return mp.binomial(n, k) * pr ** k * (1 - pr) ** (n - k)


def _binom_logpmf(k, n, pr):
    k, n, pr = mp.mpf(k), mp.mpf(n), mp.mpf(pr)
    return mp.log(mp.binomial(n, k)) + k * mp.log(pr) + (n - k) * mp.log(1 - pr)


def _binom_cdf(k, n, pr):
    k, n, pr = mp.mpf(k), mp.mpf(n), mp.mpf(pr)
    if k >= n:
        return mp.mpf(1)
    return _ibeta(1 - pr, n - k, k + 1)


def _binom_sf(k, n, pr):
    k, n, pr = mp.mpf(k), mp.mpf(n), mp.mpf(pr)
    if k >= n:
        return mp.mpf(0)
    return _ibeta(pr, k + 1, n - k)


def _binom_inv(prob, n, pr):
    prob = mp.mpf(prob)
    lo, hi = -1, int(n)
    while hi - lo > 1:
        mid = (lo + hi) // 2
        if _binom_cdf(mid, n, pr) >= prob:
            hi = mid
        else:
            lo = mid
    return mp.mpf(hi)


def _pois_pmf(k, lam):
    k, lam = mp.mpf(k), mp.mpf(lam)
    return mp.e ** (k * mp.log(lam) - lam - mp.loggamma(k + 1))


def _pois_logpmf(k, lam):
    k, lam = mp.mpf(k), mp.mpf(lam)
    return k * mp.log(lam) - lam - mp.loggamma(k + 1)


def _pois_cdf(k, lam):
    return mp.gammainc(mp.mpf(k) + 1, mp.mpf(lam), mp.inf, regularized=True)


def _pois_sf(k, lam):
    return mp.gammainc(mp.mpf(k) + 1, 0, mp.mpf(lam), regularized=True)


def _pois_inv(prob, lam):
    prob = mp.mpf(prob)
    lo = -1
    hi = int(mp.floor(mp.mpf(lam) + 12 * mp.sqrt(mp.mpf(lam)) + 40))
    while hi - lo > 1:
        mid = (lo + hi) // 2
        if _pois_cdf(mid, lam) >= prob:
            hi = mid
        else:
            lo = mid
    return mp.mpf(hi)


def _geo_pmf(k, pr):
    k, pr = mp.mpf(k), mp.mpf(pr)
    return pr * (1 - pr) ** k


def _geo_logpmf(k, pr):
    k, pr = mp.mpf(k), mp.mpf(pr)
    return mp.log(pr) + k * mp.log(1 - pr)


def _geo_cdf(k, pr):
    k, pr = mp.mpf(k), mp.mpf(pr)
    return 1 - (1 - pr) ** (k + 1)


def _geo_sf(k, pr):
    k, pr = mp.mpf(k), mp.mpf(pr)
    return (1 - pr) ** (k + 1)


def _geo_inv(prob, pr):
    prob, pr = mp.mpf(prob), mp.mpf(pr)
    k = int(mp.ceil(mp.log(1 - prob) / mp.log(1 - pr) - 1))
    if k < 0:
        k = 0
    while k > 0 and _geo_cdf(k - 1, pr) >= prob:
        k -= 1
    while _geo_cdf(k, pr) < prob:
        k += 1
    return mp.mpf(k)



def _nb_pmf(k, r, pr):
    k, r, pr = mp.mpf(k), mp.mpf(r), mp.mpf(pr)
    return mp.binomial(k + r - 1, k) * pr ** r * (1 - pr) ** k


def _nb_logpmf(k, r, pr):
    k, r, pr = mp.mpf(k), mp.mpf(r), mp.mpf(pr)
    return mp.log(mp.binomial(k + r - 1, k)) + r * mp.log(pr) + k * mp.log(1 - pr)


def _nb_cdf(k, r, pr):
    return _ibeta(mp.mpf(pr), mp.mpf(r), mp.mpf(k) + 1)


def _nb_sf(k, r, pr):
    return _ibeta(1 - mp.mpf(pr), mp.mpf(k) + 1, mp.mpf(r))


def _nb_inv(prob, r, pr):
    prob = mp.mpf(prob)
    mean = float(r) * (1 - float(pr)) / float(pr)
    sd = mean ** 0.5 / float(pr) if pr < 1 else 1.0
    lo, hi = -1, int(mean + 40 * sd + 60)
    while hi - lo > 1:
        m = (lo + hi) // 2
        if _nb_cdf(m, r, pr) >= prob:
            hi = m
        else:
            lo = m
    return mp.mpf(hi)


def _hy_pmf(k, n, K, N):
    k, n, K, N = mp.mpf(k), mp.mpf(n), mp.mpf(K), mp.mpf(N)
    return mp.binomial(K, k) * mp.binomial(N - K, n - k) / mp.binomial(N, n)


def _hy_logpmf(k, n, K, N):
    k, n, K, N = mp.mpf(k), mp.mpf(n), mp.mpf(K), mp.mpf(N)
    return (mp.log(mp.binomial(K, k)) + mp.log(mp.binomial(N - K, n - k))
            - mp.log(mp.binomial(N, n)))


def _hy_cdf(k, n, K, N):
    lo = max(0, int(n) + int(K) - int(N))
    return sum((_hy_pmf(j, n, K, N) for j in range(lo, int(k) + 1)), mp.mpf(0))


def _hy_sf(k, n, K, N):
    hi = min(int(n), int(K))
    return sum((_hy_pmf(j, n, K, N) for j in range(int(k) + 1, hi + 1)), mp.mpf(0))


def _hy_inv(prob, n, K, N):
    prob = mp.mpf(prob)
    lo = max(0, int(n) + int(K) - int(N)) - 1
    hi = min(int(n), int(K))
    cum = mp.mpf(0); k = lo
    for j in range(max(0, int(n) + int(K) - int(N)), hi + 1):
        cum += _hy_pmf(j, n, K, N)
        if cum >= prob:
            return mp.mpf(j)
    return mp.mpf(hi)



def _du_n(lo, hi):
    return mp.mpf(hi) - mp.mpf(lo) + 1


def _du_pmf(x, lo, hi):
    if x != int(x) or x < lo or x > hi:
        return mp.mpf(0)
    return mp.mpf(1) / _du_n(lo, hi)


def _du_logpmf(lo, hi):
    return -mp.log(_du_n(lo, hi))


def _du_cdf(x, lo, hi):
    if x < lo:
        return mp.mpf(0)
    if x >= hi:
        return mp.mpf(1)
    return (mp.floor(mp.mpf(x)) - lo + 1) / _du_n(lo, hi)


def _du_sf(x, lo, hi):
    if x < lo:
        return mp.mpf(1)
    if x >= hi:
        return mp.mpf(0)
    return (mp.mpf(hi) - mp.floor(mp.mpf(x))) / _du_n(lo, hi)


def _du_inv(prob, lo, hi):
    # least k in the support with CDF(k) >= prob
    return mp.mpf(lo) + mp.ceil(mp.mpf(prob) * _du_n(lo, hi)) - 1


def _du_mean(lo, hi):
    return (mp.mpf(lo) + mp.mpf(hi)) / 2


def _du_var(lo, hi):
    n = _du_n(lo, hi)
    return (n - 1) * (n + 1) / 12


def build_discrete_rows():
    rows = []
    REL12, REL9, REL14, ABS9 = "rel<=1E-12", "rel<=1E-9", "rel<=1E-14", "abs<=1E-9"

    def row(func, kernel, args, ref, claim, metric):
        rows.append({
            "function": func, "vba_kernel": kernel, "claim": claim, "metric": metric,
            "arg1": mp.nstr(args[0], 17) if len(args) > 0 else "",
            "arg2": mp.nstr(args[1], 17) if len(args) > 1 else "",
            "arg3": mp.nstr(args[2], 17) if len(args) > 2 else "",
            "arg4": mp.nstr(args[3], 17) if len(args) > 3 else "",
            "reference": mp.nstr(ref, 25), "observed_vba": "",
            "regime": "all", "evidence_set": "main grid",
        })

    for n in [mp.mpf(20), mp.mpf(1000), mp.mpf(100000), mp.mpf(1000000), mp.mpf(10000000)]:
        for pr in [mp.mpf("0.02"), mp.mpf("0.5"), mp.mpf("0.9")]:
            sd = mp.sqrt(n * pr * (1 - pr))
            kmid = mp.floor(n * pr)
            ktail = mp.floor(n * pr + 3 * sd)
            if ktail > n:
                ktail = n
            for k in sorted(set([mp.mpf(kmid), mp.mpf(ktail)])):
                row("Binomial_PMF", "K_STATS_Binomial_PMF", (k, n, pr), _binom_pmf(k, n, pr), REL12, "rel")
                row("Binomial_LogPMF", "K_STATS_Binomial_LogPMF", (k, n, pr), _binom_logpmf(k, n, pr), REL12, "rel")
                row("Binomial_Cumulative", "K_STATS_Binomial_Cumulative", (k, n, pr), _binom_cdf(k, n, pr), REL9, "rel")
                row("Binomial_Survival", "K_STATS_Binomial_Survival", (k, n, pr), _binom_sf(k, n, pr), REL9, "rel")
            for prob in [mp.mpf("0.05"), mp.mpf("0.5"), mp.mpf("0.975")]:
                row("Binomial_InverseCumulative", "K_STATS_Binomial_InverseCumulative", (prob, n, pr), _binom_inv(prob, n, pr), ABS9, "abs")
            row("Binomial_Mean", "K_STATS_Binomial_Mean", (n, pr), n * pr, REL14, "rel")
            row("Binomial_Variance", "K_STATS_Binomial_Variance", (n, pr), n * pr * (1 - pr), REL14, "rel")
            row("Binomial_StdDev", "K_STATS_Binomial_StdDev", (n, pr), mp.sqrt(n * pr * (1 - pr)), REL14, "rel")

    for lam in [mp.mpf(3), mp.mpf(50), mp.mpf(1000), mp.mpf(1000000)]:
        sd = mp.sqrt(lam)
        for k in sorted(set([mp.mpf(mp.floor(lam)), mp.mpf(mp.floor(lam + 3 * sd))])):
            row("Poisson_PMF", "K_STATS_Poisson_PMF", (k, lam), _pois_pmf(k, lam), REL12, "rel")
            row("Poisson_LogPMF", "K_STATS_Poisson_LogPMF", (k, lam), _pois_logpmf(k, lam), REL12, "rel")
            row("Poisson_Cumulative", "K_STATS_Poisson_Cumulative", (k, lam), _pois_cdf(k, lam), REL9, "rel")
            row("Poisson_Survival", "K_STATS_Poisson_Survival", (k, lam), _pois_sf(k, lam), REL9, "rel")
        row("Poisson_LogPMF", "K_STATS_Poisson_LogPMF", (mp.mpf(0), lam), _pois_logpmf(0, lam), REL12, "rel")
        for prob in [mp.mpf("0.05"), mp.mpf("0.5"), mp.mpf("0.975")]:
            row("Poisson_InverseCumulative", "K_STATS_Poisson_InverseCumulative", (prob, lam), _pois_inv(prob, lam), ABS9, "abs")
        row("Poisson_Mean", "K_STATS_Poisson_Mean", (lam,), lam, REL14, "rel")
        row("Poisson_Variance", "K_STATS_Poisson_Variance", (lam,), lam, REL14, "rel")
        row("Poisson_StdDev", "K_STATS_Poisson_StdDev", (lam,), mp.sqrt(lam), REL14, "rel")

    for pr in [mp.mpf("0.5"), mp.mpf("0.05"), mp.mpf("0.001"), mp.mpf("1e-6")]:
        mean = (1 - pr) / pr
        for k in sorted(set([mp.mpf(0), mp.mpf(mp.floor(mean)), mp.mpf(mp.floor(3 * mean + 5))])):
            row("Geometric_PMF", "K_STATS_Geometric_PMF", (k, pr), _geo_pmf(k, pr), REL12, "rel")
            row("Geometric_LogPMF", "K_STATS_Geometric_LogPMF", (k, pr), _geo_logpmf(k, pr), REL12, "rel")
            row("Geometric_Cumulative", "K_STATS_Geometric_Cumulative", (k, pr), _geo_cdf(k, pr), REL9, "rel")
            row("Geometric_Survival", "K_STATS_Geometric_Survival", (k, pr), _geo_sf(k, pr), REL9, "rel")
        for prob in [mp.mpf("0.05"), mp.mpf("0.5"), mp.mpf("0.975")]:
            row("Geometric_InverseCumulative", "K_STATS_Geometric_InverseCumulative", (prob, pr), _geo_inv(prob, pr), ABS9, "abs")
        row("Geometric_Mean", "K_STATS_Geometric_Mean", (pr,), (1 - pr) / pr, REL14, "rel")
        row("Geometric_Variance", "K_STATS_Geometric_Variance", (pr,), (1 - pr) / pr ** 2, REL14, "rel")
        row("Geometric_StdDev", "K_STATS_Geometric_StdDev", (pr,), mp.sqrt(1 - pr) / pr, REL14, "rel")

    # --- Negative Binomial (failures before r-th success; args k, r, p) ---
    for r in [mp.mpf(1), mp.mpf(5), mp.mpf(50), mp.mpf(500), mp.mpf(5000)]:
        for pr in [mp.mpf("0.2"), mp.mpf("0.5"), mp.mpf("0.85")]:
            mean = r * (1 - pr) / pr
            sd = mp.sqrt(r * (1 - pr)) / pr
            for k in sorted(set([mp.mpf(mp.floor(mean)), mp.mpf(mp.floor(mean + 3 * sd))])):
                row("NegativeBinomial_PMF", "K_STATS_NegativeBinomial_PMF", (k, r, pr), _nb_pmf(k, r, pr), REL12, "rel")
                row("NegativeBinomial_LogPMF", "K_STATS_NegativeBinomial_LogPMF", (k, r, pr), _nb_logpmf(k, r, pr), REL12, "rel")
                row("NegativeBinomial_Cumulative", "K_STATS_NegativeBinomial_Cumulative", (k, r, pr), _nb_cdf(k, r, pr), REL9, "rel")
                row("NegativeBinomial_Survival", "K_STATS_NegativeBinomial_Survival", (k, r, pr), _nb_sf(k, r, pr), REL9, "rel")
            for prob in [mp.mpf("0.07"), mp.mpf("0.53"), mp.mpf("0.94")]:
                row("NegativeBinomial_InverseCumulative", "K_STATS_NegativeBinomial_InverseCumulative", (prob, r, pr), _nb_inv(prob, r, pr), ABS9, "abs")
            row("NegativeBinomial_Mean", "K_STATS_NegativeBinomial_Mean", (r, pr), r * (1 - pr) / pr, REL14, "rel")
            row("NegativeBinomial_Variance", "K_STATS_NegativeBinomial_Variance", (r, pr), r * (1 - pr) / pr ** 2, REL14, "rel")
            row("NegativeBinomial_StdDev", "K_STATS_NegativeBinomial_StdDev", (r, pr), mp.sqrt(r * (1 - pr)) / pr, REL14, "rel")

    # --- Hypergeometric (args k, n, K, N: sample succ, sample size, pop succ, pop size) ---
    for n, K, N in [(10, 20, 50), (30, 40, 100), (100, 500, 1000), (50, 200, 1000), (500, 5000, 100000)]:
        n, K, N = mp.mpf(n), mp.mpf(K), mp.mpf(N)
        lo = max(0, int(n) + int(K) - int(N)); hi = min(int(n), int(K))
        mode = int(mp.floor((n + 1) * (K + 1) / (N + 2)))
        sd = mp.sqrt(n * (K / N) * ((N - K) / N) * ((N - n) / (N - 1)))
        # PMF/LogPMF near the mode where the mass is representable (support
        # extremes for large N underflow far below double precision, where a
        # relative-error contract is meaningless).
        kset = sorted(set(max(lo, min(hi, mode + d)) for d in (0, int(mp.ceil(2 * sd)))))
        for k in [mp.mpf(x) for x in kset]:
            row("Hypergeometric_PMF", "K_STATS_Hypergeometric_PMF", (k, n, K, N), _hy_pmf(k, n, K, N), REL12, "rel")
            row("Hypergeometric_LogPMF", "K_STATS_Hypergeometric_LogPMF", (k, n, K, N), _hy_logpmf(k, n, K, N), REL12, "rel")
            row("Hypergeometric_Cumulative", "K_STATS_Hypergeometric_Cumulative", (k, n, K, N), _hy_cdf(k, n, K, N), REL9, "rel")
            row("Hypergeometric_Survival", "K_STATS_Hypergeometric_Survival", (k, n, K, N), _hy_sf(k, n, K, N), REL9, "rel")
        for prob in [mp.mpf("0.13"), mp.mpf("0.57"), mp.mpf("0.91")]:
            row("Hypergeometric_InverseCumulative", "K_STATS_Hypergeometric_InverseCumulative", (prob, n, K, N), _hy_inv(prob, n, K, N), ABS9, "abs")
        row("Hypergeometric_Mean", "K_STATS_Hypergeometric_Mean", (n, K, N), n * K / N, REL14, "rel")
        row("Hypergeometric_Variance", "K_STATS_Hypergeometric_Variance", (n, K, N), n * (K / N) * ((N - K) / N) * ((N - n) / (N - 1)), REL14, "rel")
        row("Hypergeometric_StdDev", "K_STATS_Hypergeometric_StdDev", (n, K, N), mp.sqrt(n * (K / N) * ((N - K) / N) * ((N - n) / (N - 1))), REL14, "rel")

    # --- Discrete Uniform (inclusive integer support [Lower, Upper]) ---
    # Supports deliberately include a negative range (exercises floor vs
    # truncate-toward-zero), the degenerate single-point support, and large
    # supports with non-round cardinality.
    #
    # Two grid-design rules keep every row a meaningful accuracy claim:
    #   * inverse probabilities are built as (j + 0.37) / n, so prob * n never
    #     lands on an integer - a discrete quantile evaluated exactly on a CDF
    #     step is ill-conditioned and would not be a real claim;
    #   * rows whose reference is exactly zero are skipped for the relative
    #     metric (a relative error against zero is vacuous). The degenerate
    #     n = 1 support is still covered through its non-zero rows, and the
    #     zero-valued cases are asserted exactly in the VBA test suite.
    for lo, hi in [(1, 6), (0, 9), (-5, 6), (1, 1), (-999983, 1000000), (0, 999983)]:
        lo_m, hi_m = mp.mpf(lo), mp.mpf(hi)
        n = _du_n(lo_m, hi_m)

        def du_row(fn, args, ref, claim, metric):
            if metric == "rel" and ref == 0:
                return
            row(fn, "K_STATS_" + fn, args, ref, claim, metric)

        for x in [mp.mpf(v) for v in sorted(set([lo, (lo + hi) // 2, hi]))]:
            du_row("DiscreteUniform_PMF", (x, lo_m, hi_m), _du_pmf(x, lo_m, hi_m), REL14, "rel")
            du_row("DiscreteUniform_LogPMF", (x, lo_m, hi_m), _du_logpmf(lo_m, hi_m), REL14, "rel")
            du_row("DiscreteUniform_Cumulative", (x, lo_m, hi_m), _du_cdf(x, lo_m, hi_m), REL14, "rel")
            if x < hi_m:
                du_row("DiscreteUniform_Survival", (x, lo_m, hi_m), _du_sf(x, lo_m, hi_m), REL14, "rel")

        for j in sorted(set([0, int(n) // 3, int(n) - 1])):
            prob = (mp.mpf(j) + mp.mpf("0.37")) / n
            du_row("DiscreteUniform_InverseCumulative", (prob, lo_m, hi_m), _du_inv(prob, lo_m, hi_m), ABS9, "abs")

        du_row("DiscreteUniform_Mean", (lo_m, hi_m), _du_mean(lo_m, hi_m), REL14, "rel")
        du_row("DiscreteUniform_Variance", (lo_m, hi_m), _du_var(lo_m, hi_m), REL14, "rel")
        du_row("DiscreteUniform_StdDev", (lo_m, hi_m), mp.sqrt(_du_var(lo_m, hi_m)), REL14, "rel")

    return rows



# ===========================================================================
# DEEP-TAIL INVERSE NORMAL (regime "deep_tail")
#   The central inverse contracts are tight but domain-restricted: beyond the
#   z = PROB_CDF_SPLIT (7.07) branch the achievable RELATIVE accuracy is bounded
#   by the relative accuracy of the normal CDF in its tail. These rows measure
#   what the kernels actually deliver where the README advertises them
#   (K_STATS_NormalStandard_InverseSurvival(1E-18) and beyond), so the claim is
#   evidence rather than a disclaimer.
# ===========================================================================
def _dt_Q(z):
    return mp.mpf("0.5") * mp.erfc(mp.mpf(z) / mp.sqrt(2))


def _dt_inv_surv(q):
    """z such that Q(z) = q.

    Newton in log space at extra working precision, seeded with the standard
    tail asymptote. A plain root-find loses conditioning once q is tiny, so the
    result is round-trip verified before it is allowed into the grid.
    """
    q = mp.mpf(q)
    with mp.workdps(mp.mp.dps + 30):
        t = mp.sqrt(-2 * mp.log(q))
        z = t - (mp.log(t) + mp.log(2 * mp.pi) / 2) / t
        if z <= 0:
            z = mp.mpf("0.5")
        for _ in range(200):
            step = (mp.log(_dt_Q(z)) - mp.log(q)) / (-mp.npdf(z) / _dt_Q(z))
            z = z - step
            if abs(step) < mp.mpf(10) ** (-(mp.mp.dps - 5)):
                break
        rel = abs(_dt_Q(z) - q) / q
    # Self-check: a reference that does not round-trip must never reach the grid.
    if rel > mp.mpf("1e-40"):
        raise SystemExit(f"deep-tail reference failed round-trip at q={q}: rel={rel}")
    return +z


def build_deep_tail_rows():
    rows = []
    REL = "measured"

    def row(func, args, ref):
        rows.append({
            "function": func, "vba_kernel": "K_STATS_" + func,
            "claim": REL, "metric": "rel",
            "arg1": mp.nstr(args[0], 17) if len(args) > 0 else "",
            "arg2": mp.nstr(args[1], 17) if len(args) > 1 else "",
            "arg3": mp.nstr(args[2], 17) if len(args) > 2 else "",
            "arg4": "",
            "reference": mp.nstr(ref, 25), "observed_vba": "",
            "regime": "deep_tail", "evidence_set": "main grid",
        })

    probs = ["1e-12", "1e-15", "1e-18", "1e-30", "1e-50", "1e-100", "1e-200", "1e-300"]
    for qs in probs:
        q = mp.mpf(qs)
        z = _dt_inv_surv(q)
        row("NormalStandard_InverseSurvival", (q,), z)
        # left tail: Phi(z) = p  =>  z = -InverseSurvival(p)
        row("NormalStandard_InverseCumulative", (q,), -z)
        for mean, sd in [(mp.mpf(100), mp.mpf(15)), (mp.mpf("-2.5"), mp.mpf("0.75"))]:
            row("Normal_InverseSurvival", (q, mean, sd), mean + sd * z)
        for ml, sl in [(mp.mpf(0), mp.mpf(1)), (mp.mpf(1), mp.mpf("0.5"))]:
            row("Lognormal_InverseSurvival", (q, ml, sl), mp.e ** (ml + sl * z))

    return rows


def build_rows():
    rows = []

    def add(func, vba_kernel, args, ref, claim=None, metric=None):
        # Claim and metric come from the single source of truth,
        # benchmark/accuracy_contracts.csv, so grid, summary, and README cannot drift.
        regime = _regime_for(func)
        if claim is None:
            contract = _CONTRACTS.get((func, regime)) or _CONTRACTS.get((func, "all"))
            claim = contract["claim"]
            metric = contract["metric"]
        rows.append(
            {
                "function": func,
                "vba_kernel": vba_kernel,
                "claim": claim,
                "metric": metric,
                "arg1": mp.nstr(args[0], 17) if len(args) > 0 else "",
                "arg2": mp.nstr(args[1], 17) if len(args) > 1 else "",
                "arg3": mp.nstr(args[2], 17) if len(args) > 2 else "",
                "arg4": mp.nstr(args[3], 17) if len(args) > 3 else "",
                "reference": mp.nstr(ref, 25),
                "observed_vba": "",
                "regime": regime,
                "evidence_set": "main grid",
            }
        )

    # --- PROB_LogGamma : Z in [1E-8, 1E+50], rel < 6.1E-14 ---
    for z in logspace("1e-8", "1e50", 40):
        add("LogGamma", "PROB_LogGamma", (z,), _loggamma(z))

    # --- PROB_LogGammaHalfDiff : Z > 0, rel <= 2E-14 (tested range) ---
    for z in logspace("1e-6", "1e12", 30):
        add("LogGammaHalfDiff", "PROB_LogGammaHalfDiff", (z,), _loggamma_halfdiff(z))

    # --- PROB_StirlingError : N >= 0.5, abs <= 3E-17 (include N=501 hot spot) ---
    ns = [mp.mpf("0.5"), mp.mpf(1), mp.mpf(2), mp.mpf(3), mp.mpf(5), mp.mpf(10),
          mp.mpf(50), mp.mpf(100), mp.mpf(500), mp.mpf(501), mp.mpf(1000), mp.mpf(1e6)]
    for n in ns:
        add("StirlingError", "PROB_StirlingError", (n,), _stirling_error(n))

    # --- PROB_LogChoose : N in [2, 2^53], all K, rel <= 3.2E-16 ---
    for n in [mp.mpf(2), mp.mpf(10), mp.mpf(100), mp.mpf(1030), mp.mpf(1e6), mp.mpf(2) ** 53]:
        for frac in [mp.mpf("0.0"), mp.mpf("0.01"), mp.mpf("0.5"), mp.mpf("0.99"), mp.mpf("1.0")]:
            k = mp.floor(n * frac)
            add("LogChoose", "PROB_LogChoose", (n, k), _logchoose(n, k))

    # --- Student t ---
    for df in [mp.mpf(1), mp.mpf(2), mp.mpf(5), mp.mpf(30), mp.mpf(1000)]:
        for x in [mp.mpf("0.1"), mp.mpf(1), mp.mpf(2), mp.mpf(5), mp.mpf(20)]:
            add("StudentT_Density", "K_STATS_StudentT_Density", (x, df), _student_t_pdf(x, df))
            add("StudentT_Cumulative", "K_STATS_StudentT_Cumulative", (x, df), _student_t_cdf(x, df))
            add("StudentT_Survival", "K_STATS_StudentT_Survival", (x, df), _student_t_sf(x, df))
        for p in [mp.mpf("0.001"), mp.mpf("0.05"), mp.mpf("0.5"), mp.mpf("0.95"), mp.mpf("0.999")]:
            add("StudentT_InverseCumulative", "K_STATS_StudentT_InverseCumulative", (p, df), _student_t_ppf(p, df))

    # --- Chi-square ---
    for df in [mp.mpf(1), mp.mpf(2), mp.mpf(5), mp.mpf(30), mp.mpf(100)]:
        for x in [mp.mpf("0.5"), mp.mpf(1), mp.mpf(5), mp.mpf(20), mp.mpf(80)]:
            add("ChiSquare_Cumulative", "K_STATS_ChiSquare_Cumulative", (x, df), _chi2_cdf(x, df))
            add("ChiSquare_Survival", "K_STATS_ChiSquare_Survival", (x, df), _chi2_sf(x, df))
        for p in [mp.mpf("0.001"), mp.mpf("0.05"), mp.mpf("0.5"), mp.mpf("0.95"), mp.mpf("0.999")]:
            add("ChiSquare_InverseCumulative", "K_STATS_ChiSquare_InverseCumulative", (p, df), _chi2_ppf(p, df))

    # --- F ---
    for d1, d2 in [(mp.mpf(1), mp.mpf(1)), (mp.mpf(5), mp.mpf(2)), (mp.mpf(10), mp.mpf(30)), (mp.mpf(100), mp.mpf(100))]:
        for x in [mp.mpf("0.25"), mp.mpf(1), mp.mpf(2), mp.mpf(10)]:
            add("F_Cumulative", "K_STATS_F_Cumulative", (x, d1, d2), _f_cdf(x, d1, d2))
            add("F_Survival", "K_STATS_F_Survival", (x, d1, d2), _f_sf(x, d1, d2))
        for p in [mp.mpf("0.05"), mp.mpf("0.5"), mp.mpf("0.95")]:
            add("F_InverseCumulative", "K_STATS_F_InverseCumulative", (p, d1, d2), _f_ppf(p, d1, d2))


    # ===================== NORMAL FAMILY =====================
    FIVE_E15 = "rel<=5E-15"

    # --- Standard Normal ---
    for z in [mp.mpf("-2"), mp.mpf("-0.5"), mp.mpf("0.5"), mp.mpf(1), mp.mpf("1.96"), mp.mpf(3)]:
        add("NormalStandard_Density", "K_STATS_NormalStandard_Density", (z,), _phi(z))
        add("NormalStandard_Cumulative", "K_STATS_NormalStandard_Cumulative", (z,), _Phi(z))
        add("NormalStandard_Survival", "K_STATS_NormalStandard_Survival", (z,), _Phi_sf(z))
    for pq in [mp.mpf("0.01"), mp.mpf("0.25"), mp.mpf("0.5"), mp.mpf("0.975"), mp.mpf("0.999")]:
        add("NormalStandard_InverseCumulative", "K_STATS_NormalStandard_InverseCumulative", (pq,), _Phi_inv(pq))
        add("NormalStandard_InverseSurvival", "K_STATS_NormalStandard_InverseSurvival", (pq,), -_Phi_inv(pq))
        add("NormalStandard_InverseCumulativeFast", "K_STATS_NormalStandard_InverseCumulativeFast", (pq,), _Phi_inv(pq))
    for lo, up in [(mp.mpf("-1.96"), mp.mpf("1.96")), (mp.mpf("-1"), mp.mpf(2)), (mp.mpf("0"), mp.mpf(3))]:
        add("NormalStandard_IntervalProbability", "K_STATS_NormalStandard_IntervalProbability", (lo, up), _Phi(up) - _Phi(lo))

    # --- General Normal ---
    for (x, m, sd) in [(mp.mpf("1.96"), mp.mpf(0), mp.mpf(1)), (mp.mpf(110), mp.mpf(100), mp.mpf(15)),
                       (mp.mpf(3), mp.mpf(5), mp.mpf(2))]:
        z = (x - m) / sd
        add("Normal_Density", "K_STATS_Normal_Density", (x, m, sd), _phi(z) / sd)
        add("Normal_Cumulative", "K_STATS_Normal_Cumulative", (x, m, sd), _Phi(z))
        add("Normal_Survival", "K_STATS_Normal_Survival", (x, m, sd), _Phi_sf(z))
        add("Normal_ZScore", "K_STATS_Normal_ZScore", (x, m, sd), z)
    for (pq, m, sd) in [(mp.mpf("0.99"), mp.mpf(100), mp.mpf(15)), (mp.mpf("0.025"), mp.mpf(10), mp.mpf(2))]:
        add("Normal_InverseCumulative", "K_STATS_Normal_InverseCumulative", (pq, m, sd), m + sd * _Phi_inv(pq))
        add("Normal_InverseSurvival", "K_STATS_Normal_InverseSurvival", (pq, m, sd), m - sd * _Phi_inv(pq))

    # --- Lognormal ---
    for (x, ml, sl) in [(mp.mpf(1), mp.mpf(0), mp.mpf(1)), (mp.mpf(2), mp.mpf("0.5"), mp.mpf("0.25")),
                        (mp.mpf("0.5"), mp.mpf(0), mp.mpf(1))]:
        zz = (mp.log(x) - ml) / sl
        add("Lognormal_Density", "K_STATS_Lognormal_Density", (x, ml, sl), _phi(zz) / (x * sl))
        add("Lognormal_Cumulative", "K_STATS_Lognormal_Cumulative", (x, ml, sl), _Phi(zz))
        add("Lognormal_Survival", "K_STATS_Lognormal_Survival", (x, ml, sl), _Phi_sf(zz))
    # Deep-underflow regression for the log-domain density reconstruction: |Z| = 40,
    # so the naive phi(Z)/(X*sl) underflows its numerator to zero, yet the density
    # (~1.6E-305) is representable. MeanLog = 0 keeps Z clean, isolating the density
    # path from input conditioning. This row fails on the old form, passes on the fix.
    for (x, ml, sl) in [(mp.e ** -100, mp.mpf(0), mp.mpf("2.5"))]:
        zz = (mp.log(x) - ml) / sl
        add("Lognormal_Density", "K_STATS_Lognormal_Density", (x, ml, sl), _phi(zz) / (x * sl))
    for (pq, ml, sl) in [(mp.mpf("0.5"), mp.mpf(0), mp.mpf(1)), (mp.mpf("0.025"), mp.mpf(0), mp.mpf(1))]:
        add("Lognormal_InverseCumulative", "K_STATS_Lognormal_InverseCumulative", (pq, ml, sl), mp.e ** (ml + sl * _Phi_inv(pq)))
        add("Lognormal_InverseSurvival", "K_STATS_Lognormal_InverseSurvival", (pq, ml, sl), mp.e ** (ml - sl * _Phi_inv(pq)))
    for (ml, sl) in [(mp.mpf(0), mp.mpf(1)), (mp.mpf("0.5"), mp.mpf("0.25"))]:
        add("Lognormal_Mean", "K_STATS_Lognormal_Mean", (ml, sl), mp.e ** (ml + sl * sl / 2))
        add("Lognormal_Variance", "K_STATS_Lognormal_Variance", (ml, sl), (mp.e ** (sl * sl) - 1) * mp.e ** (2 * ml + sl * sl))
        add("Lognormal_StdDev", "K_STATS_Lognormal_StdDev", (ml, sl), mp.sqrt((mp.e ** (sl * sl) - 1) * mp.e ** (2 * ml + sl * sl)))
    # ParametersFromMeanStdDev returns a 1x2 array; test each output separately
    for (mean, sd) in [(mp.mpf(2), mp.mpf("0.5")), (mp.mpf(10), mp.mpf(3))]:
        mlref, slref = _lognorm_params(mean, sd)


    # ===================== CONTINUOUS FAMILY =====================
    # Bounds set from measured worst-case error over the tested grid (5E-15 for
    # near-machine-epsilon functions, 2E-14 where a digit is lost). Exponential is
    # parameterized by RATE (Lambda), not scale.
    PROV = "rel<=1E-8"

    # --- Gamma(X, Shape k, ScaleParam theta) ---
    for (x, k, th) in [(mp.mpf(2), mp.mpf(2), mp.mpf(1)), (mp.mpf(5), mp.mpf(3), mp.mpf(2)),
                       (mp.mpf("0.5"), mp.mpf("1.5"), mp.mpf(1))]:
        add("Gamma_Density", "K_STATS_Gamma_Density", (x, k, th), _gamma_pdf(x, k, th))
        add("Gamma_Cumulative", "K_STATS_Gamma_Cumulative", (x, k, th), _gamma_cdf(x, k, th))
        add("Gamma_Survival", "K_STATS_Gamma_Survival", (x, k, th), _gamma_sf(x, k, th))
    for (pq, k, th) in [(mp.mpf("0.5"), mp.mpf(2), mp.mpf(1)), (mp.mpf("0.95"), mp.mpf(3), mp.mpf(2))]:
        add("Gamma_InverseCumulative", "K_STATS_Gamma_InverseCumulative", (pq, k, th), _gamma_ppf(pq, k, th))
    for (k, th) in [(mp.mpf(2), mp.mpf(3)), (mp.mpf(5), mp.mpf(2))]:
        add("Gamma_Mean", "K_STATS_Gamma_Mean", (k, th), mp.mpf(k) * mp.mpf(th))
        add("Gamma_Variance", "K_STATS_Gamma_Variance", (k, th), mp.mpf(k) * mp.mpf(th) ** 2)
        add("Gamma_StdDev", "K_STATS_Gamma_StdDev", (k, th), mp.sqrt(mp.mpf(k)) * mp.mpf(th))

    # --- Beta(X, Alpha a, Beta b) ---
    for (x, a, b) in [(mp.mpf("0.5"), mp.mpf(2), mp.mpf(2)), (mp.mpf("0.3"), mp.mpf(2), mp.mpf(5)),
                      (mp.mpf("0.8"), mp.mpf(5), mp.mpf(1))]:
        add("Beta_Density", "K_STATS_Beta_Density", (x, a, b), _beta_pdf(x, a, b))
        add("Beta_Cumulative", "K_STATS_Beta_Cumulative", (x, a, b), _beta_cdf(x, a, b))
        add("Beta_Survival", "K_STATS_Beta_Survival", (x, a, b), _beta_sf(x, a, b))
    for (pq, a, b) in [(mp.mpf("0.5"), mp.mpf(2), mp.mpf(2)), (mp.mpf("0.95"), mp.mpf(2), mp.mpf(5))]:
        add("Beta_InverseCumulative", "K_STATS_Beta_InverseCumulative", (pq, a, b), _beta_ppf(pq, a, b))
    for (a, b) in [(mp.mpf(2), mp.mpf(3)), (mp.mpf(5), mp.mpf(2))]:
        add("Beta_Mean", "K_STATS_Beta_Mean", (a, b), mp.mpf(a) / (mp.mpf(a) + mp.mpf(b)))
        add("Beta_Variance", "K_STATS_Beta_Variance", (a, b), mp.mpf(a) * mp.mpf(b) / ((mp.mpf(a) + mp.mpf(b)) ** 2 * (mp.mpf(a) + mp.mpf(b) + 1)))
        add("Beta_StdDev", "K_STATS_Beta_StdDev", (a, b), mp.sqrt(mp.mpf(a) * mp.mpf(b) / ((mp.mpf(a) + mp.mpf(b)) ** 2 * (mp.mpf(a) + mp.mpf(b) + 1))))

    # --- Exponential(X, Lambda=rate) ---
    for (x, lam) in [(mp.mpf(1), mp.mpf(1)), (mp.mpf("0.5"), mp.mpf(2)), (mp.mpf(3), mp.mpf("0.5"))]:
        lam = mp.mpf(lam)
        add("Exponential_Density", "K_STATS_Exponential_Density", (x, lam), lam * mp.e ** (-lam * mp.mpf(x)))
        add("Exponential_Cumulative", "K_STATS_Exponential_Cumulative", (x, lam), 1 - mp.e ** (-lam * mp.mpf(x)))
        add("Exponential_Survival", "K_STATS_Exponential_Survival", (x, lam), mp.e ** (-lam * mp.mpf(x)))
    for (pq, lam) in [(mp.mpf("0.5"), mp.mpf(1)), (mp.mpf("0.95"), mp.mpf(2))]:
        add("Exponential_InverseCumulative", "K_STATS_Exponential_InverseCumulative", (pq, lam), -mp.log(1 - mp.mpf(pq)) / mp.mpf(lam))

    # --- Weibull(X, Shape k, ScaleParam lam) ---
    for (x, k, lam) in [(mp.mpf(1), mp.mpf("1.5"), mp.mpf(1)), (mp.mpf(2), mp.mpf(2), mp.mpf(2)),
                        (mp.mpf("0.5"), mp.mpf(3), mp.mpf(1))]:
        k2, lam2 = mp.mpf(k), mp.mpf(lam)
        add("Weibull_Density", "K_STATS_Weibull_Density", (x, k, lam), (k2 / lam2) * (mp.mpf(x) / lam2) ** (k2 - 1) * mp.e ** (-(mp.mpf(x) / lam2) ** k2))
        add("Weibull_Cumulative", "K_STATS_Weibull_Cumulative", (x, k, lam), 1 - mp.e ** (-(mp.mpf(x) / lam2) ** k2))
        add("Weibull_Survival", "K_STATS_Weibull_Survival", (x, k, lam), mp.e ** (-(mp.mpf(x) / lam2) ** k2))
    for (pq, k, lam) in [(mp.mpf("0.5"), mp.mpf("1.5"), mp.mpf(1)), (mp.mpf("0.95"), mp.mpf(2), mp.mpf(2))]:
        add("Weibull_InverseCumulative", "K_STATS_Weibull_InverseCumulative", (pq, k, lam), mp.mpf(lam) * (-mp.log(1 - mp.mpf(pq))) ** (1 / mp.mpf(k)))
    for (k, lam) in [(mp.mpf("1.5"), mp.mpf(1)), (mp.mpf(2), mp.mpf(2))]:
        add("Weibull_Mean", "K_STATS_Weibull_Mean", (k, lam), _weibull_mean(k, lam))
        add("Weibull_Variance", "K_STATS_Weibull_Variance", (k, lam), _weibull_var(k, lam))
        add("Weibull_StdDev", "K_STATS_Weibull_StdDev", (k, lam), mp.sqrt(_weibull_var(k, lam)))

    # --- Uniform(X, LowerBound a, UpperBound b) ---
    for (x, a, b) in [(mp.mpf(3), mp.mpf(0), mp.mpf(10)), (mp.mpf("2.5"), mp.mpf(1), mp.mpf(4))]:
        a2, b2 = mp.mpf(a), mp.mpf(b)
        add("Uniform_Density", "K_STATS_Uniform_Density", (x, a, b), 1 / (b2 - a2))
        add("Uniform_Cumulative", "K_STATS_Uniform_Cumulative", (x, a, b), (mp.mpf(x) - a2) / (b2 - a2))
        add("Uniform_Survival", "K_STATS_Uniform_Survival", (x, a, b), (b2 - mp.mpf(x)) / (b2 - a2))
    for (pq, a, b) in [(mp.mpf("0.5"), mp.mpf(0), mp.mpf(10)), (mp.mpf("0.9"), mp.mpf(1), mp.mpf(4))]:
        add("Uniform_InverseCumulative", "K_STATS_Uniform_InverseCumulative", (pq, a, b), mp.mpf(a) + mp.mpf(pq) * (mp.mpf(b) - mp.mpf(a)))

    rows += build_discrete_rows()
    rows += build_deep_tail_rows()
    return rows


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--digits", type=int, default=50)
    ap.add_argument("--out", default="probability_accuracy_grid.csv")
    args = ap.parse_args()

    mp.mp.dps = args.digits
    rows = build_rows()

    fields = ["function", "vba_kernel", "claim", "metric",
              "arg1", "arg2", "arg3", "arg4", "reference", "observed_vba",
              "regime", "evidence_set"]
    with open(args.out, "w", newline="") as f:
        w = csv.DictWriter(f, fieldnames=fields)
        w.writeheader()
        w.writerows(rows)

    print(f"wrote {args.out}: {len(rows)} reference rows at {args.digits} digits")
    print(f"generated {_dt.date.today().isoformat()}")


if __name__ == "__main__":
    main()
