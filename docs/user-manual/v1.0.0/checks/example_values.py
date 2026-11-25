"""Reproduce every numeric result quoted in the user manual.

Each value is computed from the mathematical definition with mpmath at
30 significant digits. Nothing here calls the VBA library, so the values
are independent of it. Output: one line per example,
'example id, formula, value', with the value rounded to 10 significant
digits.

    python3 example_values.py
"""
import sys

from mpmath import (
    mp, mpf, sqrt, exp, log, erfc, gamma, gammainc, betainc, binomial,
    nstr, pi,
)

mp.dps = 30


# --- continuous building blocks -------------------------------------------
def norm_sf(z):
    return erfc(mpf(z) / sqrt(2)) / 2


def norm_cdf(z):
    return norm_sf(-mpf(z))


def norm_pdf(z):
    z = mpf(z)
    return exp(-z * z / 2) / sqrt(2 * pi)


def bisect(func, lo, hi, steps=220):
    """Root of an increasing function by plain bisection."""
    lo, hi = mpf(lo), mpf(hi)
    for _ in range(steps):
        mid = (lo + hi) / 2
        if func(mid) < 0:
            lo = mid
        else:
            hi = mid
    return (lo + hi) / 2


def norm_isf(q):
    """z with P(Z > z) = q, solved on the upper tail directly."""
    q = mpf(q)
    return bisect(lambda z: q - norm_sf(z), -60, 60)


def norm_inv(p):
    return -norm_isf(p)


def t_sf(x, df):
    x, df = mpf(x), mpf(df)
    tail = betainc(df / 2, mpf(1) / 2, 0, df / (df + x * x), regularized=True) / 2
    return tail if x >= 0 else 1 - tail


def t_inv(p, df):
    return bisect(lambda x: (1 - t_sf(x, df)) - mpf(p), -1000, 1000)


def gamma_cdf(x, shape, scale):
    return gammainc(mpf(shape), 0, mpf(x) / mpf(scale), regularized=True)


def gamma_sf(x, shape, scale):
    return gammainc(mpf(shape), mpf(x) / mpf(scale), mp.inf, regularized=True)


def gamma_inv(p, shape, scale):
    return bisect(lambda x: gamma_cdf(x, shape, scale) - mpf(p), 0, 10000)


def chi2_sf(x, df):
    return gamma_sf(x, mpf(df) / 2, 2)


def f_cdf(x, d1, d2):
    x, d1, d2 = mpf(x), mpf(d1), mpf(d2)
    return betainc(d1 / 2, d2 / 2, 0, d1 * x / (d1 * x + d2), regularized=True)


def f_inv(p, d1, d2):
    return bisect(lambda x: f_cdf(x, d1, d2) - mpf(p), 0, 100000)


# --- discrete building blocks ---------------------------------------------
def binom_pmf(k, n, p):
    p = mpf(p)
    return binomial(n, k) * p ** k * (1 - p) ** (n - k)


def binom_cdf(k, n, p):
    return sum(binom_pmf(i, n, p) for i in range(0, k + 1))


def binom_sf(k, n, p):
    return sum(binom_pmf(i, n, p) for i in range(k + 1, n + 1))


def pois_pmf(k, mean):
    mean = mpf(mean)
    return exp(-mean) * mean ** k / gamma(k + 1)


def pois_cdf(k, mean):
    return sum(pois_pmf(i, mean) for i in range(0, k + 1))


def pois_sf(k, mean):
    return gammainc(k + 1, 0, mpf(mean), regularized=True)


def least_k(cdf, p, start=0):
    """Smallest integer k with cdf(k) >= p."""
    k = start
    while cdf(k) < mpf(p):
        k += 1
    return k


def geom_cdf(k, p):
    return 1 - (1 - mpf(p)) ** (k + 1)


def nbinom_pmf(k, r, p):
    p = mpf(p)
    return binomial(k + r - 1, k) * p ** r * (1 - p) ** k


def nbinom_cdf(k, r, p):
    return sum(nbinom_pmf(i, r, p) for i in range(0, k + 1))


def hyper_pmf(k, n, K, N):
    return binomial(K, k) * binomial(N - K, n - k) / binomial(N, n)


# --- the examples ---------------------------------------------------------
def build():
    rows = []

    def add(example_id, formula, value):
        rows.append((example_id, formula, value))

    # Chapter 4: first formula and the five-minute check
    add("4.3-a", "K_STATS_NormalStandard_Cumulative(1.64485362695147)",
        norm_cdf("1.64485362695147"))
    add("4.4-a", "K_STATS_NormalStandard_Density(0)", norm_pdf(0))
    add("4.4-b", "K_STATS_NormalStandard_Survival(8)", norm_sf(8))
    add("4.4-c", "K_STATS_NormalStandard_InverseCumulative(0.975)",
        norm_inv("0.975"))
    add("4.4-d", "K_STATS_Gamma_Mean(3,2)", mpf(3) * 2)
    add("4.4-e", "K_STATS_Binomial_PMF(7,20,0.35)", binom_pmf(7, 20, "0.35"))
    add("4.4-f", "K_STATS_Poisson_InverseCumulative(0.99,25)",
        least_k(lambda k: pois_cdf(k, 25), "0.99"))

    # Chapter 5: the VBA example
    add("5.7-a", "K_STATS_Poisson_InverseCumulative(0.99,25)",
        least_k(lambda k: pois_cdf(k, 25), "0.99"))

    # Chapter 7: conventions
    add("7.3-a", "K_STATS_NormalStandard_Survival(15)", norm_sf(15))
    add("7.4-a", "K_STATS_NormalStandard_InverseSurvival(1E-18)",
        norm_isf("1e-18"))
    add("7.5-a", "K_STATS_Binomial_Cumulative(3,10,0.3)", binom_cdf(3, 10, "0.3"))
    add("7.5-b", "K_STATS_Binomial_Cumulative(4,10,0.3)", binom_cdf(4, 10, "0.3"))
    add("7.5-c", "K_STATS_Binomial_InverseCumulative(0.8,10,0.3)",
        least_k(lambda k: binom_cdf(k, 10, "0.3"), "0.8"))
    add("7.5-d", "K_STATS_Binomial_InverseCumulative(0.75,2,0.5)",
        least_k(lambda k: binom_cdf(k, 2, "0.5"), "0.75"))
    add("7.5-e", "K_STATS_Geometric_InverseCumulative(0.5000000000001,0.5)",
        least_k(lambda k: geom_cdf(k, "0.5"), "0.5000000000001"))
    add("7.6-a", "K_STATS_DiscreteUniform_Cumulative(2.8,-3,5)", mpf(6) / 9)

    # 8.1 A threshold and a tail from a Normal model
    add("8.1-a", "K_STATS_Normal_InverseCumulative(0.99,100,15)",
        100 + 15 * norm_inv("0.99"))
    add("8.1-b", "K_STATS_Normal_Survival(130,100,15)", norm_sf(2))
    add("8.1-c", "K_STATS_Normal_IntervalProbability(85,115,100,15)",
        norm_cdf(1) - norm_cdf(-1))

    # 8.2 A very small exceedance probability
    add("8.2-a", "K_STATS_NormalStandard_Survival(8)", norm_sf(8))
    add("8.2-b", "K_STATS_Normal_InverseSurvival(0.000001,0.0005,0.012)",
        mpf("0.0005") + mpf("0.012") * norm_isf("1e-6"))

    # 8.3 A two-sided t test
    add("8.3-a", "2*K_STATS_StudentT_Survival(2.5,12)", 2 * t_sf("2.5", 12))
    add("8.3-b", "K_STATS_StudentT_InverseCumulative(0.975,12)",
        t_inv("0.975", 12))

    # 8.4 Chi-square and F tests
    add("8.4-a", "K_STATS_ChiSquare_Survival(18.3,10)", chi2_sf("18.3", 10))
    add("8.4-b", "K_STATS_F_Survival(4.2,3,20)", 1 - f_cdf("4.2", 3, 20))
    add("8.4-c", "K_STATS_F_InverseCumulative(0.95,3,20)", f_inv("0.95", 3, 20))

    # 8.5 A Lognormal loss from its mean and standard deviation
    mean, sd = mpf(100), mpf(30)
    var_log = log(1 + (sd / mean) ** 2)
    mean_log = log(mean) - var_log / 2
    sd_log = sqrt(var_log)
    add("8.5-a", "INDEX(K_STATS_Lognormal_ParametersFromMeanStdDev(100,30),1,1)",
        mean_log)
    add("8.5-b", "INDEX(K_STATS_Lognormal_ParametersFromMeanStdDev(100,30),1,2)",
        sd_log)
    add("8.5-c", "K_STATS_Lognormal_Survival(200,B1,B2)",
        norm_sf((log(200) - mean_log) / sd_log))
    add("8.5-d", "K_STATS_Lognormal_InverseCumulative(0.995,B1,B2)",
        exp(mean_log + sd_log * norm_inv("0.995")))
    add("8.5-e", "K_STATS_Lognormal_Mean(B1,B2)",
        exp(mean_log + var_log / 2))

    # 8.6 Gamma claim severity and a Weibull lifetime
    add("8.6-a", "K_STATS_Gamma_Mean(3,2)", mpf(6))
    add("8.6-b", "K_STATS_Gamma_InverseCumulative(0.99,3,2)",
        gamma_inv("0.99", 3, 2))
    add("8.6-c", "K_STATS_Gamma_Survival(15,3,2)", gamma_sf(15, 3, 2))
    add("8.6-d", "K_STATS_Weibull_Survival(1000,1.5,2000)",
        exp(-(mpf(1000) / 2000) ** mpf("1.5")))

    # 8.7 Binomial default counts
    add("8.7-a", "K_STATS_Binomial_PMF(7,20,0.35)", binom_pmf(7, 20, "0.35"))
    add("8.7-b", "K_STATS_Binomial_Survival(14,20,0.35)",
        binom_sf(14, 20, "0.35"))
    add("8.7-c", "K_STATS_Binomial_InverseCumulative(0.99,100,0.02)",
        least_k(lambda k: binom_cdf(k, 100, "0.02"), "0.99"))
    add("8.7-d", "K_STATS_Binomial_LogPMF(7,20,0.35)",
        log(binom_pmf(7, 20, "0.35")))

    # 8.8 Event counts and sampling
    add("8.8-a", "K_STATS_Poisson_Cumulative(12,8.5)", pois_cdf(12, "8.5"))
    add("8.8-b", "K_STATS_Poisson_Survival(60,2)", pois_sf(60, 2))
    add("8.8-c", "K_STATS_Geometric_Cumulative(4,0.2)", geom_cdf(4, "0.2"))
    add("8.8-d", "K_STATS_NegativeBinomial_Cumulative(5,3,0.4)",
        nbinom_cdf(5, 3, "0.4"))
    add("8.8-e", "K_STATS_Hypergeometric_PMF(4,10,20,100)",
        hyper_pmf(4, 10, 20, 100))
    return rows


def main():
    rows = build()
    print("example id, formula, value")
    for example_id, formula, value in rows:
        text = str(value) if isinstance(value, int) else nstr(mpf(value), 10)
        print(f"{example_id}, {formula}, {text}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
