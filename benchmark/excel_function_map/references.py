"""
#330: the value of a library expression from 60-digit arithmetic.

evaluate("K_STATS_Gamma_Cumulative(3,2,1.5)") is the mathematical value of
that call at the Doubles its arguments parse to. It is independent of the
VBA: every function here is written from its definition with mpmath.
Arithmetic around the calls (a factor, a difference, 1 - p) is evaluated as
Excel evaluates it, in binary64, on the arguments; the functions themselves
are evaluated at 60 digits.
"""
import mpmath as mp

mp.mp.dps = 60
F = mp.mpf
HALF = F(1) / 2


def _root(function, target, low, high):
    """The x in [low, high] with function(x) = target, function increasing."""
    low, high = F(low), F(high)
    for _ in range(400):
        middle = (low + high) / 2
        if function(middle) < target:
            low = middle
        else:
            high = middle
    return (low + high) / 2


def _upper(function, target):
    high = F(1)
    while function(high) < target:
        high *= 2
    return high


def norm_cdf(z):
    return mp.erfc(-F(z) / mp.sqrt(2)) / 2


def norm_sf(z):
    return mp.erfc(F(z) / mp.sqrt(2)) / 2


def norm_pdf(z):
    return mp.exp(-F(z) ** 2 / 2) / mp.sqrt(2 * mp.pi)


def norm_inv(p):
    return mp.sqrt(2) * mp.erfinv(2 * F(p) - 1)


def t_sf(x, df):
    x, df = F(x), F(df)
    tail = mp.betainc(df / 2, HALF, 0, df / (df + x * x), regularized=True) / 2
    return tail if x >= 0 else 1 - tail


def t_pdf(x, df):
    x, df = F(x), F(df)
    return mp.gamma((df + 1) / 2) / (mp.sqrt(df * mp.pi) * mp.gamma(df / 2)) \
        * (1 + x * x / df) ** (-(df + 1) / 2)


def gamma_cdf(x, shape, scale=1):
    return mp.gammainc(F(shape), 0, F(x) / F(scale), regularized=True)


def gamma_sf(x, shape, scale=1):
    return mp.gammainc(F(shape), F(x) / F(scale), mp.inf, regularized=True)


def gamma_pdf(x, shape, scale=1):
    x, shape, scale = F(x), F(shape), F(scale)
    return x ** (shape - 1) * mp.exp(-x / scale) / (mp.gamma(shape) * scale ** shape)


def beta_cdf(x, a, b):
    return mp.betainc(F(a), F(b), 0, F(x), regularized=True)


def beta_pdf(x, a, b):
    x, a, b = F(x), F(a), F(b)
    return x ** (a - 1) * (1 - x) ** (b - 1) / mp.beta(a, b)


def f_cdf(x, d1, d2):
    x, d1, d2 = F(x), F(d1), F(d2)
    return mp.betainc(d1 / 2, d2 / 2, 0, d1 * x / (d1 * x + d2), regularized=True)


def f_sf(x, d1, d2):
    x, d1, d2 = F(x), F(d1), F(d2)
    return mp.betainc(d2 / 2, d1 / 2, 0, d2 / (d1 * x + d2), regularized=True)


def f_pdf(x, d1, d2):
    x, d1, d2 = F(x), F(d1), F(d2)
    return mp.sqrt((d1 * x) ** d1 * d2 ** d2 / (d1 * x + d2) ** (d1 + d2)) / (x * mp.beta(d1 / 2, d2 / 2))


def binom_pmf(k, n, p):
    k, n, p = int(k), int(n), F(p)
    return mp.binomial(n, k) * p ** k * (1 - p) ** (n - k) if 0 <= k <= n else F(0)


def binom_cdf(k, n, p):
    return mp.fsum(binom_pmf(i, n, p) for i in range(0, int(mp.floor(k)) + 1))


def negbin_pmf(f, s, p):
    f, s, p = int(f), F(s), F(p)
    return mp.binomial(f + s - 1, f) * p ** s * (1 - p) ** f if f >= 0 else F(0)


def poisson_pmf(k, mean):
    k, mean = int(k), F(mean)
    return mp.exp(-mean) * mean ** k / mp.factorial(k) if k >= 0 else F(0)


def hyper_pmf(k, n, big_k, big_n):
    k, n, big_k, big_n = int(k), int(n), int(big_k), int(big_n)
    return mp.binomial(big_k, k) * mp.binomial(big_n - big_k, n - k) / mp.binomial(big_n, n)


def _smallest(cdf, p):
    k = 0
    while cdf(k) < F(p):
        k += 1
    return F(k)


NAMES = {
    "K_STATS_NormalStandard_Cumulative": norm_cdf,
    "K_STATS_NormalStandard_Density": norm_pdf,
    "K_STATS_NormalStandard_Survival": norm_sf,
    "K_STATS_NormalStandard_InverseCumulative": norm_inv,
    "K_STATS_NormalStandard_IntervalProbability": lambda a, b: norm_cdf(b) - norm_cdf(a),
    "K_STATS_Normal_Cumulative": lambda x, m, s: norm_cdf((F(x) - F(m)) / F(s)),
    "K_STATS_Normal_Density": lambda x, m, s: norm_pdf((F(x) - F(m)) / F(s)) / F(s),
    "K_STATS_Normal_InverseCumulative": lambda p, m, s: F(m) + F(s) * norm_inv(p),
    "K_STATS_Normal_ZScore": lambda x, m, s: (F(x) - F(m)) / F(s),
    "K_STATS_Lognormal_Cumulative": lambda x, m, s: norm_cdf((mp.log(F(x)) - F(m)) / F(s)),
    "K_STATS_Lognormal_Density":
        lambda x, m, s: norm_pdf((mp.log(F(x)) - F(m)) / F(s)) / (F(s) * F(x)),
    "K_STATS_Lognormal_InverseCumulative": lambda p, m, s: mp.exp(F(m) + F(s) * norm_inv(p)),
    "K_STATS_StudentT_Cumulative": lambda x, df: t_sf(-F(x), df),
    "K_STATS_StudentT_Density": t_pdf,
    "K_STATS_StudentT_Survival": t_sf,
    "K_STATS_StudentT_InverseCumulative":
        lambda p, df: _root(lambda x: t_sf(-x, df), F(p), -1000, 1000),
    "K_STATS_ChiSquare_Cumulative": lambda x, df: gamma_cdf(F(x) / 2, F(df) / 2),
    "K_STATS_ChiSquare_Density": lambda x, df: gamma_pdf(x, F(df) / 2, 2),
    "K_STATS_ChiSquare_Survival": lambda x, df: gamma_sf(F(x) / 2, F(df) / 2),
    "K_STATS_ChiSquare_InverseCumulative":
        lambda p, df: _root(lambda x: gamma_cdf(x / 2, F(df) / 2), F(p), 0,
                            _upper(lambda x: gamma_cdf(x / 2, F(df) / 2), F(p))),
    "K_STATS_F_Cumulative": f_cdf,
    "K_STATS_F_Density": f_pdf,
    "K_STATS_F_Survival": f_sf,
    "K_STATS_F_InverseCumulative":
        lambda p, d1, d2: _root(lambda x: f_cdf(x, d1, d2), F(p), 0,
                                _upper(lambda x: f_cdf(x, d1, d2), F(p))),
    "K_STATS_Gamma_Cumulative": gamma_cdf,
    "K_STATS_Gamma_Density": gamma_pdf,
    "K_STATS_Gamma_InverseCumulative":
        lambda p, a, b: _root(lambda x: gamma_cdf(x, a, b), F(p), 0,
                              _upper(lambda x: gamma_cdf(x, a, b), F(p))),
    "K_STATS_Beta_Cumulative": beta_cdf,
    "K_STATS_Beta_Density": beta_pdf,
    "K_STATS_Beta_InverseCumulative": lambda p, a, b: _root(lambda x: beta_cdf(x, a, b), F(p), 0, 1),
    "K_STATS_Exponential_Cumulative": lambda x, lam: -mp.expm1(-F(lam) * F(x)),
    "K_STATS_Exponential_Density": lambda x, lam: F(lam) * mp.exp(-F(lam) * F(x)),
    "K_STATS_Weibull_Cumulative": lambda x, k, s: -mp.expm1(-(F(x) / F(s)) ** F(k)),
    "K_STATS_Weibull_Density":
        lambda x, k, s: F(k) / F(s) * (F(x) / F(s)) ** (F(k) - 1) * mp.exp(-(F(x) / F(s)) ** F(k)),
    "K_STATS_Binomial_Cumulative": binom_cdf,
    "K_STATS_Binomial_PMF": binom_pmf,
    "K_STATS_Binomial_InverseCumulative": lambda p, n, q: _smallest(lambda k: binom_cdf(k, n, q), p),
    "K_STATS_NegativeBinomial_Cumulative":
        lambda f, s, p: mp.fsum(negbin_pmf(i, s, p) for i in range(0, int(mp.floor(f)) + 1)),
    "K_STATS_NegativeBinomial_PMF": negbin_pmf,
    "K_STATS_Poisson_Cumulative":
        lambda k, m: mp.fsum(poisson_pmf(i, m) for i in range(0, int(mp.floor(k)) + 1)),
    "K_STATS_Poisson_PMF": poisson_pmf,
    "K_STATS_Hypergeometric_Cumulative":
        lambda k, n, bk, bn: mp.fsum(hyper_pmf(i, n, bk, bn) for i in range(0, int(mp.floor(k)) + 1)),
    "K_STATS_Hypergeometric_PMF": hyper_pmf,
}


def evaluate(expression):
    """The value of a library expression; arithmetic outside the calls is
    binary64 on its literals, as Excel evaluates it."""
    return eval(expression, {"__builtins__": {}}, dict(NAMES))  # noqa: S307 - fixed, tracked input
