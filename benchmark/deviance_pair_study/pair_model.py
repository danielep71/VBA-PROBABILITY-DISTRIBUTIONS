"""
#498: the Loader deviance as a pair of Doubles, and the mass functions and
densities that carry the pair to the final exponential.

Today PROB_TryDeviancePart returns bd0(X, Mean) = X Log(X / Mean) + Mean - X
as one Double. Away from the mean it is a difference of two larger numbers,
and even a correctly rounded Double of size D is off by D * 2^-53, which is
the relative error of the density.

The candidate forms the deviance as a pair (high, low) whose sum is exact to
far below one unit of the high part:

  - Log(X / Mean) = k Log 2 + 2 atanh(w), w = (s - Mean) / (s + Mean),
    s = X / 2^k, every step in pairs;
  - bd0 = X * Log(X / Mean) + (Mean - X), in pairs;
  - the mean itself is a pair where the caller forms it by a rounded
    operation (N * p and N * (1 - p) in the Binomial).

The caller adds the other terms of the logarithm to the pair and returns
Exp(high) * (1 + low).

This file is a model in plain binary64, the arithmetic the error-free
transforms need. VBA keeps 64-bit intermediates inside an expression, which
breaks them unless every operation is stored in a Double; the pair toolkit of
#35 (PROB_BP_Sum, PROB_BP_Add, PROB_BP_Mul, PROB_BP_Div) already does that.
"today" is the port of the source in ../beta_balanced_survival/beta_route_model.py.
"""
import math
import os
import sys

import mpmath as mp

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(os.path.dirname(HERE), "beta_balanced_survival"))
import beta_route_model as M  # noqa: E402

SPLITTER = 134217729.0                                   # 2^27 + 1
LOG_TWO = (0.693147180559945 + 3.33066907387547e-16, 2.3190468138463e-17)   # the source's literals
ATANH_MAX_TERMS = 80
ATANH_STOP = 1e-34
HALF_LOG_TWO_PI = M.C["PROB_HALF_LOG_TWO_PI"]


def two_sum(a, b):
    s = a + b
    bb = s - a
    return s, (a - (s - bb)) + (b - bb)


def _split(a):
    t = SPLITTER * a
    h = t - (t - a)
    return h, a - h


def two_prod(a, b):
    p = a * b
    ah, al = _split(a)
    bh, bl = _split(b)
    return p, ((ah * bh - p) + ah * bl + al * bh) + al * bl


def add(a, b):
    s, e = two_sum(a[0], b[0])
    return two_sum(s, e + (a[1] + b[1]))


def mul(a, b):
    p, e = two_prod(a[0], b[0])
    return two_sum(p, e + (a[0] * b[1] + a[1] * b[0]))


def div(a, b):
    q = a[0] / b[0]
    p = mul((q, 0.0), b)
    r = add(a, (-p[0], -p[1]))
    return two_sum(q, r[0] / b[0])


def neg(a):
    return -a[0], -a[1]


def log_ratio(x, mean):
    """Log(x / mean) as a pair; x a Double, mean a pair, both positive."""
    k = round(math.log2(x / mean[0]))
    s = math.ldexp(x, -k)                                # exact
    w = div(add((s, 0.0), neg(mean)), add((s, 0.0), mean))
    w2 = mul(w, w)
    total = term = w
    for j in range(1, ATANH_MAX_TERMS + 1):
        term = mul(term, w2)
        piece = div(term, (2.0 * j + 1.0, 0.0))
        total = add(total, piece)
        if abs(piece[0]) <= ATANH_STOP * abs(total[0]):     # also x equal to the mean: w is zero
            break
    else:
        raise ArithmeticError("atanh series did not converge")
    return add((2.0 * total[0], 2.0 * total[1]), mul((float(k), 0.0), LOG_TWO))


def deviance(x, mean):
    """bd0(x, mean) as a pair; mean a Double or a pair."""
    if not isinstance(mean, tuple):
        mean = (mean, 0.0)
    return add(mul((x, 0.0), log_ratio(x, mean)), add(mean, (-x, 0.0)))


def _finish(terms):
    """Exp of a sum of Doubles and pairs: Exp(high) * (1 + low), as PROB_TryExpPair."""
    total = (0.0, 0.0)
    for t in terms:
        total = add(total, t if isinstance(t, tuple) else (t, 0.0))
    return math.exp(total[0]) * (1.0 + total[1])


def poisson_pmf(k, mean, rule=True):
    if rule:
        return _finish([neg(deviance(k, mean)), -M.stirling_error(k), -HALF_LOG_TWO_PI,
                        -0.5 * math.log(k)])
    return math.exp(-M.stirling_error(k) - M.deviance_part(k, mean) - HALF_LOG_TWO_PI
                    - 0.5 * math.log(k))


def binomial_pmf(k, n, p, rule=True):
    j = n - k
    corr = 0.5 * (math.log(k) + math.log(j) - math.log(n))
    if rule:
        q = two_sum(1.0, -p)
        return _finish([M.stirling_error(n), -M.stirling_error(k), -M.stirling_error(j),
                        neg(deviance(k, two_prod(n, p))), neg(deviance(j, mul((n, 0.0), q))),
                        -HALF_LOG_TWO_PI, -corr])
    return math.exp(M.stirling_error(n) - M.stirling_error(k) - M.stirling_error(j)
                    - M.deviance_part(k, n * p) - M.deviance_part(j, n * (1.0 - p))
                    - HALF_LOG_TWO_PI - corr)


def gamma_density(x, shape, scale=1.0, rule=True):
    """K_STATS_Gamma_Density on its direct branch; the chi-square density is shape df / 2, scale 2."""
    log_ratio_ = math.log(x) - math.log(scale)
    if rule:
        standard = div((x, 0.0), (scale, 0.0))
        return _finish([neg(deviance(shape, standard)), 0.5 * math.log(shape), -log_ratio_,
                        -HALF_LOG_TWO_PI, -M.stirling_error(shape), -math.log(scale)])
    return math.exp(-M.deviance_part(shape, x / scale) + 0.5 * math.log(shape) - log_ratio_
                    - HALF_LOG_TWO_PI - M.stirling_error(shape) - math.log(scale))


def chisquare_density(x, df, rule=True):
    return gamma_density(x, 0.5 * df, 2.0, rule=rule)


def reference(fn, *args, dps=60):
    with mp.workdps(dps):
        a = [mp.mpf(v) for v in args]
        if fn == "poisson":
            return mp.exp(a[0] * mp.log(a[1]) - a[1] - mp.loggamma(a[0] + 1))
        if fn == "binomial":
            return mp.binomial(a[1], a[0]) * a[2] ** a[0] * (1 - a[2]) ** (a[1] - a[0])
        if fn == "gamma":
            scale = a[2] if len(a) > 2 else mp.mpf(1)
            return mp.exp((a[1] - 1) * mp.log(a[0] / scale) - a[0] / scale - mp.loggamma(a[1])) / scale
        if fn == "deviance":
            return a[0] * mp.log(a[0] / a[1]) + a[1] - a[0]
    raise ValueError(fn)


def rel(value, ref):
    with mp.workdps(60):
        return float(abs((mp.mpf(value) - ref) / ref))
