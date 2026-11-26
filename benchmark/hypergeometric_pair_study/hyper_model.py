"""
#496: the Hypergeometric mass function with its three Binomial logarithms
carried as pairs of Doubles, on the pair arithmetic of #498
(../deviance_pair_study).

The library forms the mass as Loader does:

    dhyper(k; n, K, N) = dbinom(k; K, p) * dbinom(n - k; N - K, p) / dbinom(n; N, p),   p = n / N

Today each Binomial logarithm is one Double (PROB_DS_TryBinomialLogMass), the
three are added and subtracted, and p is rounded before any of them sees it.
The candidate takes p as a pair, each logarithm as a pair by the route of
#498, adds them as pairs and returns Exp(high) * (1 + low).

This file is a model in plain binary64. "today" stands on the ports of
../beta_balanced_survival/beta_route_model.py.
"""
import math
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(os.path.dirname(HERE), "deviance_pair_study"))
import pair_model as P  # noqa: E402

M = P.M
HALF_LOG_TWO_PI = P.HALF_LOG_TWO_PI


def binomial_log_pair(k, n, p):
    """The logarithm of dbinom(k; n, p) as a pair, for 0 < k < n and a pair p in (0, 1)."""
    j = n - k
    q = P.add((1.0, 0.0), P.neg(p))
    total = P.add(P.neg(P.deviance(k, P.mul((n, 0.0), p))),
                  P.neg(P.deviance(j, P.mul((n, 0.0), q))))
    for v in (M.stirling_error(n), -M.stirling_error(k), -M.stirling_error(j), -HALF_LOG_TWO_PI,
              -0.5 * (math.log(k) + math.log(j) - math.log(n))):
        total = P.add(total, (v, 0.0))
    return total


def _binomial_log_today(k, n, p):
    j = n - k
    return (M.stirling_error(n) - M.stirling_error(k) - M.stirling_error(j)
            - M.deviance_part(k, n * p) - M.deviance_part(j, n * (1.0 - p))
            - HALF_LOG_TWO_PI - 0.5 * (math.log(k) + math.log(j) - math.log(n)))


def interior(k, n, successes, population):
    """True where each of the three Binomial terms is away from its edges."""
    return (0 < k < successes and 0 < n - k < population - successes and 0 < n < population)


def hypergeometric_pmf(k, n, successes, population, rule=True):
    """K_STATS_Hypergeometric_PMF(k, n, K, N) where the three terms are interior."""
    if not interior(k, n, successes, population):
        raise ValueError("an edge of a Binomial term: not modelled")
    if rule:
        p = P.div((n, 0.0), (population, 0.0))
        total = P.add(P.add(binomial_log_pair(k, successes, p),
                            binomial_log_pair(n - k, population - successes, p)),
                      P.neg(binomial_log_pair(n, population, p)))
        return math.exp(total[0]) * (1.0 + total[1])
    p = n / population
    return math.exp(_binomial_log_today(k, successes, p)
                    + _binomial_log_today(n - k, population - successes, p)
                    - _binomial_log_today(n, population, p))
