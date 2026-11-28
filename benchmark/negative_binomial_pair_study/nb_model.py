"""
#504: the Negative Binomial and Geometric mass functions with their large
terms carried as pairs of Doubles, on the pair arithmetic of #498.

Negative Binomial. The library forms the mass as Loader does,

    dnbinom(x; r, p) = dbinom(r; r + x, p) * r / (r + x),

with the Binomial logarithm in one Double (PROB_DS_TryBinomialLogMass). The
candidate takes that logarithm as a pair (PROB_TryBinomialLogPair, #496),
adds Log(r / (r + x)) as a pair, and returns Exp(high) * (1 + low).

Geometric. The library forms Log(p) + k * Log(1 - p), the product rounded at
its own size. The candidate forms 1 - p exactly, Log(1 - p) as a pair, the
product as a pair, adds Log(p), and returns Exp(high) * (1 + low).

This file is a model in plain binary64. "today" stands on the ports of
../beta_balanced_survival/beta_route_model.py.
"""
import math
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(os.path.dirname(HERE), "hypergeometric_pair_study"))
import hyper_model as H  # noqa: E402

P, M = H.P, H.M


def _exp(pair):
    return math.exp(pair[0]) * (1.0 + pair[1])


def negative_binomial_pmf(failures, successes, p, rule=True):
    """K_STATS_NegativeBinomial_PMF(x, r, p) for x > 0, r > 0 and 0 < p < 1."""
    trials = successes + failures
    if rule:
        total = P.add(H.binomial_log_pair(successes, trials, (p, 0.0)),
                      P.log_ratio(successes, (trials, 0.0)))
        return _exp(total)
    return math.exp(H._binomial_log_today(successes, trials, p)
                    + math.log(successes) - math.log(trials))


def geometric_pmf(k, p, rule=True):
    """K_STATS_Geometric_PMF(k, p) for k > 0 and 0 < p < 1."""
    if rule:
        log_q = P.neg(P.log_ratio(1.0, P.two_sum(1.0, -p)))     # Log(1 - p), a pair
        return _exp(P.add(P.mul((k, 0.0), log_q), (math.log(p), 0.0)))
    return math.exp(math.log(p) + k * M.log1p(-p))
