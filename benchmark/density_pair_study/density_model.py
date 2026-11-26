"""
#501: the Beta, F and Student t densities with their large terms carried as
pairs of Doubles, on the pair arithmetic of #498 (../deviance_pair_study).

Three densities, three causes:

  Beta        the two Loader deviances bd0(A, N X) and bd0(B, N Y) in one
              Double each, with N = A + B, N X and N Y rounded before the
              deviance sees them;
  F           the same, and before it the arguments U = r / (1 + r) and
              V = 1 / (1 + r), r = X d1 / d2, formed through an exponential
              and rounded;
  Student t   no deviance: the product (df + 1) / 2 * Log(1 + t^2 / df),
              rounded at its own size, and a constant from a difference of
              log-gammas.

The candidate carries every one of these as a pair to Exp(high) * (1 + low).
For the Student t it takes Loader's form, which has no log-gamma:

  f(t) = Exp(-bd0(df/2, (df+1)/2) + d((df+1)/2) - d(df/2)
             + (df+1)/2 * Log(df / (df + t^2))) / Sqr(2 Pi)

with d the Stirling error.

This file is a model in plain binary64, as pair_model.py is. "today" is the
port of the source: the Beta density of ../beta_balanced_survival, and ports
written here of K_STATS_F_Density and PROB_TryStudentTPDF, which stand on
that file's procedures. None is bit-faithful to Excel everywhere.
"""
import math
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(os.path.dirname(HERE), "deviance_pair_study"))
import pair_model as P  # noqa: E402

M = P.M
HALF_LOG_TWO_PI = P.HALF_LOG_TWO_PI


RATIO_MIN, RATIO_MAX = M.C["PROB_BD0_RATIO_MIN"], M.C["PROB_BD0_RATIO_MAX"]
X, R = M.X, M.D            # 64-bit intermediates inside an expression; the store rounds


def _total(pair, doubles):
    for v in doubles:
        pair = P.add(pair, (v, 0.0))
    return pair


def _exp(pair):
    """PROB_TryExpPair: Exp(high) * (1 + low), one expression."""
    return R(X(math.exp(pair[0])) * (1 + X(pair[1])))


def _in_range(*values):
    return all(math.isfinite(v) and RATIO_MIN <= v <= RATIO_MAX for v in values)


def _beta_core(x, y, a, b):
    """PROB_TryBetaCorePair: the Loader log-density of the Beta without its two
    logarithms, as a pair; x and y pairs. None where a deviance has no pair."""
    n = P.two_sum(a, b)
    mx, my = P.mul(n, x), P.mul(n, y)
    if not _in_range(a, b, mx[0], my[0]):
        return None
    core = P.add(P.neg(P.deviance(a, mx)), P.neg(P.deviance(b, my)))
    half_logs = R(X(0.5) * (X(math.log(a)) + X(math.log(b)) - X(math.log(n[0]))))
    return _total(core, (half_logs, -HALF_LOG_TWO_PI, -M.stirling_error(a),
                         -M.stirling_error(b), M.stirling_error(n[0])))


def beta_density_pair(x, a, b):
    """PROB_TryBetaDensityPair, or None where it returns FALSE."""
    core = _beta_core((x, 0.0), P.two_sum(1.0, -x), a, b)
    if core is None:
        return None
    return _exp(_total(core, (-math.log(x), -M.log1p(-x))))


def beta_density(x, a, b, rule=True):
    if not rule:
        return M.beta_density(x, a, b, pair=False)
    value = beta_density_pair(x, a, b)
    return M.beta_density(x, a, b, pair=False) if value is None else value


def f_density(x, d1, d2, rule=True):
    a, b = 0.5 * d1, 0.5 * d2
    if rule:
        r = P.div(P.two_prod(x, d1), (d2, 0.0))
        one_plus = P.add((1.0, 0.0), r)
        core = _beta_core(P.div(r, one_plus), P.div((1.0, 0.0), one_plus), a, b)
        if core is not None:
            return _exp(_total(core, (-math.log(x),)))
    log_ratio = math.log(x) + math.log(d1) - math.log(d2)
    if log_ratio >= 0:
        e = math.exp(-log_ratio)
        log_one_plus = M.log1p(e)
        u, v = 1.0 / (1.0 + e), e / (1.0 + e)
        log_u, log_v = -log_one_plus, -log_ratio - log_one_plus
    else:
        ratio = math.exp(log_ratio)
        log_one_plus = math.log1p(ratio)
        u, v = ratio / (1.0 + ratio), 1.0 / (1.0 + ratio)
        log_u, log_v = log_ratio - log_one_plus, -log_one_plus
    return math.exp(M.beta_log_pdf(u, v, a, b, log_u, log_v) + log_u + log_v - math.log(x))


def student_t_density(t, df, rule=True):
    if rule:
        half, half_plus = 0.5 * df, 0.5 * (df + 1.0)
        log_part = P.mul((half_plus, 0.0), P.log_ratio(df, P.add(P.two_prod(t, t), (df, 0.0))))
        core = P.add(log_part, P.neg(P.deviance(half, half_plus)))
        return _exp(_total(core, (M.stirling_error(half_plus), -M.stirling_error(half),
                                  -HALF_LOG_TWO_PI)))
    constant = M.log_gamma_half_diff(0.5 * df) - 0.5 * (math.log(df) + math.log(math.pi))
    return math.exp(constant - 0.5 * (df + 1.0) * math.log1p(t * t / df))
