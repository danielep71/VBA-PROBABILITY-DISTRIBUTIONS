"""
#213: binary64 model of the candidate Student-t tail above |t| = 8 at large df.

The source before #213 serves df >= 1E5 and |t| > 8 with the incomplete-beta
continued fraction at x = df / (df + t^2), a Double within 1E-5 of one. The
fraction's leading denominator, 1 - (a + b) x / (a + 1) with a = df / 2, is a
difference of nearly equal numbers: the rounding of x, half a unit in the
last place, is multiplied by about a / (1/2 + a y), up to 1.5E6 at df = 1E8.

The candidate keeps the prefactor and gives the fraction its argument
exactly:

    y    = t^2 / (df + t^2)              formed directly, not through logs
    x    = 1 - y                          a pair of Doubles, exact
    tail = Bt * CF_pair(a, 1/2, x) / a / 2

with Bt the Loader prefactor of PROB_TryBetaRegularized and CF_pair the pair
continued fraction of #35, stopped at 2^-100.

Used for the predictions of ABOVE8_PREREGISTRATION.md. It reuses the models
of benchmark/f_inverse_large_df_study/fix_comparison.
"""
import math
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(os.path.dirname(HERE), "f_inverse_large_df_study", "fix_comparison"))
sys.path.insert(0, HERE)
import pair_cf_screen as pair  # noqa: E402
import stagnation_model as p3  # noqa: E402
import production_route_model as ldf  # noqa: E402

forward = p3.forward
STOP = 2.0 ** -100
T_MAX = 8.0
DF_MIN = 1e5


def candidate_tail(abs_t, df):
    """The candidate's upper tail at |t| > 8, df >= 1E5; None where it declines."""
    a, b = df / 2.0, 0.5
    t2 = abs_t * abs_t
    y = t2 / (df + t2)
    if not 0.0 < y < 0.5:
        return None
    x = 1.0 - y
    if not x < (a + 1.0) / (a + b + 2.0):
        return None
    log_bt = forward.beta_log_pdf(x, y, a, b, math.log(x), math.log(y)) + math.log(x) + math.log(y)
    bt = math.exp(log_bt)
    saved_model, saved_eps = pair.model, forward.C["PROB_NUM_EPS"]
    try:
        pair.model = p3
        forward.C["PROB_NUM_EPS"] = STOP
        cf, _ = pair.cf_pair(a, b, pair.DD(1.0) - pair.DD(y))
    finally:
        pair.model = saved_model
        forward.C["PROB_NUM_EPS"] = saved_eps
    return min(max(bt * cf / a, 0.0), 1.0) * 0.5


def current_tail(abs_t, df):
    """The source before #213 above 8: logistic pair, binary64 fraction."""
    ratio = math.exp(2.0 * math.log(abs_t) - math.log(df))
    return 0.5 * forward.beta_regularized(1.0 / (1.0 + ratio), ratio / (1.0 + ratio), df / 2.0, 0.5)


def tail(abs_t, df, candidate=True):
    """The upper tail as the routed source returns it for df >= 1E5."""
    if abs_t <= T_MAX:
        return ldf.route_tail(abs_t, df)
    if candidate:
        value = candidate_tail(abs_t, df)
        if value is not None:
            return value
    return current_tail(abs_t, df)
