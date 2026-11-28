"""
#365: the Student t tail beyond |t| = 16 below df = 1E5, with the prefactor of
the incomplete beta, its two arguments and its continued fraction carried as
pairs of Doubles.

The library takes the tail there as half of I_x(df/2, 1/2) at
x = df / (df + t^2), through PROB_TryBetaRegularized: the logarithm of the
prefactor x^a y^b / B(a, b) in one Double, of the size of t^2 / 2, then one
exponential, and the continued fraction at x rounded.

The candidate forms

    r = t^2 / df,  y = r / (1 + r),  x = 1 / (1 + r)        as pairs
    core = Log(x^a y^b / B(a, b))                           as a pair (#501)
    tail = Exp(core.high) * (1 + core.low) * CF(a, 1/2; x) / a / 2

with CF the pair continued fraction of #35 at the pair x.

This file is a model in plain binary64. "today" is the routed model of
../student_t_large_df_study (below1e5_model.rule_tail); the Beta core is that
of ../density_pair_study and the fraction that of
../f_inverse_large_df_study/fix_comparison. None is bit-faithful to Excel.
"""
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
BENCH = os.path.dirname(HERE)
sys.path.insert(0, os.path.join(BENCH, "student_t_large_df_study"))
sys.path.insert(0, os.path.join(BENCH, "density_pair_study"))
import above8_model as A8  # noqa: E402
import below1e5_model as B5  # noqa: E402
import density_model as D  # noqa: E402

P = D.P
T_MIN = 16.0             # the rule serves |t| above it
DF_SEAM = 1e5            # and df below it


def served(abs_t, df):
    """Whether the rule is tried at (|t|, df)."""
    return abs_t > T_MIN and df < DF_SEAM


def _fraction(a, b, x):
    """PROB_TryBetaContinuedFractionPair at the pair x, stopped at 2^-100."""
    pair, forward = A8.pair, A8.forward
    saved_model, saved_eps = pair.model, forward.C["PROB_NUM_EPS"]
    try:
        pair.model = A8.p3
        forward.C["PROB_NUM_EPS"] = A8.STOP
        value, _ = pair.cf_pair(a, b, pair.DD(x[0]) + pair.DD(x[1]))
    except ValueError:
        return None
    finally:
        pair.model = saved_model
        forward.C["PROB_NUM_EPS"] = saved_eps
    return float(value)


def pair_tail(abs_t, df):
    """The candidate's upper tail, or None where it declines."""
    if not D._in_range(abs_t, df):
        return None
    ratio = abs_t / df * abs_t
    if not D._in_range(ratio):
        return None
    a, b = 0.5 * df, 0.5
    r = P.div(P.two_prod(abs_t, abs_t), (df, 0.0))
    one_plus = P.add((1.0, 0.0), r)
    y, x = P.div(r, one_plus), P.div((1.0, 0.0), one_plus)
    if not x[0] < (a + 1.0) / (a + 2.5):
        return None
    core = D._beta_core(x, y, a, b)
    if core is None:
        return None
    cf = _fraction(a, b, x)
    if cf is None:
        return None
    return 0.5 * (D._exp(core) * cf / a)


def tail(abs_t, df, rule=True):
    """P(T > |t|) for df < 1E5, as the routed source returns it."""
    if rule and served(abs_t, df):
        value = pair_tail(abs_t, df)
        if value is not None:
            return value
    return B5.rule_tail(abs_t, df)
