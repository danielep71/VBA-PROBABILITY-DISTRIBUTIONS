"""
#365: binary64 model of the candidate Student-t tail below df = 1E5.

The rule of BELOW1E5_PREREGISTRATION.md extends the rule of #213 downward:

    df < 1E5 and |t| <= 16:
        y    = t^2 / (df + t^2)              formed directly
        tail = PROB_TryBetaLowerNearOne(y, df / 2, 1/2) / 2

PROB_TryBetaLowerNearOne serves a + b >= 1000 (df >= 1999) and an argument
below the mean (t^2 > 3 df / (df + 2)); where it declines, and for
|t| > 16, the route of today answers. Nothing changes for df >= 1E5.

The model reuses above8_model: candidate_tail is the near-one rule and
current_tail the route of today. It is not bit-faithful to Excel (14 of the
314 map rows agree bit for bit), so its figures bound the rule; they do not
predict observations.
"""
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import above8_model as M  # noqa: E402

DF_SEAM = 1e5
T_MAX = 16.0
LOADER_MIN_SHAPE = 1000.0


def served(abs_t, df):
    """Whether the rule answers at (|t|, df); the tests of the source, in its order."""
    if not (df < DF_SEAM and abs_t <= T_MAX):
        return False
    a, b = df / 2.0, 0.5
    t2 = abs_t * abs_t
    y = t2 / (df + t2)
    if not 0.0 < y < 0.5 or a + b < LOADER_MIN_SHAPE:
        return False
    return 1.0 - y < (a + 1.0) / (a + b + 2.0)


def rule_tail(abs_t, df):
    """The upper tail under the rule, for df < 1E5."""
    if served(abs_t, df):
        value = M.candidate_tail(abs_t, df)
        if value is not None:
            return value
    return M.current_tail(abs_t, df)
