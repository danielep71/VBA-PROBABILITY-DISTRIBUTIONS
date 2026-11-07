"""Study-only model of #26; no production implementation is imported by the gate.

Reuse the source-derived Loader/Stirling helpers and 64-bit-significand
intermediate model. Every assignment is rounded to binary64. The high-precision
arm evaluates exactly the same finite continued fraction to separate stopping
error from recurrence rounding. Default cap exhaustion raises, never scores.
"""

import sys
import math
import mpmath as mp
from pathlib import Path

HERE = Path(__file__).resolve().parent
ROOT = HERE.parents[1]
sys.path.insert(0, str(HERE.parent / "beta_balanced_survival"))
import beta_route_model as m

X, D = m.X, m.D


def cf(a, x, eps=3e-14, steps=None, exact=False, budget=100000):
    """Port the production recurrence, or evaluate its finite exact counterpart.

    Only the documented continued-fraction domain is modelled. ``steps`` fixes
    the finite fraction depth for decomposition; it is not a convergence claim.
    """
    if not (math.isfinite(a) and math.isfinite(x) and a > 0 and x >= a + 1):
        raise ValueError("outside the modelled CF domain")
    if budget < 1 or (steps is not None and not 1 <= steps <= budget):
        raise ValueError("invalid iteration budget/depth")
    conv = mp.mpf if exact else X
    put = (lambda x: x) if exact else D
    floor = conv(1e-300)
    a, x = conv(a), conv(x)
    b = put(x + 1 - a)
    if abs(b) < floor:
        b = put(floor)
    c = put(1 / floor)
    d = put(1 / conv(b))
    h = d
    for i in range(1, budget + 1):
        an = put(-conv(i) * (conv(i) - a))
        b = put(conv(b) + 2)
        d = put(conv(an) * conv(d) + conv(b))
        if abs(d) < floor:
            d = put(floor)
        c = put(conv(b) + conv(an) / conv(c))
        if abs(c) < floor:
            c = put(floor)
        d = put(1 / conv(d))
        delta = put(conv(d) * conv(c))
        h = put(conv(h) * conv(delta))
        if (steps and i == steps) or (not steps and abs(delta - 1) <= eps):
            return h, i
    raise RuntimeError("continued fraction exhausted its iteration budget")


def pref(a, x):
    lp = D(
        -X(m.deviance_part(a, x))
        + X(0.5) * X(math.log(a))
        - X(math.log(x))
        - X(m.C["PROB_HALF_LOG_TWO_PI"])
        - X(m.stirling_error(a))
    )
    logpref = D(X(lp) + X(math.log(x)))
    return math.exp(logpref), logpref


def read(t):
    return float(sum(mp.mpf(s) for s in t.split(";")))


def backward(a, x, n):
    """Exploratory finite-fraction reevaluation; no convergence claim of its own."""
    tail = 0.0
    for k in range(n, 0, -1):
        tail = D((-X(k) * (X(k) - X(a))) / (X(x) + 1 - X(a) + 2 * X(k) + X(tail)))
    return D(1 / (X(x) + 1 - X(a) + X(tail)))
