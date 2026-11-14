"""
#314: the routed Gamma P and Q, as a model of the rule in
FAST_ROUTE_PREREGISTRATION.md.

The rule answers from the repaired binary64 route (route_model.py) wherever its
directly computed tail can be trusted, and delegates everything else to the
double-double entry points PROB_TryGammaRegularizedP and _Q, modelled by
gamma_cf_study/coverage/wide_gamma.public:

  X < fl(A + 1)   the lower series gives P.
                  P returns it when FLOOR <= value <= 1.
                  Q returns 1 - value when value <= COMPLEMENT_MAX.
  otherwise       the continued fraction gives Q.
                  Q returns it when FLOOR <= value <= UPPER_MAX.
                  P returns 1 - value when value <= UPPER_MAX.

A directly returned tail must be at least FLOOR, so that a result near the
bottom of the Double range keeps the behaviour of the double-double kernel.
A complemented tail has no floor: 1 - value is then exactly one.

There is no shape ceiling. The binary64 routines decline by themselves where
their arithmetic would leave the normal range or their iteration cap is
reached, and every decline delegates.

paths(a, x) names what each tail used: "series", "cf", "complement" or "dd".
"""
import math
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
sys.path.insert(0, os.path.join(os.path.dirname(HERE), "gamma_cf_study", "coverage"))
import route_model as rm  # noqa: E402

FLOOR = 1e-290
COMPLEMENT_MAX = 0.9
UPPER_MAX = 0.5


def _dd(a, x):
    from wide_gamma import public
    return public(a, x)


def fast(a, x):
    """(routine, value) from the binary64 route, value None where it declines
    or returns something outside [0, 1]."""
    try:
        routine, value, _ = rm.route(a, x)
    except ArithmeticError:
        # VBA raises a runtime error here (an overflow, a division by zero);
        # the routine's handler turns it into a decline.
        return ("series" if rm.below_seam(a, x) else "cf"), None
    if value is None or not math.isfinite(value) or not 0.0 <= value <= 1.0:
        return routine, None
    return routine, value


def decide(routine, value):
    """(P source, Q source) for a binary64 result, by the rule alone."""
    if value is None:
        return "dd", "dd"
    if routine == "series":
        return ("series" if value >= FLOOR else "dd",
                "complement" if value <= COMPLEMENT_MAX else "dd")
    if value > UPPER_MAX:
        return "dd", "dd"
    return "complement", ("cf" if value >= FLOOR else "dd")


def paths(a, x):
    """(P source, Q source) at finite A > 0, X > 0."""
    return decide(*fast(a, x))


def routed(a, x):
    """(P, Q) as the routed entry points return them, at finite A > 0, X > 0."""
    routine, value = fast(a, x)
    p_source, q_source = decide(routine, value)
    dd = _dd(a, x) if "dd" in (p_source, q_source) else None
    direct = {"series": value, "cf": value, "complement": None}
    p = dd[0] if p_source == "dd" else direct[p_source]
    q = dd[1] if q_source == "dd" else direct[q_source]
    if p_source == "complement":
        p = rm.sub(1.0, value)
    if q_source == "complement":
        q = rm.sub(1.0, value)
    return p, q
