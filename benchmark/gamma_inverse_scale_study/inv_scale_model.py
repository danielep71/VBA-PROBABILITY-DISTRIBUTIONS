"""
#14: the Gamma inverse at the final scale, as a model of the rule in
PREREGISTRATION.md.

Where the unit-scale quantile is tiny, the lower tail is its leading term to
far below one unit in the last place:

    P(A, x) = x^A / Gamma(A + 1) * (1 - A x / (A + 1) + ...)

so  ln x = (ln p + ln Gamma(1 + A)) / A + x / (A + 1) + O(x^2).

The rule classifies a call by the plain Double

    L0 = (Log(p) + LogGamma1p(A)) / A

and answers in log space when L0 <= CROSSOVER: there x <= exp(-40), about
4.2E-18, and the neglected term is below 4.2E-18 relative. The quantile at
the final scale is formed once,

    exp((ln p + ln Gamma(1 + A)) / A + ln Scale)

with the logarithms, the quotient and the sum carried as pairs of Doubles
(the PROB_B64_Pair routines of the binary64 Gamma route), so the result does
not inherit the size of the exponent. Above the crossover the unit-scale
solver answers and its result is multiplied by the scale, as before #14.

LogGamma1p is modelled as the correctly rounded value. Production calls
PROB_TryLogGamma1p; the two may differ by a few units in the last place,
which moves the crossover by as much and the result by far less than one.
"""
import math
import os
import sys

from mpmath import loggamma, mp, mpf

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(os.path.dirname(HERE), "gamma_route_study"))
import route_model as rm  # noqa: E402

CROSSOVER = -40.0
MIN_NORMAL = 2.2250738585072014e-308


def log_gamma1p(a):
    with mp.workdps(40):
        return float(loggamma(1 + mpf(a)))


def log_unit(p, a):
    """L0, the plain Double the rule classifies by."""
    return rm.div(rm.add(math.log(p), log_gamma1p(a)), a)


def uses_log(p, a):
    return log_unit(p, a) <= CROSSOVER


PAIR_LOG_LIMIT = math.ldexp(1.0, 978)      # exponent field 2001
SHIFT = math.ldexp(1.0, 128)


def _log_pair(v):
    # PROB_B64_PairLog is defined for exponent fields 1 to 2000. Outside it
    # the argument is moved by an exact factor of 2^128, whose logarithm is
    # the pair logarithm of that power of two (Addendum 1).
    if v < MIN_NORMAL:
        hi, lo = rm.pair_log(v * SHIFT, 0.0)
        sh, sl = rm.pair_log(SHIFT, 0.0)
        return rm.pair_add(hi, lo, -sh, -sl)
    if v >= PAIR_LOG_LIMIT:
        hi, lo = rm.pair_log(v / SHIFT, 0.0)
        sh, sl = rm.pair_log(SHIFT, 0.0)
        return rm.pair_add(hi, lo, sh, sl)
    return rm.pair_log(v, 0.0)


def candidate(p, a, scale):
    """The log-space quantile at the final scale: a Double, 0.0 where the
    exponential underflows, or None where it overflows (#NUM! in VBA)."""
    lh, ll = _log_pair(p)
    nh, nl = rm.pair_add(lh, ll, log_gamma1p(a), 0.0)
    qh, ql = rm.pair_div(nh, nl, a)
    sh, sl = _log_pair(scale)
    fh, fl = rm.pair_add(qh, ql, sh, sl)
    try:
        e = math.exp(fh)
    except OverflowError:
        return None
    return rm.add(e, rm.mul(e, fl))


def exact(p, a, scale=1.0, digits=60):
    """The quantile at the final scale, as an mpf, for a unit quantile small
    enough for three terms of the lower series (x below about 1E-6)."""
    with mp.workdps(digits):
        fa = mpf(a)
        log0 = (mp.log(mpf(p)) + loggamma(1 + fa)) / fa
        x = mp.exp(log0)
        for _ in range(4):
            series = 1 + x / (fa + 1) + x * x / ((fa + 1) * (fa + 2)) \
                + x ** 3 / ((fa + 1) * (fa + 2) * (fa + 3))
            x = mp.exp(log0 + (x - mp.log(series)) / fa)
        return x * mpf(scale)


# --------------------------------------------------------------------------
# Second candidate (PREREGISTRATION_2.md): LogGamma(1 + A) from the degree-32
# polynomial PROB_GammaSmallLogGamma1p up to shape 2, which route_model
# reproduces bit for bit, and from PROB_TryLogGamma1p above it.

def log_gamma1p_2(a):
    if a <= 1.0:
        return rm.small_log_gamma1p(a)
    if a <= 2.0:
        # Gamma(1 + A) = A Gamma(A), and A - 1 is exact for A in (1, 2]
        return rm.add(rm.small_log_gamma1p(rm.sub(a, 1.0)), math.log(a))
    return log_gamma1p(a)


def log_unit_2(p, a):
    return rm.div(rm.add(math.log(p), log_gamma1p_2(a)), a)


def uses_log_2(p, a):
    return log_unit_2(p, a) <= CROSSOVER


def candidate_2(p, a, scale):
    lh, ll = _log_pair(p)
    nh, nl = rm.pair_add(lh, ll, log_gamma1p_2(a), 0.0)
    qh, ql = rm.pair_div(nh, nl, a)
    sh, sl = _log_pair(scale)
    fh, fl = rm.pair_add(qh, ql, sh, sl)
    try:
        e = math.exp(fh)
    except OverflowError:
        return None
    return rm.add(e, rm.mul(e, fl))
