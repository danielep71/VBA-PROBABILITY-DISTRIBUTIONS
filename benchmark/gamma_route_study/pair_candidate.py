"""B4 development: unscaled pair recurrences, scaled prefactor/composition.

No reference/oracle calls in the numerical path. This is not a VBA port and
has no host fidelity or cost evidence. Bounds deliberately decline outside
a development envelope; they are NOT a frozen public routing rule.
"""
import math
from functools import lru_cache

import preflight_phase_b as old  # noqa: F401; establishes the audited model import paths
from precision_probe import DD
from scaled_dd import ScaledDD as S
from scaled_gamma import ONE, log_scaled, exp_scaled
from extended_gamma import log_gamma1p, log_pair, lower_pair
from wide_gamma import deviance, stirling, LOG_SQRT_2PI

MAX_ITER = 100000


class Declined(ArithmeticError):
    """Candidate cannot safely answer in its development envelope."""


def checked(v):
    if not math.isfinite(v.hi) or not math.isfinite(v.lo):
        raise Declined('unscaled recurrence range')
    # Protect Dekker splitting and keep significant low parts normal.
    if v.hi and not 2.**-450 <= abs(v.hi) <= 2.**450:
        raise Declined('unscaled recurrence magnitude')
    return v


@lru_cache(maxsize=1)
def lg(a):
    return log_gamma1p(S.from_float(a))


def prefactor(a, x):
    if a <= 128:
        return log_scaled(x)*S.from_float(a)-lg(a)-S.from_float(x)
    return -deviance(a, x)-log_scaled(a)*.5-LOG_SQRT_2PI-stirling(S.from_float(a))


def public(a, x, trace=None):
    if trace is None:
        trace = {}
    if not (math.isfinite(a) and math.isfinite(x) and 1e-8 <= a <= 10000 and 1e-12 <= x <= 20000):
        raise Declined('outside development envelope')
    aa, xx = DD(a), DD(x)
    lp = prefactor(a, x)
    if x < a+1:
        term, total = DD(1), DD(1)
        for n in range(1, MAX_ITER+1):
            term = checked(checked(term*xx)/(aa+n))
            total = checked(total+term)
            ratio = checked(xx/(aa+n+1))
            gap = DD(1)-ratio
            if gap.hi > 0:
                remainder = checked(checked(term*ratio)/gap)
                if abs(remainder.hi)*4 < abs(total.hi)*2.**-100:
                    break
        else:
            raise Declined('lower cap')
        p, q = lower_pair(lp+log_pair(S._normalize(total, 0)))
        trace.update(route='pair lower', iterations=n)
    else:
        base = xx+1-aa
        b = base
        # Sentinel replaces an infinite first Lentz denominator. 2^400 keeps
        # Dekker products in range; its perturbation is below DD precision.
        c, d = DD(2.**400), checked(DD(1)/b)
        for n in range(1, MAX_ITER+1):
            an = DD(-n)*(DD(n)-aa)
            b = checked(b+2)
            d = checked(DD(1)/checked(an*d+b))
            c = checked(b+checked(an/c))
            delta = checked(c*d)-1
            if abs(delta.hi) < 2.**-100:
                break
        else:
            raise Declined('upper cap')
        tail = DD(0)
        for k in range(n, 0, -1):
            tail = checked(DD(-k)*(DD(k)-aa)/checked(base+2*k+tail))
        h = checked(DD(1)/checked(base+tail))
        q = exp_scaled(lp)*S.from_float(a)*S._normalize(h, 0)
        p = ONE-q
        trace.update(route='pair upper', iterations=n)
    values = p.rounded(), q.rounded()
    if not all(math.isfinite(v) and 0 <= v <= 1 for v in values):
        raise Declined('invalid probability')
    return values
