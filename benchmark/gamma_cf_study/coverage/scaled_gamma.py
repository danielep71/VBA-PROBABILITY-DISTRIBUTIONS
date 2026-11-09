"""Revealed small-shape Gamma experiment; not public routing or acceptance.

Combines the tested scaled arithmetic with the frozen polynomial coefficients.
The domain is finite positive binary64 a<=1 and x<=32. Series convergence
screens are heuristic; they do not certify accumulated error.
"""
import math
from scaled_dd import ScaledDD as S
from precision_probe import LN2
from precision_v2 import COEFFICIENTS

ONE = S.from_float(1.)
LN_TWO = S._normalize(LN2, 0)


def small(term, total):
    return term.hi == 0 or (total.hi != 0 and term.exponent < total.exponent-108)


def log_scaled(x):
    if not math.isfinite(x) or x <= 0:
        raise ValueError('positive finite log input required')
    f, e = math.frexp(x)
    z = (S.from_float(f)-1)/(S.from_float(f)+1)
    square = z*z
    term, total = z, z
    for n in range(3, 1001, 2):
        term = term*square
        delta = term/n
        total = total+delta
        if small(delta, total):
            return total*2+LN_TWO*e
    if z.hi == 0:
        return LN_TWO*e
    raise RuntimeError('log cap')


def exp_scaled(z):
    value = z.rounded()
    if not -1500 <= value <= 8:
        raise ValueError('outside exploratory exponential range')
    n = round(value/LN2.hi)
    r = z-LN_TWO*n
    term, total = ONE, ONE
    for k in range(1, 1001):
        term = term*r/k
        total = total+term
        if small(term, total):
            return S(total.hi, total.lo, total.exponent+n)
    raise RuntimeError('exp cap')


def expm1_scaled(z):
    if abs(z.rounded()) > .5:
        raise ValueError('outside exploratory expm1 range')
    term, total = z, z
    for k in range(2, 1001):
        term = term*z/k
        total = total+term
        if small(term, total):
            return total
    raise RuntimeError('expm1 cap')


def log1p_scaled(z):
    if not -.8 <= z.rounded() <= 0:
        raise ValueError('outside exploratory log1p range')
    t = z/(z+2)
    square = t*t
    term, total = t, t
    for k in range(3, 2001, 2):
        term = term*square
        delta = term/k
        total = total+delta
        if small(delta, total):
            return total*2
    raise RuntimeError('log1p cap')


def public(a, x):
    if not (math.isfinite(a) and math.isfinite(x) and 0 < a <= 1 and 0 < x <= 32):
        raise ValueError('outside scaled small-shape study domain')
    aa, xx = S.from_float(a), S.from_float(x)
    u = aa-.5
    lg = S._normalize(COEFFICIENTS[-1], 0)
    for c in reversed(COEFFICIENTS[:-1]):
        lg = lg*u+S._normalize(c, 0)
    lg = aa*(aa-1)*lg
    if x < a+1:
        term = -xx
        total = term/(aa+1)
        for n in range(2, 1001):
            term = term*(-xx)/n
            delta = term/(aa+n)
            total = total+delta
            if small(delta, total):
                break
        else:
            raise RuntimeError('lower cap')
        lp = log_scaled(x)*aa-lg+log1p_scaled(total*aa)
        if lp.hi > 0:
            raise RuntimeError('positive log probability')
        if lp.rounded() >= -.5:
            q = -expm1_scaled(lp)
            p = ONE-q
        else:
            p = exp_scaled(lp)
            q = ONE-p
    else:
        b = xx+1-aa
        c, d = S.from_float(1e300), ONE/b
        for n in range(1, 100001):
            an = S.from_float(-n)*(S.from_float(n)-aa)
            b = b+2
            d = ONE/(an*d+b)
            c = b+an/c
            delta = c*d-1
            if delta.hi == 0 or delta.exponent < -100:
                break
        else:
            raise RuntimeError('upper cap')
        tail = S.from_float(0.)
        for k in range(n, 0, -1):
            tail = S.from_float(-k)*(S.from_float(k)-aa)/(xx+1-aa+2*k+tail)
        h = ONE/(xx+1-aa+tail)
        q = exp_scaled(log_scaled(x)*aa-lg-xx)*aa*h
        p = ONE-q
    return p.rounded(), q.rounded()
