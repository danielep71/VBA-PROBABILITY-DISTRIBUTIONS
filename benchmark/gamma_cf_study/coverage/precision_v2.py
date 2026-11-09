"""Second, bounded study prototype. Frozen v1 remains unchanged.

Use two-Double arithmetic for the complete upper prefactor and backward
fraction. The lower branch retains the v1 formula with exact log(1).
This is not a full-domain production implementation.
"""
import math
import json
from pathlib import Path
from precision_probe import DD, LN2, log_dd as old_log, log1p_dd, expm1_dd


def log_dd(x):
    if x == 1.0:
        return DD(0.)
    return old_log(x)


def exp_dd(z):
    if not math.isfinite(z.value()) or not -64 <= z.value() <= 8:
        raise ValueError('outside v2 exploratory exp range')
    n = round(z.hi / LN2.hi)
    r = z - LN2 * n
    term, total = DD(1), DD(1)
    for k in range(1, 1001):
        term = term * r / k
        total = total + term
        if abs(term.hi) <= abs(total.hi) * 2**-105:
            return DD(math.ldexp(total.hi, n), math.ldexp(total.lo, n))
    raise RuntimeError('exp cap')


def lower(a, x, loggamma1p):
    if not 2**-400 <= a <= 1 or not 2**-400 <= x < a + 1:
        raise ValueError('outside v2 lower domain')
    term = DD(-x)
    total = term / (DD(a) + 1)
    for n in range(2, 1001):
        term = term * -x / n
        delta = term / (DD(a) + n)
        total = total + delta
        if abs(delta.hi) <= abs(total.hi) * 2**-105:
            break
    else:
        raise RuntimeError('lower cap')
    lp = log_dd(x) * a - loggamma1p + log1p_dd(total * a)
    if lp.value() > 0:
        raise ValueError('positive log probability')
    if lp.value() >= -0.5:
        em1 = expm1_dd(lp)
        return (em1 + 1).value(), (-em1).value()
    p = exp_dd(lp)
    return p.value(), (DD(1) - p).value()


def upper(a, x, loggamma1p):
    if not 2**-400 <= a <= 1 or not a + 1 <= x <= 32:
        raise ValueError('outside v2 upper domain')
    b = DD(x) + 1 - a
    c = DD(1e300)
    d = DD(1) / b
    h = d
    for n in range(1, 100001):
        an = DD(-n) * (DD(n) - a)
        b = b + 2
        d = an * d + b
        c = b + an / c
        if abs(d.hi) < 1e-300 or abs(c.hi) < 1e-300:
            raise RuntimeError('unexpected near-zero denominator')
        d = DD(1) / d
        delta = c * d
        h = h * delta
        if abs((delta - 1).value()) <= 2**-100:
            break
    else:
        raise RuntimeError('upper cap')
    tail = DD(0)
    for k in range(n, 0, -1):
        tail = DD(-k) * (DD(k) - a) / (DD(x) + 1 - a + 2*k + tail)
    h = DD(1) / (DD(x) + 1 - a + tail)
    e = log_dd(x) * a - loggamma1p - x
    q = exp_dd(e) * a * h
    return (DD(1) - q).value(), q.value()


COEFFICIENTS = [DD(float.fromhex(hi), float.fromhex(lo)) for hi, lo in
                json.loads(Path(__file__).with_name('precision_coefficients.json').read_text())['coefficients_hi_lo_hex']]


def loggamma_dd(a):
    if not 0 <= a <= 1:
        raise ValueError('outside loggamma range')
    u = DD(a) - .5
    v = COEFFICIENTS[-1]
    for c in reversed(COEFFICIENTS[:-1]):
        v = v * u + c
    return DD(a) * (DD(a) - 1) * v


def public(a, x):
    lg = loggamma_dd(a)
    return lower(a, x, lg) if x < a + 1 else upper(a, x, lg)
