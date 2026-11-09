"""Exploratory two-Double lower-tail arithmetic; not VBA or acceptance evidence.

All elementary operations below use separate binary64 assignments. No FMA or
arbitrary precision is used in the numerical path. Range limits intentionally
fail closed; extending this prototype to the public domain requires a design.
"""
import math


class DD:
    def __init__(self, hi, lo=0.0):
        self.hi, self.lo = float(hi), float(lo)

    @staticmethod
    def sum(a, b):
        s = a + b
        v = s - a
        av = s - v
        ar = a - av
        br = b - v
        return s, ar + br

    @staticmethod
    def product(a, b):
        p = a * b
        ca = 134217729.0 * a
        cb = 134217729.0 * b
        ah = ca - (ca - a)
        bh = cb - (cb - b)
        al, bl = a - ah, b - bh
        e = ah * bh - p
        e = e + ah * bl
        e = e + al * bh
        e = e + al * bl
        return p, e

    def __add__(self, other):
        b = other if isinstance(other, DD) else DD(other)
        s, e = self.sum(self.hi, b.hi)
        t, f = self.sum(self.lo, b.lo)
        e = e + t
        s, e = self.sum(s, e)
        e = e + f
        return DD(*self.sum(s, e))

    __radd__ = __add__

    def __neg__(self):
        return DD(-self.hi, -self.lo)

    def __sub__(self, other):
        return self + -(other if isinstance(other, DD) else DD(other))

    def __mul__(self, other):
        b = other if isinstance(other, DD) else DD(other)
        p, e = self.product(self.hi, b.hi)
        e = e + self.hi * b.lo
        e = e + self.lo * b.hi
        e = e + self.lo * b.lo
        return DD(*self.sum(p, e))

    __rmul__ = __mul__

    def __truediv__(self, other):
        b = other if isinstance(other, DD) else DD(other)
        if b.hi == 0:
            raise ValueError('zero divisor')
        q = DD(self.hi / b.hi)
        for _ in range(2):
            r = self - q * b
            q = q + r.hi / b.hi
        return q

    def value(self):
        return self.hi + self.lo


# Nearest binary64 plus its nearest binary64 residual, fixed independently.
LN2 = DD(float.fromhex('0x1.62e42fefa39efp-1'),
         float.fromhex('0x1.abc9e3b39803fp-56'))


def log_dd(x):
    if not math.isfinite(x) or not 2**-500 <= x <= 2**500:
        raise ValueError('outside exploratory log range')
    f, e = math.frexp(x)
    # f in [.5,1); atanh argument magnitude <= 1/3.
    z = (DD(f) - 1) / (DD(f) + 1)
    square = z * z
    term, total = z, z
    for n in range(3, 1001, 2):
        term = term * square
        delta = term / n
        total = total + delta
        if abs(delta.hi) <= abs(total.hi) * 2**-105:
            return total * 2 + LN2 * e
    raise RuntimeError('log series cap')


def log1p_dd(z):
    if not -0.8 <= z.value() <= 0:
        raise ValueError('outside exploratory log1p range')
    t = z / (z + 2)
    square = t * t
    term, total = t, t
    for n in range(3, 2001, 2):
        term = term * square
        delta = term / n
        total = total + delta
        if abs(delta.hi) <= abs(total.hi) * 2**-105:
            return total * 2
    raise RuntimeError('log1p series cap')


def expm1_dd(z):
    if not -8 <= z.value() <= 0:
        raise ValueError('outside exploratory expm1 range')
    term, total = z, z
    for n in range(2, 1001):
        term = term * z / n
        total = total + term
        if abs(term.hi) <= abs(total.hi) * 2**-105:
            return total
    raise RuntimeError('expm1 series cap')


def lower(a, x, loggamma1p):
    if not 2**-500 <= a <= 1 or not 2**-500 <= x < a + 1:
        raise ValueError('outside exploratory lower-tail range')
    term = DD(-x)
    total = term / (DD(a) + 1)
    for n in range(2, 1001):
        term = term * -x / n
        delta = term / (DD(a) + n)
        total = total + delta
        if abs(delta.hi) <= abs(total.hi) * 2**-105:
            break
    else:
        raise RuntimeError('lower series cap')
    lp = log_dd(x) * a - loggamma1p + log1p_dd(total * a)
    if lp.value() > 0:
        raise ValueError('positive log probability')
    em1 = expm1_dd(lp)
    return (em1 + 1).value(), (-em1).value()
