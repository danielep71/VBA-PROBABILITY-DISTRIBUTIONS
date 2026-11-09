"""Offline outward-rounded enclosure of DLMF 8.9.2, only 0<a<=1, x>0.

Positive continued-fraction coefficients bracket the limit between consecutive
convergents. Decimal directed rounding encloses each finite convergent. This
bounds the fraction H=x**(-a)*exp(x)*Gamma(a,x), NOT its Gamma Q prefactor.
It is independent of the production Lentz stopping test; not VBA runtime code.
"""
import math
from decimal import Context, Decimal, ROUND_CEILING, ROUND_FLOOR


class PositiveIntervals:
    def __init__(self, precision=120):
        self.down = Context(prec=precision, rounding=ROUND_FLOOR)
        self.up = Context(prec=precision, rounding=ROUND_CEILING)

    @staticmethod
    def point(x):
        d = Decimal(x)
        return d, d

    def add(self, a, b):
        return self.down.add(a[0], b[0]), self.up.add(a[1], b[1])

    def mul(self, a, b):
        return self.down.multiply(a[0], b[0]), self.up.multiply(a[1], b[1])

    def div(self, a, b):
        if b[0] <= 0:
            raise ValueError('nonpositive interval divisor')
        return self.down.divide(a[0], b[1]), self.up.divide(a[1], b[0])

    def integer_minus(self, n, a):
        return self.down.subtract(Decimal(n), a[1]), self.up.subtract(Decimal(n), a[0])


def enclose(a, x, *, bits=105, cap=4096):
    if not math.isfinite(a) or not math.isfinite(x) or not 0 < a <= 1 or x <= 0:
        raise ValueError('requires finite 0<a<=1 and x>0')
    ar = PositiveIntervals()
    aa, xx = ar.point(a), ar.point(x)
    zero, one = ar.point(0), ar.point(1)
    previous = ar.div(one, xx)
    if a == 1:
        return previous[0], previous[1], 0
    am2, am1, bm2, bm1 = one, xx, zero, one
    # 2^-bits is rounded DOWN so acceptance cannot exploit tolerance rounding.
    tolerance = ar.down.power(Decimal(2), -bits)
    for n in range(1, cap + 1):
        an = ar.integer_minus((n + 1)//2, aa) if n % 2 else ar.point(n//2)
        bn = one if n % 2 else xx
        av = ar.add(ar.mul(bn, am1), ar.mul(an, am2))
        bv = ar.add(ar.mul(bn, bm1), ar.mul(an, bm2))
        current = ar.div(bv, av)
        lower, upper = min(current[0], previous[0]), max(current[1], previous[1])
        width = ar.up.subtract(upper, lower)
        if width <= ar.down.multiply(lower, tolerance):
            return lower, upper, n
        am2, am1, bm2, bm1 = am1, av, bm1, bv
        previous = current
    raise RuntimeError('positive CF enclosure did not converge')
