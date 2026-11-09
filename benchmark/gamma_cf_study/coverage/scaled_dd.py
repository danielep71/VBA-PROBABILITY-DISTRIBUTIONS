"""Study-only scaled two-Double arithmetic; frozen v2 remains unchanged.

Each value is (hi+lo)*2**exponent. The mantissa operation graphs are the
existing audited DD graphs, executed on bounded operands. Results are
approximations, not correctly rounded real operations. rounded() rounds the
represented pair only. No high-precision arithmetic runs in this module.
"""
from dataclasses import dataclass
import math
from precision_probe import DD
from scaled_product import rounded

MAX_EXPONENT = 32767  # explicit study bound, not a public Gamma-domain policy
ALIGNMENT_LIMIT = 120


@dataclass(frozen=True)
class ScaledDD:
    hi: float
    lo: float = 0.
    exponent: int = 0

    def __post_init__(self):
        if not isinstance(self.exponent, int) or isinstance(self.exponent, bool):
            raise ValueError('integer exponent required')
        if abs(self.exponent) > MAX_EXPONENT:
            raise ValueError('outside scaled arithmetic study exponent range')
        if not math.isfinite(self.hi) or not math.isfinite(self.lo):
            raise ValueError('finite mantissa required')
        if self.hi == 0:
            if self.lo != 0 or self.exponent != 0:
                raise ValueError('noncanonical zero')
        elif (not .5 <= abs(self.hi) < 1 or
              abs(self.lo) > math.ulp(self.hi)*.5 or self.hi+self.lo != self.hi):
            raise ValueError('normalized mantissa required')

    @classmethod
    def from_float(cls, x):
        x = float(x)
        if not math.isfinite(x):
            raise ValueError('finite input required')
        h, e = math.frexp(x)
        return cls(h, 0., e)

    @classmethod
    def _normalize(cls, value, exponent):
        h, l = DD.sum(value.hi, value.lo)
        if h == 0:
            return cls(h)
        h, shift = math.frexp(h)
        return cls(h, math.ldexp(l, -shift), exponent+shift)

    @staticmethod
    def _coerce(value):
        return value if isinstance(value, ScaledDD) else ScaledDD.from_float(value)

    def __neg__(self):
        return ScaledDD(-self.hi, -self.lo, self.exponent)

    def __add__(self, other):
        a, b = self, self._coerce(other)
        if a.hi == 0:
            return b if b.hi else ScaledDD(a.hi+b.hi)
        if b.hi == 0:
            return a
        if a.exponent < b.exponent:
            a, b = b, a
        gap = a.exponent-b.exponent
        if gap > ALIGNMENT_LIMIT:
            # Smaller magnitude < 2**b.exponent. Relative omission error
            # < 2**-119 for gap >= 121; no cancellation is possible here.
            return a
        small = DD(math.ldexp(b.hi, -gap), math.ldexp(b.lo, -gap))
        value = DD(a.hi, a.lo) + small
        return self._normalize(value, a.exponent)

    __radd__ = __add__

    def __sub__(self, other):
        return self + -self._coerce(other)

    def __mul__(self, other):
        b = self._coerce(other)
        if self.hi == 0 or b.hi == 0:
            return ScaledDD(self.hi*b.hi)
        value = DD(self.hi, self.lo)*DD(b.hi, b.lo)
        return self._normalize(value, self.exponent+b.exponent)

    __rmul__ = __mul__

    def __truediv__(self, other):
        b = self._coerce(other)
        if b.hi == 0:
            raise ZeroDivisionError('scaled divisor is zero')
        if self.hi == 0:
            return ScaledDD(self.hi/b.hi)
        value = DD(self.hi, self.lo)/DD(b.hi, b.lo)
        return self._normalize(value, self.exponent-b.exponent)

    def rounded(self):
        return rounded(self.hi, self.lo, self.exponent)
