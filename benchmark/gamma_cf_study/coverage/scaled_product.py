"""Study-only exponent-separated product and final binary64 rounding.

The exact product of two finite binary64 operands is retained as (hi+lo)*2**e.
Only bounded binary64 arithmetic is used in the numerical path. This is NOT a
replacement DD class, a Gamma kernel, or an Excel-validated implementation.
Frozen precision_probe/precision_v2 code is deliberately not modified.
"""
import math
from precision_probe import DD


def product(a, b):
    """Return normalized hi, lo, exponent without materializing a*b."""
    if not math.isfinite(a) or not math.isfinite(b):
        raise ValueError('finite binary64 operands required')
    if a == 0 or b == 0:
        return a * b, 0.0, 0
    am, ae = math.frexp(a)
    bm, be = math.frexp(b)
    # |am|, |bm| in [.5,1): splitter, products and nonzero residuals
    # are normal and finite; scaling by powers of two is exact here.
    hi, lo = DD.product(am, bm)
    hi, shift = math.frexp(hi)
    lo = math.ldexp(lo, -shift)
    return hi, lo, ae + be + shift


def rounded(hi, lo, exponent):
    """Round a normalized scaled pair, ties to even; raise on overflow.

    Subnormal results are rounded once in units of 2**-1074, retaining the
    low component at halfway cases. Scaling hi and lo separately to the final
    exponent and adding them would introduce double rounding.
    """
    if not isinstance(exponent, int) or isinstance(exponent, bool):
        raise ValueError('integer exponent required')
    if not math.isfinite(hi) or not math.isfinite(lo):
        raise ValueError('finite components required')
    if hi == 0:
        if lo != 0:
            raise ValueError('noncanonical zero')
        return hi
    if not .5 <= abs(hi) < 1:
        raise ValueError('normalized high component required')
    if abs(lo) > math.ulp(hi) * .5 or hi + lo != hi:
        raise ValueError('normalized low component required')
    if exponent > 1024:
        raise OverflowError('scaled result outside binary64')
    if exponent >= -1021:
        # hi already is the rounded significand of the pair. Normal power-
        # of-two scaling preserves that rounding (including overflow at top).
        return math.ldexp(hi, exponent)
    if exponent < -1074:
        return math.copysign(0., hi)
    sign = -1. if hi < 0 else 1.
    units = math.ldexp(abs(hi), exponent + 1074)
    low_units = math.ldexp(sign * lo, exponent + 1074)
    nearest = round(units)
    remainder, residual = DD.sum(units - nearest, low_units)
    if (remainder > .5 or
        (remainder == .5 and (residual > 0 or
                             (residual == 0 and nearest % 2)))):
        nearest += 1
    elif (remainder < -.5 or
          (remainder == -.5 and (residual < 0 or
                                (residual == 0 and nearest % 2)))):
        nearest -= 1
    return math.copysign(math.ldexp(float(nearest), -1074), sign)
