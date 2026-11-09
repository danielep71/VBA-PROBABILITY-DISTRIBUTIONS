"""Single-round an exact finite mpmath dyadic to binary64 using integers.

This is validation tooling, not numerical runtime code. mpmath's ordinary
float conversion may round the significand and then scale into subnormal range.
Retain all mantissa bits until rounding to the final binary64 lattice instead.
"""
import math


def rounded(value):
    sign,mantissa,exponent,bits=value._mpf_
    if bits<0:raise ValueError('finite reference required')
    if mantissa==0:return -0. if sign else 0.
    power=exponent+bits
    if power < -1074:return -0. if sign else 0.
    if power > 1024:return -math.inf if sign else math.inf
    quantum=max(power-53,-1074)
    shift=exponent-quantum
    if shift>=0:
        nearest=mantissa<<shift
    else:
        denominator=1<<(-shift)
        nearest,remainder=divmod(mantissa,denominator)
        twice=remainder*2
        if twice>denominator or (twice==denominator and nearest%2):nearest+=1
    try:answer=math.ldexp(float(nearest),quantum)
    except OverflowError:answer=math.inf
    return -answer if sign else answer
