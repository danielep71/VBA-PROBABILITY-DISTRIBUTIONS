"""Separate wide Gamma candidate; development only, with explicit cap failures.

The verified extended-v2 interior stays frozen. Outside it, a stable deviance
and ten-term Stirling correction avoid an arbitrary shape cutoff. A geometric
lower-series remainder replaces a last-term-only screen for larger shapes.
Arithmetic and CF stopping still require validation; this is not a global
rounding certificate or production dispatch.
"""
import math

from scaled_dd import ScaledDD as S
from scaled_gamma import ONE, small, log_scaled, exp_scaled
from extended_gamma import log_gamma1p, log_pair, lower_pair
from extended_gamma_v2 import public as interior


STIRLING = ((1,12),(-1,360),(1,1260),(-1,1680),(1,1188),
            (-691,360360),(1,156),(-3617,122400),(43867,244188),
            (-174611,125400))
# Nearest binary64 plus nearest residual; independently regenerated offline.
LOG_SQRT_2PI = S.from_float(float.fromhex('0x1.d67f1c864beb5p-1')) + S.from_float(
    float.fromhex('-0x1.65b5a1b7ff5dfp-55'))
MAX_ITER = 100000
LOWER_REMAINDER_BITS = 68


def stirling(a):
    inverse = ONE/a
    square = inverse*inverse
    term = S.from_float(STIRLING[-1][0])/STIRLING[-1][1]
    for num,den in reversed(STIRLING[:-1]):
        term = term*square+S.from_float(num)/den
    return term*inverse


def deviance(a,x):
    """x-a-a*log(x/a), without subtracting first-order terms near x=a."""
    aa,xx = S.from_float(a),S.from_float(x)
    u = (xx-aa)/aa
    if abs(u.rounded()) <= .5:
        term = u*u
        total = term/2
        for n in range(3,1001):
            term = term*(-u)
            delta = term/n
            total = total+delta
            if small(delta,total):
                value = aa*total
                break
        else:
            raise RuntimeError('deviance series cap')
    else:
        value = xx-aa-aa*(log_scaled(x)-log_scaled(a))
    if value.hi < 0:
        raise ValueError('negative deviance')
    return value


def upper(a,x,log_prefactor,trace):
    aa,xx = S.from_float(a),S.from_float(x)
    b = xx+1-aa
    c,d = S.from_float(1e300),ONE/b
    for n in range(1,MAX_ITER+1):
        an = S.from_float(-n)*(S.from_float(n)-aa)
        b = b+2
        d = ONE/(an*d+b)
        c = b+an/c
        delta = c*d-1
        if delta.hi==0 or delta.exponent < -100:
            break
    else:
        raise RuntimeError('continued fraction failed to converge in 100000 iterations')
    tail = S.from_float(0.)
    for k in range(n,0,-1):
        tail = S.from_float(-k)*(S.from_float(k)-aa)/(xx+1-aa+2*k+tail)
    h = ONE/(xx+1-aa+tail)
    q = exp_scaled(log_prefactor)*aa*h
    trace.update(route='upper',iterations=n)
    return (ONE-q).rounded(),q.rounded()


def public(a,x,trace=None):
    if trace is None:
        trace = {}
    if not (math.isfinite(a) and math.isfinite(x) and a>0 and x>=0):
        raise ValueError('finite positive shape and nonnegative argument required')
    if x==0:
        trace['route']='endpoint'
        return 0.,1.
    if a<=128 and x<=1024:
        trace['route']='frozen extended-v2 interior'
        return interior(a,x)
    aa,xx = S.from_float(a),S.from_float(x)
    if a<=128:
        # Chernoff: for a<=128 and x>=1280, log Q <= -832 < -1075*ln(2).
        if x>=1280:
            trace['route']='small-shape Chernoff underflow'
            return 1.,0.
        lp = log_scaled(x)*aa-log_gamma1p(aa)-xx
        return upper(a,x,lp,trace)
    dev = deviance(a,x)
    # Do not round a potentially unrepresentable intermediate to binary64.
    if dev.exponent>=11 or dev.rounded()>750:
        trace['route']='deviance underflow'
        return (0.,1.) if x<a else (1.,0.)
    lp = -dev-log_scaled(a)*.5-LOG_SQRT_2PI-stirling(aa)
    if x>=a+1:
        return upper(a,x,lp,trace)
    term,total = ONE,ONE
    for n in range(1,MAX_ITER+1):
        term = term*xx/(aa+n)
        total = total+term
        ratio = xx/(aa+n+1)
        gap = ONE-ratio
        if gap.hi>0:
            remainder = term*ratio/gap
            guarded = remainder*4
            if guarded.hi==0 or guarded.exponent<total.exponent-LOWER_REMAINDER_BITS:
                break
    else:
        raise RuntimeError('lower series failed to converge in 100000 iterations')
    p,q = lower_pair(lp+log_pair(total))
    trace.update(route='positive lower with geometric remainder',iterations=n,
                 remainder_exponent=remainder.exponent,total_exponent=total.exponent)
    return p.rounded(),q.rounded()
