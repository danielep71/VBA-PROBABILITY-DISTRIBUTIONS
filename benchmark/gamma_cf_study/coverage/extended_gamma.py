"""Separate development candidate: 0<a<=128, 0<x<=1024; no public dispatch.

Positive lower series for a>1 avoids the cancellation of an alternating
expansion at large x. LogGamma1p uses recurrence from the frozen polynomial.
Small and large Q use direct expm1 with argument halving and doubling.
Stopping screens remain heuristic; this is not a general error certificate.
"""
import math
from scaled_dd import ScaledDD as S
from scaled_gamma import ONE, LN_TWO, small, log_scaled, exp_scaled, expm1_scaled, log1p_scaled
from precision_v2 import COEFFICIENTS


def log_pair(value):
    if value.hi <= 0:
        raise ValueError('positive scaled logarithm required')
    f = S(value.hi,value.lo,0)
    z = (f-1)/(f+1)
    square = z*z
    term,total = z,z
    for n in range(3,1001,2):
        term = term*square
        delta = term/n
        total = total+delta
        if small(delta,total):
            return total*2+LN_TWO*value.exponent
    raise RuntimeError('log pair cap')


def log_gamma1p(a):
    n = int(a.rounded()) if a.rounded()>1 else 0
    b = a-n
    u = b-.5
    lg = S._normalize(COEFFICIENTS[-1],0)
    for c in reversed(COEFFICIENTS[:-1]):
        lg = lg*u+S._normalize(c,0)
    lg = b*(b-1)*lg
    for k in range(1,n+1):
        lg = lg+log_pair(b+k)
    return lg


def lower_pair(lp):
    value = lp.rounded()
    if lp.hi>0:
        raise ValueError('positive log probability')
    # This final-probability shortcut is far below log(2^-1075), not an
    # exponential helper range limit. It still needs a full error budget.
    if value < -750:
        return S.from_float(0.),ONE
    halvings = 0
    z = lp
    while z.rounded() < -.5:
        z = z*.5
        halvings += 1
    q = -expm1_scaled(z)
    for _ in range(halvings):
        q = q*(ONE*2-q)
    p = ONE-q if value >= -.5 else exp_scaled(lp)
    return p,q


def public(a,x):
    if not (math.isfinite(a) and math.isfinite(x) and 0<a<=128 and 0<x<=1024):
        raise ValueError('outside extended study domain')
    aa,xx = S.from_float(a),S.from_float(x)
    lg = log_gamma1p(aa)
    if x<a+1:
        if a<=1:
            term = -xx
            total = term/(aa+1)
            for n in range(2,1001):
                term = term*(-xx)/n
                delta = term/(aa+n)
                total = total+delta
                if small(delta,total):
                    break
            else:
                raise RuntimeError('alternating lower cap')
            lp = log_scaled(x)*aa-lg+log1p_scaled(total*aa)
        else:
            term,total = ONE,ONE
            for n in range(1,10001):
                term = term*xx/(aa+n)
                total = total+term
                if small(term,total):
                    break
            else:
                raise RuntimeError('positive lower cap')
            lp = log_scaled(x)*aa-xx-lg+log_pair(total)
        p,q = lower_pair(lp)
    else:
        b = xx+1-aa
        c,d = S.from_float(1e300),ONE/b
        for n in range(1,100001):
            an = S.from_float(-n)*(S.from_float(n)-aa)
            b = b+2
            d = ONE/(an*d+b)
            c = b+an/c
            delta = c*d-1
            if delta.hi==0 or delta.exponent < -100:
                break
        else:
            raise RuntimeError('upper cap')
        tail = S.from_float(0.)
        for k in range(n,0,-1):
            tail = S.from_float(-k)*(S.from_float(k)-aa)/(xx+1-aa+2*k+tail)
        h = ONE/(xx+1-aa+tail)
        q = exp_scaled(log_scaled(x)*aa-lg-xx)*aa*h
        p = ONE-q
    return p.rounded(),q.rounded()
