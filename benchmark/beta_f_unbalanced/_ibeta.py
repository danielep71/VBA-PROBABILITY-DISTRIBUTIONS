"""Robust regularized incomplete beta (continued fraction) for high-precision refs."""
import mpmath as mp

def betacf(a, b, x):
    tiny = mp.mpf('1e-300'); qab = a + b; qap = a + 1; qam = a - 1
    c = mp.mpf(1); d = 1 - qab * x / qap
    if abs(d) < tiny: d = tiny
    d = 1 / d; h = d
    for m in range(1, 20000):
        m2 = 2 * m
        aa = m * (b - m) * x / ((qam + m2) * (a + m2))
        d = 1 + aa * d
        if abs(d) < tiny: d = tiny
        c = 1 + aa / c
        if abs(c) < tiny: c = tiny
        d = 1 / d; h *= d * c
        aa = -(a + m) * (qab + m) * x / ((a + m2) * (qap + m2))
        d = 1 + aa * d
        if abs(d) < tiny: d = tiny
        c = 1 + aa / c
        if abs(c) < tiny: c = tiny
        d = 1 / d; de = d * c; h *= de
        if abs(de - 1) < mp.mpf('1e-45'): break
    return h

def ibeta(x, a, b):
    x, a, b = mp.mpf(x), mp.mpf(a), mp.mpf(b)
    if x <= 0: return mp.mpf(0)
    if x >= 1: return mp.mpf(1)
    lbt = mp.loggamma(a + b) - mp.loggamma(a) - mp.loggamma(b) + a * mp.log(x) + b * mp.log(1 - x)
    bt = mp.e ** lbt
    if x < (a + 1) / (a + b + 2): return bt * betacf(a, b, x) / a
    return 1 - bt * betacf(b, a, 1 - x) / b

def beta_invcdf(p, a, b):
    p, a, b = mp.mpf(p), mp.mpf(a), mp.mpf(b)
    lo, hi = mp.mpf(0), mp.mpf(1)
    for _ in range(230):
        mid = (lo + hi) / 2
        if ibeta(mid, a, b) < p: lo = mid
        else: hi = mid
    return (lo + hi) / 2

def f_cdf(x, d1, d2):
    d1, d2, x = mp.mpf(d1), mp.mpf(d2), mp.mpf(x)
    return ibeta(d1 * x / (d1 * x + d2), d1 / 2, d2 / 2)

def f_invcdf(p, d1, d2):
    lo, hi = mp.mpf(0), mp.mpf('1e12')
    for _ in range(270):
        mid = (lo + hi) / 2
        if f_cdf(mid, d1, d2) < mp.mpf(p): lo = mid
        else: hi = mid
    return (lo + hi) / 2
