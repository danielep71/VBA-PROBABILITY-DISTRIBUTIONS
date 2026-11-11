"""
#307 prototype: Loader's deviance D = A log(A/X) + X - A to about double-double
precision, from error-free binary64 transformations only.

  y = X / A          as an unevaluated pair (yh, yl), exact to about 2^-106
  t = y - 1          the same pair, shifted
  log y              e log 2 + log c + log1p(u): y = 2^e m, c the nearest
                     multiple of 1/256 to m, log c from a two-part table, and
                     u = (m - c) / c, with |u| < 2^-8, through a short series
  D = A (t - log y)  as a pair; the result is (Dh, Dl)

Every operation is a binary64 operation in a fixed order, so a VBA port with
the same order returns the same bits. No double-double library, no loop beyond
the eight-term series. Exploratory: its accuracy is measured here; it is not a
production kernel and nothing routes to it.

Run: python3 accurate_deviance.py   (prints the measured accuracy)
"""
import math

import mpmath as mp

SPLITTER = 134217729.0          # 2^27 + 1


def two_sum(a, b):
    s = a + b
    bb = s - a
    return s, (a - (s - bb)) + (b - bb)


def fast_two_sum(a, b):
    s = a + b
    return s, b - (s - a)


def split(a):
    c = SPLITTER * a
    hi = c - (c - a)
    return hi, a - hi


def two_prod(a, b):
    p = a * b
    ah, al = split(a)
    bh, bl = split(b)
    return p, ((ah * bh - p) + ah * bl + al * bh) + al * bl


def dd_add(ah, al, bh, bl):
    s, e = two_sum(ah, bh)
    e += al + bl
    return fast_two_sum(s, e)


def dd_mul_d(ah, al, b):
    p, e = two_prod(ah, b)
    e += al * b
    return fast_two_sum(p, e)


def dd_div_d(ah, al, b):
    """(ah + al) / b for a binary64 b, as a pair."""
    q = ah / b
    p, e = two_prod(q, b)
    r = ((ah - p) - e + al) / b
    return fast_two_sum(q, r)


def _table():
    with mp.workdps(60):
        out = []
        for k in range(128, 257):
            v = mp.log(mp.mpf(k) / 256)
            hi = float(v)
            out.append((hi, float(v - hi)))
        ln2 = mp.log(2)
        return out, (float(ln2), float(ln2 - float(ln2)))


LOG_TABLE, LN2 = _table()
# log1p(u) = u - u^2/2 + u^3 S(u), S(u) = 1/3 - u/4 + u^2/5 - ... (8 terms; |u| < 2^-8)
SERIES = tuple(((-1) ** k) / (k + 3) for k in range(8))


def dd_log(yh, yl):
    """log(yh + yl) for a positive pair, as a pair."""
    m, e = math.frexp(yh)                 # yh = m 2^e, 0.5 <= m < 1
    k = int(round(m * 256.0))
    c = k / 256.0
    scale = math.ldexp(1.0, -e)
    vh, vl = two_sum(yh * scale - c, yl * scale)   # yh*scale - c is exact (Sterbenz)
    uh, ul = dd_div_d(vh, vl, c)
    # u^2 / 2 as a pair, u^3 S(u) in binary64
    sh, sl = two_prod(uh, uh)
    sl += 2.0 * uh * ul
    sh, sl = sh * 0.5, sl * 0.5
    poly = 0.0
    for coefficient in reversed(SERIES):
        poly = poly * uh + coefficient
    cube = uh * uh * uh * poly
    lh, ll = dd_add(uh, ul, -sh, -sl)
    lh, ll = dd_add(lh, ll, cube, 0.0)
    th, tl = LOG_TABLE[k - 128]
    lh, ll = dd_add(lh, ll, th, tl)
    eh, el = dd_mul_d(LN2[0], LN2[1], float(e))
    return dd_add(lh, ll, eh, el)


def deviance(a, x):
    """D = A log(A/X) + X - A = A (t - log1p(t)), t = X/A - 1, as a pair."""
    yh, yl = dd_div_d(x, 0.0, a)
    th, tl = dd_add(yh, yl, -1.0, 0.0)
    lh, ll = dd_log(yh, yl)
    gh, gl = dd_add(th, tl, -lh, -ll)
    return dd_mul_d(gh, gl, a)


def main():
    import random
    rng = random.Random(307)
    worst_abs = worst_rel = 0.0
    with mp.workdps(80):
        for _ in range(20000):
            a = 10 ** rng.uniform(-6, 7)
            x = a * math.exp(rng.uniform(-3, 3)) if rng.random() < 0.7 else 10 ** rng.uniform(-6, 7)
            dh, dl = deviance(a, x)
            exact = mp.mpf(a) * mp.log(mp.mpf(a) / mp.mpf(x)) + mp.mpf(x) - mp.mpf(a)
            err = abs(mp.mpf(dh) + mp.mpf(dl) - exact)
            worst_abs = max(worst_abs, float(err))
            if exact > 0:
                worst_rel = max(worst_rel, float(err / exact))
    print(f"20000 random (A, X): worst absolute error of D {worst_abs:.2e}, "
          f"worst relative {worst_rel:.2e}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
