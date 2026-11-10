"""
#307: the repaired binary64 Gamma route, modelled statement by statement.

The VBA port is the generated block of b64_route_vba.py, installed in
src/M_STATS_PROBDIST_SPECIALFUNCS.bas. Every arithmetic statement of the port
has its twin here. VBA evaluates an expression with 64-bit-significand
intermediates and rounds to a Double when the value is stored
(beta_balanced_survival/beta_route_model.py identified these semantics on Excel
observations); the helpers below do exactly that. Log, Exp and Sqr take and
return Doubles: math.log and math.exp stand in for VBA's, and Sqr is the square
root at 64 bits rounded to a Double.

The route answers P below X = A + 1 by the lower series and Q from A + 1 by the
continued fraction, both times the factor x^A e^-x / Gamma(A):

  Factor  for A > 1: sqrt(A / (2 pi)) exp(-(D + stirlerr(A))), where Loader's
          deviance D = A log(A/X) + X - A is an unevaluated pair from
          error-free transformations (accurate to about 1E-17 absolute; see
          accurate_deviance.py), and the exponent is kept as a pair:
          exp(Eh) (1 + El). For A <= 1: the small-shape factor of
          PROB_TryGammaContinuedFractionQ, A e^-X exp(A log X - log Gamma(1 + A)).
  Series  PROB_TryGammaSeriesP's recurrence with a compensated (Kahan) sum. It
          stops once the next term is below an ulp of the sum and the remaining
          terms, which fall at least geometrically by X / (Ap + 1), sum to at
          most half an ulp.
  CF      PROB_TryGammaContinuedFractionQ's modified Lentz recurrence, stopped
          at machine epsilon at every shape, with its backward re-evaluation at
          shape 1 and below.

The route declines (raises Declined; the public adapter then uses the
double-double kernel) where its arithmetic would leave the normal range: for
A > 1 when the binary exponent field of X / A is outside 1 .. 2000, and on any
runtime error, such as an overflow.

Production procedures it calls, PROB_StirlingError and PROB_GammaSmallLogGamma1p,
are bit-exact ports bound to the source text by test_route_model.py.
"""
import math
import os
import re
import sys

import mpmath as mp
from mpmath.libmp import from_float, mpf_add, mpf_div, mpf_mul, mpf_sqrt, mpf_sub, to_float

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(os.path.dirname(HERE), "beta_balanced_survival"))
import beta_route_model as BRM  # noqa: E402

PREC = 64
MACH_EPS = 2.0 ** -52
HALF_EPS = 2.0 ** -53
FPMIN = 1e-300
MAX_ITER = 100000
SPLITTER = 134217729.0          # 2^27 + 1
EXPONENT_FIELD = (1, 2000)      # X / A's biased exponent field, for A > 1


class Declined(ArithmeticError):
    """The route declines; the caller answers with the double-double kernel."""


# --------------------------------------------------------------------------
# one VBA statement: the operation(s) at 64 bits, then stored to a Double

def _f(v):
    return from_float(v)


def _d(m):
    return to_float(m, rnd="n")


def add(a, b):
    return _d(mpf_add(_f(a), _f(b), PREC, "n"))


def sub(a, b):
    return _d(mpf_sub(_f(a), _f(b), PREC, "n"))


def mul(a, b):
    return _d(mpf_mul(_f(a), _f(b), PREC, "n"))


def div(a, b):
    return _d(mpf_div(_f(a), _f(b), PREC, "n"))


def sqr(a):
    return _d(mpf_sqrt(_f(a), PREC, "n"))


def mul_div(a, b, c):
    """a * b / c in one statement."""
    return _d(mpf_div(mpf_mul(_f(a), _f(b), PREC, "n"), _f(c), PREC, "n"))


def mul_sub(a, b, c):
    """a * b - c in one statement."""
    return _d(mpf_sub(mpf_mul(_f(a), _f(b), PREC, "n"), _f(c), PREC, "n"))


def mul_add(a, b, c):
    """a * b + c in one statement."""
    return _d(mpf_add(mpf_mul(_f(a), _f(b), PREC, "n"), _f(c), PREC, "n"))


def add_sub(a, b, c):
    """a + b - c in one statement."""
    return _d(mpf_sub(mpf_add(_f(a), _f(b), PREC, "n"), _f(c), PREC, "n"))


def vba_exp(x):
    try:
        return math.exp(x)
    except OverflowError:
        raise Declined("Exp overflow") from None


# --------------------------------------------------------------------------
# error-free transformations (PROB_B64_TwoSum, _FastTwoSum, _TwoProd)

def two_sum(a, b):
    s = add(a, b)
    bb = sub(s, a)
    t = sub(s, bb)
    t = sub(a, t)
    u = sub(b, bb)
    return s, add(t, u)


def fast_two_sum(a, b):
    s = add(a, b)
    t = sub(s, a)
    return s, sub(b, t)


def split(a):
    c = mul(SPLITTER, a)
    t = sub(c, a)
    hi = sub(c, t)
    return hi, sub(a, hi)


def two_prod(a, b):
    p = mul(a, b)
    ah, al = split(a)
    bh, bl = split(b)
    e = mul(ah, bh)
    e = sub(e, p)
    t = mul(ah, bl)
    e = add(e, t)
    t = mul(al, bh)
    e = add(e, t)
    t = mul(al, bl)
    e = add(e, t)
    return p, e


def pair_add(ah, al, bh, bl):
    s, e = two_sum(ah, bh)
    t = add(al, bl)
    e = add(e, t)
    return fast_two_sum(s, e)


def pair_mul(ah, al, b):
    p, e = two_prod(ah, b)
    t = mul(al, b)
    e = add(e, t)
    return fast_two_sum(p, e)


def pair_div(ah, al, b):
    q = div(ah, b)
    p, e = two_prod(q, b)
    r = sub(ah, p)
    r = sub(r, e)
    r = add(r, al)
    r = div(r, b)
    return fast_two_sum(q, r)


# --------------------------------------------------------------------------
# constants: log(k / 256) for k = 128 .. 256, log 2 and 2 pi, as exact Doubles

def _pair(value):
    hi = float(value)
    return hi, float(value - hi)


with mp.workdps(60):
    LOG_TABLE = tuple(_pair(mp.log(mp.mpf(k) / 256)) for k in range(128, 257))
    LN2 = _pair(mp.log(2))
    TWO_PI = float(2 * mp.pi)
# log(1 + u) = u - u^2/2 + u^3 S(u), S(u) = sum (-u)^j / (j + 3), j = 0 .. 7;
# each coefficient is one division in VBA.
SERIES = tuple(div(-1.0 if j % 2 else 1.0, float(j + 3)) for j in range(8))


def exponent_field(y):
    """The biased exponent field of a positive Double (0 for a subnormal)."""
    if y < 2.0 ** -1022:
        return 0
    return math.frexp(y)[1] + 1022


def pair_log(yh, yl):
    """log(yh + yl) for a positive pair whose high part has exponent field F
    in 1 .. 2000: y = m 2^e, 0.5 <= m < 1, e = F - 1022; c = k / 256 nearest
    to m; log y = e log 2 + log c + log(1 + u), u = (m - c) / c."""
    field = exponent_field(yh)
    e = field - 1022
    scale = math.ldexp(1.0, 1022 - field)      # PROB_B64_Pow2(2045 - F): exact
    m = mul(yh, scale)                          # exact
    k = int(math.floor(m * 256.0 + 0.5))        # Int(M * 256# + 0.5): exact operands
    c = div(float(k), 256.0)                    # exact
    vh = sub(m, c)                              # exact (Sterbenz)
    vl = mul(yl, scale)                         # exact
    vh, vl = two_sum(vh, vl)
    uh, ul = pair_div(vh, vl, c)
    sh, sl = two_prod(uh, uh)
    t = mul(2.0, uh)
    t = mul(t, ul)
    sl = add(sl, t)
    sh = mul(sh, 0.5)
    sl = mul(sl, 0.5)
    poly = SERIES[7]
    for j in range(6, -1, -1):
        poly = mul(poly, uh)
        poly = add(poly, SERIES[j])
    cube = mul(uh, uh)
    cube = mul(cube, uh)
    cube = mul(cube, poly)
    lh, ll = pair_add(uh, ul, -sh, -sl)
    lh, ll = pair_add(lh, ll, cube, 0.0)
    th, tl = LOG_TABLE[k - 128]
    lh, ll = pair_add(lh, ll, th, tl)
    eh, el = pair_mul(LN2[0], LN2[1], float(e))
    return pair_add(lh, ll, eh, el)


def deviance(a, x):
    """D = A log(A/X) + X - A = A (t - log(1 + t)), t = X/A - 1, as a pair."""
    q = div(x, a)
    field = exponent_field(q)
    if not EXPONENT_FIELD[0] <= field <= EXPONENT_FIELD[1]:
        raise Declined("X / A outside the normal range of the deviance")
    yh, yl = pair_div(x, 0.0, a)
    th, tl = pair_add(yh, yl, -1.0, 0.0)
    lh, ll = pair_log(yh, yl)
    gh, gl = pair_add(th, tl, -lh, -ll)
    return pair_mul(gh, gl, a)


# --------------------------------------------------------------------------
# production helpers

def _small_log_gamma1p_coefficients():
    text = BRM.procedure_text("special", "PROB_GammaSmallLogGamma1p")
    out = {}
    for m in re.finditer(r"Const C(\d+) As Double = ([^'\n]+)", text):
        parts = re.fullmatch(r"\s*([-+\dE.]+)\s*\+\s*\(([-+\dE.]+)\)\s*", m.group(2)).groups()
        out[int(m.group(1))] = add(float(parts[0]), float(parts[1]))
    if sorted(out) != list(range(33)):
        raise RuntimeError("PROB_GammaSmallLogGamma1p coefficients not found as expected")
    return tuple(out[k] for k in range(33))


SMALL_C = _small_log_gamma1p_coefficients()


def small_log_gamma1p(a):
    """PROB_GammaSmallLogGamma1p: log Gamma(1 + A) for A <= 1."""
    u = sub(a, 0.5)
    acc = SMALL_C[32]
    for k in range(31, -1, -1):
        acc = mul_add(acc, u, SMALL_C[k])
    product = _d(mpf_mul(_f(a), mpf_sub(_f(a), _f(1.0), PREC, "n"), PREC, "n"))
    return mul(product, acc)


# --------------------------------------------------------------------------
# the route (PROB_B64_TryFactor, PROB_TryGammaB64SeriesP,
# PROB_TryGammaB64ContinuedFractionQ)

def factor(a, x):
    """x^A e^-x / Gamma(A) for X > 0."""
    if a <= 1.0:
        if x >= 708.0:
            return 0.0
        f = vba_exp(-x)
        lx = math.log(x)
        ex = mul_sub(a, lx, small_log_gamma1p(a))
        corr = vba_exp(ex)
        f = mul(a, f)
        return mul(f, corr)
    dh, dl = deviance(a, x)
    stir = BRM.stirling_error(a)
    eh, el = two_sum(-dh, -stir)
    if eh < -708.0:
        return 0.0
    el = sub(el, dl)
    root = div(a, TWO_PI)
    root = sqr(root)
    f = vba_exp(eh)
    f = mul(root, f)
    t = mul(f, el)
    return add(f, t)


def series_p(a, x):
    """P(A, X) for 0 < X < A + 1: (value, iterations), or Declined."""
    ap = a
    total = div(1.0, a)
    term = total
    carry = 0.0
    for i in range(1, MAX_ITER + 1):
        ap = add(ap, 1.0)
        term = mul_div(term, x, ap)
        s = add(total, term)
        t = sub(s, total)
        t = sub(term, t)
        carry = add(carry, t)
        total = s
        if abs(term) <= abs(total) * MACH_EPS:
            room = add(ap, 1.0)
            room = sub(room, x)
            if room > 0.0:
                bound = mul(term, x)
                limit = mul(total, HALF_EPS)
                limit = mul(limit, room)
                if bound <= limit:
                    total = add(total, carry)
                    return mul(total, factor(a, x)), i
    raise Declined("series cap")


def cf_q(a, x):
    """Q(A, X) for X >= A + 1: (value, iterations), or Declined."""
    b = add_sub(x, 1.0, a)
    if abs(b) < FPMIN:
        b = FPMIN
    c = div(1.0, FPMIN)
    d = div(1.0, b)
    h = d
    for i in range(1, MAX_ITER + 1):
        fi = float(i)
        an = _d(mpf_mul(_f(-fi), mpf_sub(_f(fi), _f(a), PREC, "n"), PREC, "n"))
        b = add(b, 2.0)
        d = mul_add(an, d, b)
        if abs(d) < FPMIN:
            d = FPMIN
        c = _d(mpf_add(_f(b), mpf_div(_f(an), _f(c), PREC, "n"), PREC, "n"))
        if abs(c) < FPMIN:
            c = FPMIN
        d = div(1.0, d)
        delta = mul(d, c)
        h = mul(h, delta)
        if abs(delta - 1.0) <= MACH_EPS:
            if a <= 1.0:
                base = add_sub(x, 1.0, a)
                tail = 0.0
                for k in range(i, 0, -1):
                    fk = float(k)
                    num = mpf_mul(_f(-fk), mpf_sub(_f(fk), _f(a), PREC, "n"), PREC, "n")
                    den = mpf_add(mpf_add(_f(base), mpf_mul(_f(2.0), _f(fk), PREC, "n"), PREC, "n"),
                                  _f(tail), PREC, "n")
                    tail = _d(mpf_div(num, den, PREC, "n"))
                h = div(1.0, add(base, tail))
            return mul(factor(a, x), h), i
    raise Declined("continued fraction cap")


def below_seam(a, x):
    """VBA's X < A + 1#, as Excel evaluates it: A + 1 is rounded to a Double
    before the comparison. The fit wave (run 42798193976) showed it at
    X = fl(A + 1) for A = 1E-6, 1E-4 and 1E-3, where the exact A + 1 lies
    above X and a 64-bit intermediate would have chosen the series."""
    return x < a + 1.0


def route(a, x):
    """(routine, value, iterations) for X > 0: the series below X = A + 1, the
    continued fraction from it; value and iterations are None where the route
    declines."""
    routine = "series" if below_seam(a, x) else "cf"
    try:
        v, n = series_p(a, x) if routine == "series" else cf_q(a, x)
        return routine, v, n
    except Declined:
        return routine, None, None
