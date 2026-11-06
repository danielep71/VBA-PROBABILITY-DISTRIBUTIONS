"""
Direct upper-tail survival oracle at the exact binary64 argument (#159).

Every committed survival reference must equal the survival function evaluated at
the Double the exporter passes to the library, to at least 25 significant
digits (the SER_FLOOR of audit_references.py and audit_study_references.py).
This module recomputes that value independently of every grid generator:

  * Arguments are exact. Each token is parsed to its Double and promoted with
    _binary64.exact_binary64_mpf, so no decimal reading of the token enters.
  * Every route computes the upper tail directly. None forms 1 - CDF where the
    survival is small; a complement appears only where the survival is at least
    about one half, so the subtraction cannot cancel.
  * mpmath.betainc(a, b, x, 1) is never used. It is itself a complement of the
    lower tail, so it cannot judge a reference in the upper tail.

Routes, by function:

  Beta_Survival(x, a, b)      I_{1-x}(b, a), by mpmath.betainc(b, a, 0, 1 - x)
                              while either shape is below QUAD_MIN_SHAPE, else
                              by quadrature of the Beta(b, a) density on
                              [0, 1 - x].
  F_Survival(x, d1, d2)       the same with 1 - y = d2 / (d1 x + d2) formed
                              directly, never as 1 - y.
  StudentT_Survival(t, df)    I_z(df/2, 1/2) / 2 with z = df / (df + t^2), for
                              t > 0; 1 minus that for t < 0, where it is at least
                              one half; 1/2 at t = 0.
  ChiSquare_Survival(x, df)   regularized upper incomplete gamma Q(df/2, x/2):
  Gamma_Survival(x, k, s)     Q(k, x/s); mpmath.gammainc(a, z, inf) up to
                              GAMMA_DIRECT_MAX_SHAPE, quadrature of the gamma
                              density on [z, inf) above it.
  Normal / NormalStandard /   erfc(z / sqrt 2) / 2 at the exact standardised
  Lognormal_Survival          argument.
  Exponential, Weibull,       closed forms exp(-l x), exp(-(x/l)^k) and
  Uniform_Survival            (b - x) / (b - a).
  Binomial, Poisson,          exact sums of the probability mass above k, built
  NegativeBinomial_Survival   by the ratio recurrence; below the mode the sum
                              is taken over the mass at or below k and
                              subtracted from 1, where the survival exceeds
                              about one half.
  Hypergeometric_Survival     exact finite sum of the mass above k.
  Geometric_Survival          (1 - p)^(floor(k) + 1).
  DiscreteUniform_Survival    (hi - floor(k)) / (hi - lo + 1).

Each value is computed at two working precisions. Their agreement is the
oracle's own uncertainty; certify() refuses a value whose two evaluations agree
to fewer than CERTIFY_DIGITS significant digits, so an unconverged oracle fails
closed instead of passing or failing a reference by accident. The working
precision rises with the magnitude of the shapes, because the log-density of a
shape-A distribution subtracts terms of size A log A.

The quadrature and the closed-form routes are cross-checked against a Lentz
continued fraction and against Temme's uniform expansion where both shapes are
huge (test_survival_references.py).
"""
import math

import mpmath as mp

from _binary64 import exact_binary64_mpf

SER_FLOOR = mp.mpf(10) ** -24      # 25 significant digits, as in the audits
CERTIFY_DIGITS = 32                # the two precisions must agree this far
DPS_LOW, DPS_HIGH = 45, 65         # before the shape-magnitude allowance
QUAD_MIN_SHAPE = 20                # both beta shapes at least this: quadrature
GAMMA_DIRECT_MAX_SHAPE = 1e9       # above this, gamma by quadrature


class OracleError(ValueError):
    """The oracle cannot produce a certified value for this row."""


# --------------------------------------------------------------------------
# argument handling


def exact_args(tokens):
    """Exact binary64 value of each non-blank token, in order."""
    out = []
    for t in tokens:
        t = (t or "").strip()
        if t == "":
            continue
        v = float(t)
        if not math.isfinite(v):
            raise OracleError(f"argument {t!r} is not a finite Double")
        out.append(exact_binary64_mpf(t))
    return out


def magnitude_digits(args):
    """Extra working digits needed for shapes of this size."""
    big = max([abs(float(a)) for a in args] + [1.0])
    return int(math.ceil(math.log10(big))) + 5


# --------------------------------------------------------------------------
# continuous kernels


def _segments(lo, hi, centre, sd, dps):
    """Breakpoints one standard deviation apart across the mass in [lo, hi].

    Beyond K standard deviations a near-Gaussian density is below
    10^-(dps + 10) of its peak, so the mass there cannot reach a retained
    digit. The interval outside that band is still integrated, as one
    segment, so no mass is dropped by construction."""
    k = int(math.ceil(math.sqrt(2 * (dps + 10) * math.log(10)))) + 2
    first = max(lo, centre - k * sd)
    last = centre + k * sd if hi == mp.inf else min(hi, centre + k * sd)
    pts = [lo]
    if first > lo:
        pts.append(first)
    p = first
    while True:
        p = p + sd
        if p >= last:
            break
        if p > pts[-1]:
            pts.append(p)
    if hi == mp.inf or last < hi:
        if last > pts[-1]:
            pts.append(last)
    pts.append(hi)
    return pts


def beta_lower_quad(t, a, b):
    """I_t(a, b) by quadrature of the Beta(a, b) density on [0, t]."""
    if t <= 0:
        return mp.mpf(0)
    if t >= 1:
        return mp.mpf(1)
    lb = mp.loggamma(a) + mp.loggamma(b) - mp.loggamma(a + b)
    mode = (a - 1) / (a + b - 2)
    sd = mp.sqrt(a * b / ((a + b) ** 2 * (a + b + 1)))

    def f(u):
        return mp.e ** ((a - 1) * mp.log(u) + (b - 1) * mp.log1p(-u) - lb)

    pts = _segments(mp.mpf(0), t, mode, sd, mp.mp.dps)
    return mp.quad(f, pts)


def beta_upper(xc, a, b):
    """P(Beta(a, b) > x) given xc = 1 - x exactly: I_xc(b, a)."""
    if xc <= 0:
        return mp.mpf(0)
    if xc >= 1:
        return mp.mpf(1)
    if min(a, b) >= QUAD_MIN_SHAPE:
        return beta_lower_quad(xc, b, a)
    return mp.betainc(b, a, 0, xc, regularized=True)


def gamma_upper_quad(a, z):
    """Q(a, z) by quadrature of the gamma density on [z, inf)."""
    lg = mp.loggamma(a)

    def f(t):
        return mp.e ** ((a - 1) * mp.log(t) - t - lg)

    pts = _segments(z, mp.inf, a - 1, mp.sqrt(a), mp.mp.dps)
    return mp.quad(f, pts)


def gamma_upper(a, z):
    """Regularized upper incomplete gamma Q(a, z)."""
    if z <= 0:
        return mp.mpf(1)
    if a > GAMMA_DIRECT_MAX_SHAPE:
        return gamma_upper_quad(a, z)
    return mp.gammainc(a, z, mp.inf, regularized=True)


def normal_upper(z):
    return mp.erfc(z / mp.sqrt(2)) / 2


# --------------------------------------------------------------------------
# discrete kernels


def _mass_tail(k, log_pmf, ratio, mode, upper_end):
    """P(X > k) for an integer-valued X by exact summation.

    log_pmf(j) gives the log mass at j, ratio(j) gives pmf(j+1)/pmf(j), and
    upper_end is the last support point (mp.inf for an unbounded support).
    Above the mode the upper tail is summed outward until a term falls below
    10^-(dps + 10) of the sum. At or below the mode the survival is at least
    about one half, so the shorter lower part is summed and subtracted from 1;
    that complement cannot cancel."""
    k = int(mp.floor(k))
    if k < 0:
        return mp.mpf(1)
    if upper_end != mp.inf and k >= upper_end:
        return mp.mpf(0)
    eps = mp.mpf(10) ** -(mp.mp.dps + 10)
    if k + 1 > mode:
        j = k + 1
        term = mp.e ** log_pmf(j)
        total = mp.mpf(0)
        while True:
            total += term
            if upper_end != mp.inf and j >= upper_end:
                break
            if term <= total * eps:
                break
            term = term * ratio(j)
            j += 1
        return total
    j = k
    term = mp.e ** log_pmf(j)
    total = mp.mpf(0)
    while True:
        total += term
        if j <= 0 or term <= total * eps:
            break
        term = term / ratio(j - 1)
        j -= 1
    return 1 - total


def binomial_upper(k, n, p):
    n_int = int(n)
    q = 1 - p
    lp, lq = mp.log(p), mp.log1p(-p)

    def log_pmf(j):
        return (mp.loggamma(n + 1) - mp.loggamma(j + 1) - mp.loggamma(n - j + 1)
                + j * lp + (n - j) * lq)

    def ratio(j):
        return (n - j) / mp.mpf(j + 1) * p / q

    return _mass_tail(k, log_pmf, ratio, int(mp.floor((n + 1) * p)), n_int)


def poisson_upper(k, lam):
    ll = mp.log(lam)

    def log_pmf(j):
        return j * ll - lam - mp.loggamma(j + 1)

    def ratio(j):
        return lam / mp.mpf(j + 1)

    return _mass_tail(k, log_pmf, ratio, int(mp.floor(lam)), mp.inf)


def negbinomial_upper(k, r, p):
    lp, lq = mp.log(p), mp.log1p(-p)
    q = 1 - p

    def log_pmf(j):
        return (mp.loggamma(j + r) - mp.loggamma(j + 1) - mp.loggamma(r)
                + r * lp + j * lq)

    def ratio(j):
        return (j + r) / mp.mpf(j + 1) * q

    mode = int(mp.floor((r - 1) * q / p)) if r > 1 else 0
    return _mass_tail(k, log_pmf, ratio, mode, mp.inf)


def hypergeometric_upper(k, n, big_k, big_n):
    k, n, big_k, big_n = (int(mp.floor(v)) for v in (k, n, big_k, big_n))
    hi = min(n, big_k)
    total = mp.mpf(0)
    for j in range(max(k + 1, 0), hi + 1):
        total += mp.binomial(big_k, j) * mp.binomial(big_n - big_k, n - j)
    return total / mp.binomial(big_n, n)


def discrete_uniform_upper(x, lo, hi):
    if x < lo:
        return mp.mpf(1)
    if x >= hi:
        return mp.mpf(0)
    return (hi - mp.floor(x)) / (hi - lo + 1)


# --------------------------------------------------------------------------
# public functions


def _student_t(t, df):
    if t == 0:
        return mp.mpf(1) / 2
    h = mp.betainc(df / 2, mp.mpf(1) / 2, 0, df / (df + t * t), regularized=True) / 2
    return h if t > 0 else 1 - h


def _f(x, d1, d2):
    if x <= 0:
        return mp.mpf(1)
    return beta_upper(d2 / (d1 * x + d2), d1 / 2, d2 / 2)


def _beta(x, a, b):
    if x <= 0:
        return mp.mpf(1)
    if x >= 1:
        return mp.mpf(0)
    return beta_upper(1 - x, a, b)


def _uniform(x, a, b):
    if x <= a:
        return mp.mpf(1)
    if x >= b:
        return mp.mpf(0)
    return (b - x) / (b - a)


def _lognormal(x, m, s):
    if x <= 0:
        return mp.mpf(1)
    return normal_upper((mp.log(x) - m) / s)


ROUTES = {
    # function: (arity, route name, evaluator)
    "Beta_Survival": (3, "I_{1-x}(b,a): betainc(b,a,0,1-x) or quadrature", _beta),
    "F_Survival": (3, "I_{d2/(d1 x+d2)}(d2/2,d1/2): betainc or quadrature", _f),
    "StudentT_Survival": (2, "betainc(df/2,1/2,0,df/(df+t^2))/2", _student_t),
    "ChiSquare_Survival": (2, "upper incomplete gamma Q(df/2,x/2)",
                           lambda x, df: gamma_upper(df / 2, x / 2)),
    "Gamma_Survival": (3, "upper incomplete gamma Q(k,x/s)",
                       lambda x, k, s: gamma_upper(k, x / s)),
    "NormalStandard_Survival": (1, "erfc", normal_upper),
    "Normal_Survival": (3, "erfc", lambda x, m, s: normal_upper((x - m) / s)),
    "Lognormal_Survival": (3, "erfc", _lognormal),
    "Exponential_Survival": (2, "exp(-l x)",
                             lambda x, lam: mp.mpf(1) if x <= 0 else mp.e ** (-lam * x)),
    "Weibull_Survival": (3, "exp(-(x/l)^k)",
                         lambda x, k, lam: mp.mpf(1) if x <= 0 else mp.e ** (-(x / lam) ** k)),
    "Uniform_Survival": (3, "(b-x)/(b-a)", _uniform),
    "Binomial_Survival": (3, "exact mass sum", binomial_upper),
    "Poisson_Survival": (2, "exact mass sum", poisson_upper),
    "NegativeBinomial_Survival": (3, "exact mass sum", negbinomial_upper),
    "Hypergeometric_Survival": (4, "exact finite sum", hypergeometric_upper),
    "Geometric_Survival": (2, "(1-p)^(floor(k)+1)",
                           lambda k, p: mp.mpf(1) if k < 0 else (1 - p) ** (mp.floor(k) + 1)),
    "DiscreteUniform_Survival": (3, "(hi-floor(k))/(hi-lo+1)", discrete_uniform_upper),
}


def is_survival(function):
    return function in ROUTES


def evaluate(function, args, dps):
    """The survival at exact arguments, at one working precision."""
    if function not in ROUTES:
        raise OracleError(f"no direct route for {function}")
    arity, _, fn = ROUTES[function]
    if len(args) < arity:
        raise OracleError(f"{function} needs {arity} arguments, got {len(args)}")
    with mp.workdps(dps):
        a = [+v for v in args[:arity]]
        v = fn(*a)
        if not mp.isfinite(v) or v < 0 or v > 1:
            raise OracleError(f"{function} oracle returned {v}")
        return +v


def certify(function, tokens, dps_low=DPS_LOW, dps_high=DPS_HIGH):
    """Certified survival at the binary64 arguments of the tokens.

    Returns (value, agreement_digits, route). Raises OracleError when the two
    precisions disagree beyond CERTIFY_DIGITS."""
    args = exact_args(tokens)
    extra = magnitude_digits(args)
    lo = evaluate(function, args, dps_low + extra)
    hi = evaluate(function, args, dps_high + extra)
    with mp.workdps(dps_high + extra):
        if hi == 0:
            if lo != 0:
                raise OracleError(f"{function}: oracle unstable at zero")
            return hi, mp.inf, ROUTES[function][1]
        rel = abs(hi - lo) / abs(hi)
        digits = mp.inf if rel == 0 else -mp.log10(rel)
        if digits < CERTIFY_DIGITS:
            raise OracleError(f"{function}: oracle agrees with itself to only "
                              f"{mp.nstr(digits, 4)} digits")
        return hi, digits, ROUTES[function][1]


def relative_error(reference_text, value):
    """|reference - value| / |value|, at a precision that keeps every digit."""
    with mp.workdps(80):
        ref = mp.mpf(reference_text.strip())
        if value == 0:
            return mp.mpf(0) if ref == 0 else mp.inf
        return abs(ref - value) / abs(value)


def agrees(reference_text, value):
    """True when the reference meets the 25-digit floor against value."""
    return relative_error(reference_text, value) <= SER_FLOOR
