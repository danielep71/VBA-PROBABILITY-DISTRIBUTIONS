"""
Regularized incomplete gamma for the accuracy gate.

WHY THIS MODULE EXISTS

Scoring Chi-square inverse tail residuals needs the forward CDF
P(df/2, x/2). The obvious route is unusable: mpmath.gammainc raises
NoConvergence at shape a = 5E6 and a = 5E7 - that is, at df = 1E7 and
df = 1E8, which are exactly the envelope_domain rows the gate must score.
A helper that works only at df = 1E6 would fail on two thirds of the
surface it exists to measure.

So the incomplete gamma is evaluated here from the converging route
directly:

    x <  a + 1   lower series      -> P computed directly
    x >= a + 1   upper Lentz CF    -> Q computed directly

Each side is computed DIRECTLY in the regime where it converges, so the
small quantity is never obtained by subtracting two nearly equal large
ones.

NO FALLBACK. There is deliberately no mpmath.gammainc path, not even as a
last resort: a fallback would silently reintroduce the failure this module
exists to avoid, and at the shapes where it fails it fails loudly on some
rows and not others. Exhausting either iteration cap raises
IGammaNonConvergence. Callers must turn that into PENDING (main grid) or
INCOMPLETE (holdout) - never into a numerical verdict.

RELATIONSHIP TO THE CHI-SQUARE STUDY

benchmark/chisq_reference_study/chisq_reference.py contains an
independently developed implementation of the same route, used to generate
the frozen 69-point reference set. This module does NOT import it: the
study is evidence and the gate is production tooling, and the gate must not
depend on a study directory.

The two are kept honest by test_igamma_parity.py, which evaluates both at
the committed 69 quantiles and requires agreement. That fixture is a DRIFT
DETECTOR, not an independent oracle. The independent numerical authority
remains the quadrature and MPFR legs recorded in the study.
"""
import mpmath as mp

# Hitting either cap is non-convergence, never an approximation to return.
MAX_SERIES_TERMS = 4_000_000
MAX_CF_ITERATIONS = 4_000_000

# The regime boundary. Below it the lower series converges; at or above it the
# upper continued fraction does. Both routes are usable near the seam, and the
# parity fixture pins agreement immediately below, exactly at, and above it.
def seam(a):
    """The route-selection boundary x = a + 1, as a value rather than a
    scattered literal, so the fixtures and the dispatch cannot disagree."""
    return mp.mpf(a) + 1


class IGammaNonConvergence(Exception):
    """A converging route exhausted its iteration cap.

    Raised rather than returning a truncated sum. The gate turns this into
    PENDING and the holdout analyzer into INCOMPLETE; neither may emit a
    numerical verdict from it.
    """


def _log_prefactor(a, x):
    """log( x^a e^-x / Gamma(a) ), in log space to avoid overflow at large a."""
    return a * mp.log(x) - x - mp.loggamma(a)


def lower_series(a, x):
    """Regularized P(a, x) by the converging lower series. For x < a + 1.

    P(a,x) = x^a e^-x / Gamma(a+1) * sum_{n>=0} x^n / ((a+1)...(a+n))

    Returns (P, n_terms). Raises IGammaNonConvergence at the cap.
    """
    a, x = mp.mpf(a), mp.mpf(x)
    if x <= 0:
        return mp.mpf(0), 0
    tol = mp.mpf(10) ** (-(mp.mp.dps + 10))
    term = mp.mpf(1)
    total = mp.mpf(1)
    for n in range(1, MAX_SERIES_TERMS + 1):
        term = term * x / (a + n)
        total += term
        if abs(term) < abs(total) * tol:
            return mp.e ** (_log_prefactor(a, x) - mp.log(a)) * total, n
    raise IGammaNonConvergence(
        f"lower series did not converge in {MAX_SERIES_TERMS} terms "
        f"(a={mp.nstr(a, 8)}, x={mp.nstr(x, 8)}, dps={mp.mp.dps})")


def upper_cf(a, x):
    """Regularized Q(a, x) by the modified-Lentz continued fraction.
    For x >= a + 1. Returns (Q, n_iterations). Raises at the cap.
    """
    a, x = mp.mpf(a), mp.mpf(x)
    tiny = mp.mpf(10) ** (-(mp.mp.dps * 2 + 20))
    tol = mp.mpf(10) ** (-(mp.mp.dps + 10))

    b = x + 1 - a
    c = mp.mpf(1) / tiny
    d = mp.mpf(1) / b if b != 0 else mp.mpf(1) / tiny
    h = d
    for i in range(1, MAX_CF_ITERATIONS + 1):
        an = -mp.mpf(i) * (mp.mpf(i) - a)
        b += 2
        d = an * d + b
        if abs(d) < tiny:
            d = tiny
        c = b + an / c
        if abs(c) < tiny:
            c = tiny
        d = mp.mpf(1) / d
        delta = d * c
        h *= delta
        if abs(delta - 1) < tol:
            return mp.e ** _log_prefactor(a, x) * h, i
    raise IGammaNonConvergence(
        f"upper CF did not converge in {MAX_CF_ITERATIONS} iterations "
        f"(a={mp.nstr(a, 8)}, x={mp.nstr(x, 8)}, dps={mp.mp.dps})")


def reg_incomplete(a, x):
    """Evaluate the incomplete gamma on whichever side converges.

    Returns (value, side, route, iterations) where `side` is "P" or "Q" and
    `value` is that side computed DIRECTLY - never as one minus the other.
    Callers wanting the CDF should use chi2_cdf; this signature exists so the
    parity fixture can compare the informative tail and the route selection.
    """
    a, x = mp.mpf(a), mp.mpf(x)
    if x < seam(a):
        p, n = lower_series(a, x)
        return p, "P", "lower_series", n
    q, n = upper_cf(a, x)
    return q, "Q", "upper_cf", n


def reg_lower(a, x):
    """Regularized P(a, x). Returns (P, route, iterations)."""
    value, side, route, n = reg_incomplete(a, x)
    return (value if side == "P" else 1 - value), route, n


def chi2_cdf(x, df):
    """Chi-square CDF: P(df/2, x/2).

    Signature matches the other tail-residual forward CDFs - quantile first,
    shape parameters after - so the registry-driven dispatch can call it as
    cdf(x, *shape_args) with no special case.
    """
    x, df = mp.mpf(x), mp.mpf(df)
    if x <= 0:
        return mp.mpf(0)
    value, _route, _n = reg_lower(df / 2, x / 2)
    return value
