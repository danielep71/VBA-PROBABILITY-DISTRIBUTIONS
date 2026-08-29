"""
Frozen Chi-square inverse-CDF reference generator for the #22 feasibility
checkpoint (v1.0.0 plan, Track A2 item 6).

PREREGISTERED SCOPE - frozen before any value was generated:

  df                : 1E6, 1E7, 1E8
  fitting baseline  : 0.5, 0.9, 0.99                                    (3)
  fitting bridge    : 0.125, 0.25, 0.375, 0.625, 0.75, 0.875            (6)
  fitting tails     : 2^-30, 2^-20, 1-2^-20, 1-2^-30                    (4)
  holdout bridge    : 0.0625, 0.1875, 0.3125, 0.4375,
                      0.5625, 0.6875, 0.8125, 0.9375                    (8)
  holdout tails     : 2^-25, 1-2^-25                                    (2)

  13 fitting + 10 holdout = 23 probabilities per df = 69 references.

Every probability is carried as the EXACT IEEE-754 binary64 value a VBA Double
would hold, including 0.9 and 0.99, which are not exactly representable. The
exact value is recorded as a hex float so the binding is unambiguous.

METHOD

  Chi-square quantile q solves P(a, x) = p with a = df/2, x = q/2, where P is
  the regularized lower incomplete gamma function.

  P/Q are evaluated by the converging route required by the plan:
    * lower series      for x < a + 1
    * upper Lentz CF    for x >= a + 1
  Both carry an explicit convergence flag and iteration count. Neither is
  allowed to return a truncated value silently: exhausting the iteration cap is
  a hard non-convergence, not an approximation.

  Inversion is Newton on log x (the dynamic range across the tails is large and
  x must stay positive), seeded from Wilson-Hilferty and safeguarded by a
  bracket. The analytic derivative is used, so no finite differences enter.

STABILIZATION

  Every quantile is produced independently at two materially separated working
  precisions. A value is accepted only if the two agree to the required number
  of significant digits; otherwise it is REJECTED and recorded as such. No
  point may be silently dropped: any rejection fails the checkpoint.

This module produces the mpmath-route values only. It records no final
agreement figure and does not mark the checkpoint complete: the independent
Rmpfr leg has not run.
"""
import json
import platform
import time
from fractions import Fraction

import mpmath
from mpmath import mp, mpf

# --------------------------------------------------------------------------
# Frozen inputs
# --------------------------------------------------------------------------

DF_VALUES = (1e6, 1e7, 1e8)

# Built as Python floats so each is exactly the binary64 value VBA would hold.
# 0.9 and 0.99 are deliberately NOT written as decimal strings downstream; the
# exact hex representation travels with every record.
FITTING_BASELINE = (0.5, 0.9, 0.99)
FITTING_BRIDGE = (0.125, 0.25, 0.375, 0.625, 0.75, 0.875)
FITTING_TAILS = (2.0 ** -30, 2.0 ** -20, 1.0 - 2.0 ** -20, 1.0 - 2.0 ** -30)
HOLDOUT_BRIDGE = (0.0625, 0.1875, 0.3125, 0.4375,
                  0.5625, 0.6875, 0.8125, 0.9375)
HOLDOUT_TAILS = (2.0 ** -25, 1.0 - 2.0 ** -25)

ARMS = (
    ("fitting", "baseline", FITTING_BASELINE),
    ("fitting", "bridge", FITTING_BRIDGE),
    ("fitting", "tails", FITTING_TAILS),
    ("holdout", "bridge", HOLDOUT_BRIDGE),
    ("holdout", "tails", HOLDOUT_TAILS),
)

# Materially separated working precisions. The gap is ~2x, well beyond any
# plausible shared truncation, so agreement is evidence rather than coincidence.
DPS_LOW = 60
DPS_HIGH = 120

# A quantile must reproduce to at least this many significant digits across the
# two precisions to be accepted.
REQUIRED_AGREEMENT_DIGITS = 40

# Hard caps. Hitting one is non-convergence, never a returned approximation.
MAX_SERIES_TERMS = 4_000_000
MAX_CF_ITERATIONS = 4_000_000
MAX_NEWTON_STEPS = 200


class NonConvergence(Exception):
    """Raised when a route exhausts its iteration cap. Never swallowed."""


# --------------------------------------------------------------------------
# Exact binary64 handling
# --------------------------------------------------------------------------

def exact_mpf(value):
    """The exact binary64 value of `value`, carried exactly into mpf.

    mpf(float) is exact, but going through Fraction makes the intent explicit
    and keeps the value exact even if the working precision is later reduced
    below binary64's significand.
    """
    fr = Fraction(value)
    return mpf(fr.numerator) / mpf(fr.denominator)


def describe_probability(value):
    """Record a probability unambiguously: hex is authoritative, decimal aids reading."""
    fr = Fraction(value)
    return {
        "float_repr": repr(value),
        "hex": float.hex(value),
        "exact_numerator": str(fr.numerator),
        "exact_denominator": str(fr.denominator),
    }


# --------------------------------------------------------------------------
# Regularized incomplete gamma: converging series / Lentz continued fraction
# --------------------------------------------------------------------------

def _log_prefactor(a, x):
    """log( x^a e^-x / Gamma(a) ) computed in log space to avoid overflow."""
    return a * mp.log(x) - x - mp.loggamma(a)


def lower_series(a, x):
    """Regularized P(a, x) by the converging lower series.

    P(a,x) = x^a e^-x / Gamma(a+1) * sum_{n>=0} x^n / ((a+1)...(a+n))

    Converges well for x < a + 1. Returns (P, n_terms).
    """
    if x <= 0:
        return mpf(0), 0
    term = mpf(1)
    total = mpf(1)
    tol = mpf(10) ** (-(mp.dps + 10))
    for n in range(1, MAX_SERIES_TERMS + 1):
        term = term * x / (a + n)
        total += term
        if abs(term) < abs(total) * tol:
            pref = _log_prefactor(a, x) - mp.log(a)
            return mp.e ** (pref) * total, n
    raise NonConvergence(
        f"lower series did not converge in {MAX_SERIES_TERMS} terms "
        f"(a={mp.nstr(a, 8)}, x={mp.nstr(x, 8)}, dps={mp.dps})")


def upper_cf(a, x):
    """Regularized Q(a, x) by the modified-Lentz continued fraction.

    Q(a,x) = x^a e^-x / Gamma(a) * 1/(x+1-a - 1*(1-a)/(x+3-a - 2*(2-a)/(...)))

    Converges well for x >= a + 1. Returns (Q, n_iterations).
    """
    tiny = mpf(10) ** (-(mp.dps * 2 + 20))
    tol = mpf(10) ** (-(mp.dps + 10))

    b = x + 1 - a
    c = mpf(1) / tiny
    d = mpf(1) / b if b != 0 else mpf(1) / tiny
    h = d
    for i in range(1, MAX_CF_ITERATIONS + 1):
        an = -mpf(i) * (mpf(i) - a)
        b += 2
        d = an * d + b
        if abs(d) < tiny:
            d = tiny
        c = b + an / c
        if abs(c) < tiny:
            c = tiny
        d = mpf(1) / d
        delta = d * c
        h *= delta
        if abs(delta - 1) < tol:
            return mp.e ** _log_prefactor(a, x) * h, i
    raise NonConvergence(
        f"upper CF did not converge in {MAX_CF_ITERATIONS} iterations "
        f"(a={mp.nstr(a, 8)}, x={mp.nstr(x, 8)}, dps={mp.dps})")


def reg_lower(a, x):
    """Regularized P(a, x) via the route appropriate to the argument.

    Returns (P, route, iterations). The switch at x = a + 1 is the standard
    crossover where each route is in its converging regime.
    """
    if x < a + 1:
        p, n = lower_series(a, x)
        return p, "lower_series", n
    q, n = upper_cf(a, x)
    return 1 - q, "upper_cf", n


def reg_both(a, x):
    """Return (P, Q, route, iterations), keeping whichever side was computed
    directly so the small side never comes from cancelling the large one."""
    if x < a + 1:
        p, n = lower_series(a, x)
        return p, 1 - p, "lower_series", n
    q, n = upper_cf(a, x)
    return 1 - q, q, "upper_cf", n


# --------------------------------------------------------------------------
# Inversion
# --------------------------------------------------------------------------

def wilson_hilferty_start(df, p):
    """Wilson-Hilferty seed for the Chi-square quantile.

    Accurate to roughly 1e-8 relative at these df, so Newton reaches full
    working precision in a handful of steps.
    """
    z = mp.sqrt(2) * mp.erfinv(2 * p - 1)
    t = 2 / (9 * df)
    base = 1 - t + z * mp.sqrt(t)
    if base <= 0:
        # Deep lower tail where the cubic seed would go non-positive.
        base = mpf(1) / (9 * df)
    return df * base ** 3


def log_pdf_chi2(df, q):
    """log of the Chi-square density at q, in log space."""
    k = df / 2
    return (k - 1) * mp.log(q) - q / 2 - mp.loggamma(k) - k * mp.log(2)


def invert(df, p, dps):
    """Solve P(df/2, q/2) = p for q at working precision `dps`.

    Newton on u = log q with the analytic derivative, safeguarded by a bracket
    that is tightened from the sign of the residual. Returns a record including
    the route used, iteration counts, and the final residual.
    """
    mp.dps = dps
    a = mpf(df) / 2
    p_exact = exact_mpf(p) if isinstance(p, float) else p

    q = wilson_hilferty_start(mpf(df), p_exact)
    lo, hi = mpf(0), mp.inf

    tol = mpf(10) ** (-(dps - 8))
    route = None
    iters = 0
    steps = 0
    for steps in range(1, MAX_NEWTON_STEPS + 1):
        pv, qv, route, iters = reg_both(a, q / 2)
        resid = pv - p_exact
        # An exact hit is convergence, not a degenerate bracket. Without this,
        # the update below sets lo = q, the strict guard `lo < q_new` then
        # rejects the (correct, unchanged) Newton point, and the safeguard
        # bisects away from an exact answer until the step cap. The upper-tail
        # points reach exactly zero because Q is computed directly there, so
        # they are what expose it.
        if resid == 0:
            break
        if resid > 0:
            hi = min(hi, q)
        else:
            lo = max(lo, q)

        # d/du P where u = log q:  P'(q) * q
        dP_du = mp.e ** log_pdf_chi2(mpf(df), q) * q
        if dP_du == 0:
            raise NonConvergence(
                f"zero derivative at df={df:.0e}, p={p!r}, dps={dps}")

        u_new = mp.log(q) - resid / dP_du
        q_new = mp.e ** u_new
        if not (lo < q_new < hi):
            # Newton left the bracket; bisect in log space instead.
            if hi == mp.inf:
                q_new = q * 2
            else:
                q_new = mp.sqrt(max(lo, mpf(10) ** (-(dps))) * hi)

        rel = abs(q_new - q) / abs(q_new)
        # Evaluate BEFORE rebinding q: testing `q_new == q` after the
        # assignment is trivially true and stops Newton after one step.
        converged = (rel < tol) or (q_new == q)
        q = q_new
        if converged:
            break
    else:
        raise NonConvergence(
            f"Newton did not converge in {MAX_NEWTON_STEPS} steps "
            f"(df={df:.0e}, p={p!r}, dps={dps})")

    pv, qv, route, iters = reg_both(a, q / 2)
    resid = pv - p_exact
    denom = min(p_exact, 1 - p_exact)
    # Score the residual on whichever side was computed DIRECTLY. For an upper
    # quantile the informative quantity is Q, and forming P = 1 - Q first would
    # route the small side through a complement subtraction before it is
    # measured. Mathematically |P - p| == |Q - (1-p)|; taking the direct side
    # keeps the measured quantity free of that cancellation.
    if route == "upper_cf":
        resid_direct = qv - (1 - p_exact)
    else:
        resid_direct = pv - p_exact
    return {
        "q": q,
        "route": route,
        "route_iterations": iters,
        "newton_steps": steps,
        "residual_abs": abs(resid_direct),
        "residual_side": "Q" if route == "upper_cf" else "P",
        "tail_residual": abs(resid_direct) / denom,
        "P_at_q": pv,
        "Q_at_q": qv,
    }


# --------------------------------------------------------------------------
# Two-precision stabilization
# --------------------------------------------------------------------------

def agreement_digits(x, y):
    """Significant decimal digits to which x and y agree."""
    mp.dps = max(DPS_HIGH, DPS_LOW) + 20
    x = mpf(x)
    y = mpf(y)
    if x == y:
        return mpf("inf")
    if x == 0 or y == 0:
        return mpf(0)
    return -mp.log10(abs(x - y) / abs(x))


def generate_point(df, p, arm, band):
    """Produce one reference, stabilized across two precisions.

    A point is ACCEPTED only if the low- and high-precision inversions agree to
    at least REQUIRED_AGREEMENT_DIGITS. Anything else is REJECTED with the
    reason recorded. Non-convergence propagates as a rejection, never as a
    returned value.
    """
    started = time.time()
    record = {
        "df": df,
        "arm": arm,
        "band": band,
        "probability": describe_probability(p),
        "precision_pair": [DPS_LOW, DPS_HIGH],
    }
    try:
        low = invert(df, p, DPS_LOW)
        high = invert(df, p, DPS_HIGH)
    except NonConvergence as exc:
        record.update(status="REJECTED", reason=f"non-convergence: {exc}",
                      runtime_s=round(time.time() - started, 3))
        return record

    digits = agreement_digits(low["q"], high["q"])
    mp.dps = DPS_HIGH
    stabilized = digits >= REQUIRED_AGREEMENT_DIGITS

    record.update({
        "status": "ACCEPTED" if stabilized else "REJECTED",
        "stabilization_digits": (None if digits == mp.inf
                                 else float(mp.nstr(digits, 6))),
        # Store at MORE than the working precision. Storing fewer digits than
        # were computed silently caps every downstream agreement figure at the
        # stored width: an earlier revision wrote 50 digits from a 120-dps
        # computation, and every cross-check then reported ~49 digits of
        # agreement, which measured the truncation rather than the routes.
        "quantile": mp.nstr(high["q"], DPS_HIGH + 10),
        "quantile_low_precision": mp.nstr(mpf(low["q"]), DPS_LOW + 10),
        "route": high["route"],
        "route_iterations_low": low["route_iterations"],
        "route_iterations_high": high["route_iterations"],
        "newton_steps_low": low["newton_steps"],
        "newton_steps_high": high["newton_steps"],
        "tail_residual": mp.nstr(high["tail_residual"], 6),
        "residual_abs": mp.nstr(high["residual_abs"], 6),
        "residual_side": high["residual_side"],
        "runtime_s": round(time.time() - started, 3),
    })
    if not stabilized:
        record["reason"] = (
            f"did not stabilize: agreed to {mp.nstr(digits, 6)} digits, "
            f"required {REQUIRED_AGREEMENT_DIGITS}")
    return record


def frozen_points():
    """The 69 frozen (df, p, arm, band) inputs, in a fixed order."""
    out = []
    for df in DF_VALUES:
        for arm, band, probs in ARMS:
            for p in probs:
                out.append((df, p, arm, band))
    return out


def main(out_path="chisq_reference.json"):
    points = frozen_points()
    assert len(points) == 69, f"expected 69 frozen points, got {len(points)}"

    started = time.time()
    records = []
    for i, (df, p, arm, band) in enumerate(points, 1):
        rec = generate_point(df, p, arm, band)
        records.append(rec)
        print(f"  [{i:2d}/69] df={df:.0e} {arm}/{band:8s} p={p!r:24s} "
              f"{rec['status']}", flush=True)

    accepted = sum(1 for r in records if r["status"] == "ACCEPTED")
    rejected = [r for r in records if r["status"] != "ACCEPTED"]
    mp.dps = DPS_HIGH
    stab = [r["stabilization_digits"] for r in records
            if r.get("stabilization_digits") is not None]

    payload = {
        "checkpoint": "v1.0.0 plan Track A2 item 6 - Chi-square reference feasibility",
        "status": "PROVISIONAL - INCOMPLETE",
        "completion_blocked_on": (
            "independent Rmpfr leg has not been run; no final agreement figure "
            "is recorded and the checkpoint must not be marked complete"),
        "generated_utc": time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime()),
        "runtime_s_total": round(time.time() - started, 3),
        "versions": {
            "python": platform.python_version(),
            "mpmath": mpmath.__version__,
            "platform": platform.platform(),
        },
        "design": {
            "df_values": list(DF_VALUES),
            "fitting_baseline": list(FITTING_BASELINE),
            "fitting_bridge": list(FITTING_BRIDGE),
            "fitting_tails": [float.hex(v) for v in FITTING_TAILS],
            "holdout_bridge": list(HOLDOUT_BRIDGE),
            "holdout_tails": [float.hex(v) for v in HOLDOUT_TAILS],
            "points_total": len(points),
            "precision_pair": [DPS_LOW, DPS_HIGH],
            "required_agreement_digits": REQUIRED_AGREEMENT_DIGITS,
            "measures": [
                "quantile (Chi-square inverse CDF at exact binary64 p)",
                "tail_residual = |P(a, q/2) - p| / min(p, 1-p)",
            ],
        },
        "convergence": {
            "points_total": len(points),
            "accepted": accepted,
            "rejected": len(rejected),
            "all_converged_and_stabilized": not rejected,
            "min_stabilization_digits": (min(stab) if stab else None),
            "rejections": [
                {"df": r["df"], "arm": r["arm"], "band": r["band"],
                 "probability": r["probability"]["hex"],
                 "reason": r.get("reason")}
                for r in rejected
            ],
        },
        "references": records,
    }
    with open(out_path, "w", newline="\n") as f:
        json.dump(payload, f, indent=2)
        f.write("\n")
    print(f"\n{accepted}/{len(points)} accepted, {len(rejected)} rejected "
          f"-> {out_path}")
    if rejected:
        print("CHECKPOINT FAILS: at least one point did not converge or stabilize.")
        return 1
    print("mpmath route converged and stabilized at every frozen point.")
    print("Checkpoint remains PROVISIONAL: the Rmpfr leg has not run.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
