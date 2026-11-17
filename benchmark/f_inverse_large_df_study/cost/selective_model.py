"""Study-only candidate S for #359: the tail-aware beta inverse at a lower cost.

The accepted rule (P4, ../fix_comparison/tail_aware_model.py) evaluates the
incomplete beta with a pair continued fraction, stopped at 2^-100, at every
Newton step. In Excel that is 6 to 580 times the cost of the routine before
#35 (../production/PRODUCTION_RESULT.md).

Candidate S keeps P4's orientation, which is what repaired #35: it solves for
the member of the pair below 1/2, on the smaller probability tail. It changes
which forward evaluation the Newton iteration uses:

1. Phase A. Iterate with the binary64 forward function of the source, its
   continued fraction stopped at 4 * 2^-52 instead of 3E-14, with the
   convergence rule of the source. The two tails are
   formed as P4 forms them: below the mean the lower tail is the forward
   function at u and the upper its complement; above it the upper tail is the
   forward function at 1 - u with the shapes exchanged and the lower its
   complement. Below the mean nothing is evaluated at 1 - u.
   A settled iterate below the mean is then compared with its two adjacent
   Doubles, and the one with the smallest residual is returned.
2. Phase B, only when the iterate phase A settles on lies on the reflected
   branch of the forward function, u >= (sa + 1) / (sa + sb + 2), or when
   phase A has not settled within PHASE_A_ITER iterations: continue from
   it with P4's loop, pair fraction and guarded stagnation included. That
   branch evaluates the fraction at 1 - u, which is exact only as a pair; it
   is where the second loss of #35 was found
   (../fix_comparison/RESIDUAL_DIAGNOSIS.md).

On the direct branch the argument of the fraction is u itself, an exact Double,
and no pair arithmetic is used at all.

Nothing here is the VBA: it is the model the registration is written from.
The frozen files of ../fix_comparison are imported, not edited.
"""
import math
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE.parent / "fix_comparison"))

import pair_cf_screen as pair  # noqa: E402
import tail_aware_model as p4  # noqa: E402

p3 = p4.p3
forward = p4.forward
MACH_EPS = 2.0 ** -52
MAX_ITER = 200
# The fitting sweep chose these two. With the source's stop of 3E-14 the
# candidate is up to 2.8E-14 farther from the reference than P4 at small
# shapes; with the tight stop and the neighbour comparison it is not.
BINARY64_STOP = 4 * 2.0 ** -52
PHASE_A_ITER = 60

# Counters of forward evaluations, for the cost model; reset by the caller.
COUNTS = {"binary64": 0, "pair": 0, "pair_terms": 0}
_cf_pair = pair.cf_pair


def _counting_cf_pair(a, b, x):
    value, terms = _cf_pair(a, b, x)
    COUNTS["pair"] += 1
    COUNTS["pair_terms"] += terms
    return value, terms


pair.cf_pair = _counting_cf_pair


def reset_counts():
    for key in COUNTS:
        COUNTS[key] = 0


def lower_binary64(u, sa, sb):
    COUNTS["binary64"] += 1
    saved = forward.C["PROB_NUM_EPS"]
    try:
        forward.C["PROB_NUM_EPS"] = BINARY64_STOP
        return forward.beta_regularized(u, 1.0 - u, sa, sb)
    finally:
        forward.C["PROB_NUM_EPS"] = saved


def reflected(u, sa, sb):
    """Whether the forward function evaluates its fraction at 1 - u."""
    return not (u < (sa + 1.0) / (sa + sb + 2.0))


def density_at(u, sa, sb, lb):
    if not 0.0 < u < 1.0:
        return 0.0
    return p3.try_exp((sa - 1.0) * math.log(u) + (sb - 1.0) * math.log(1.0 - u) - lb) or 0.0


def tails_binary64(u, sa, sb):
    """(lower, upper) at u with the binary64 forward function, split as P4's
    tails: the continued fraction is asked for the smaller-argument side."""
    if not reflected(u, sa, sb):
        lower = lower_binary64(u, sa, sb)
        return lower, 1.0 - lower
    upper = lower_binary64(1.0 - u, sb, sa)
    return 1.0 - upper, upper


def phase_a(u, sa, sb, lb, target_lower, target_upper, match_lower):
    """The Newton loop of the source before #35, on the matched tail."""
    def residual_at(v):
        lower, upper = tails_binary64(v, sa, sb)
        return lower - target_lower if match_lower else target_upper - upper

    low, high = 0.0, 1.0
    for _ in range(PHASE_A_ITER):
        residual = residual_at(u)
        if residual < 0:
            low = max(low, u)
        else:
            high = min(high, u)
        density = density_at(u, sa, sb, lb)
        new = 0.5 * (low + high) if density <= 0 else u - residual / density
        if density > 0 and (new <= low or new >= high):
            new = 0.5 * (low + high)
        settled = abs(new - u) <= MACH_EPS * abs(new) or new == u
        u = new
        if settled:
            break
    else:
        return u, False
    if reflected(u, sa, sb):
        return u, True
    # Below the mean: the best of the settled Double and its two neighbours.
    candidates = [u]
    for toward in (0.0, 1.0):
        neighbour = math.nextafter(u, toward)
        if 0.0 < neighbour < 1.0 and not reflected(neighbour, sa, sb):
            candidates.append(neighbour)
    return min(candidates, key=lambda v: abs(residual_at(v))), True


def phase_b(u, sa, sb, lb, target_lower, target_upper, match_lower):
    """P4's loop from u: pair fraction, guarded stagnation, fresh bracket."""
    def residual_at(v):
        lower, upper = p4.tails(v, sa, sb)
        return lower - target_lower if match_lower else target_upper - upper

    low, high = 0.0, 1.0
    for _ in range(MAX_ITER):
        residual = residual_at(u)
        if residual < 0:
            low = max(low, u)
        else:
            high = min(high, u)
        density = density_at(u, sa, sb, lb)
        fallback = density <= 0
        new = 0.5 * (low + high) if fallback else u - residual / density
        stagnation = False
        if not fallback and new == u and 0.0 < u < 1.0:
            below, above = math.nextafter(u, 0.0), math.nextafter(u, 1.0)
            if 0.0 < below < u < above < 1.0:
                lo, up = residual_at(below), residual_at(above)
                stagnation = (all(math.isfinite(v) for v in (lo, residual, up))
                              and lo <= residual <= up and lo <= 0.0 <= up
                              and abs(residual) <= min(abs(lo), abs(up)))
        if not fallback and not stagnation and (new <= low or new >= high):
            new = 0.5 * (low + high)
        settled = abs(new - u) <= MACH_EPS * abs(new) or new == u
        u = new
        if settled:
            return u
    raise ValueError("Incomplete beta inverse failed to converge (phase B)")


def inverse_pair(p, complement, a, b):
    """(x, y, used_pair) with I_x(a, b) = p."""
    if p <= 0:
        return 0.0, 1.0, False
    if complement <= 0:
        return 1.0, 0.0, False
    direct = p <= lower_binary64(0.5, a, b)
    sa, sb = (a, b) if direct else (b, a)
    target_lower, target_upper = (p, complement) if direct else (complement, p)
    if target_lower >= 1.0:
        raise ValueError("legacy regime: the source answers with the routine before #35")
    match_lower = target_lower <= target_upper
    lb = p3.log_beta(sa, sb)
    u = p3.seed(sa, sb, target_lower, lb)
    u, settled = phase_a(u, sa, sb, lb, target_lower, target_upper, match_lower)
    used_pair = reflected(u, sa, sb) or not settled
    if used_pair:
        u = phase_b(u, sa, sb, lb, target_lower, target_upper, match_lower)
    x, y = (u, 1.0 - u) if direct else (1.0 - u, u)
    return x, y, used_pair


def f_quantile(p, d1, d2, x, y):
    result = p3.try_exp(math.log(d2) - math.log(d1) + math.log(x) - math.log(y))
    if result is None:
        raise ValueError("F transform overflow")
    return result
