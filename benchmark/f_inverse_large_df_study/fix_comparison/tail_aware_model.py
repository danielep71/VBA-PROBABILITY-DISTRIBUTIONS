"""Study-only P4: P3 with a tail-aware residual.

P3 solves for the small member of the pair (orientation F-A) and matches the
lower tail of that member. Where the member is small and the probability is
near one, the lower tail is a number near one and its residual has lost the
upper tail: P3's error reaches 1.3E-5 at 1 - p = 1E-13, d = (5, 2E6).

P4 keeps P3's orientation, pair continued fraction, 2^-100 stop and guarded
stagnation, and changes one thing: the residual is formed on the smaller of
the two probability tails, and that tail is computed directly. The forward
evaluation returns both tails; below the mean the lower is the continued
fraction and the upper its complement, above it the upper is the continued
fraction and the lower its complement.

Copied from stagnation_model.py and pair_cf_screen.py; those files stay frozen.
"""
import math

import pair_cf_screen as pair
import stagnation_model as p3

forward = p3.forward
DD = pair.DD
STOP = 2.0 ** -100


def tails(x, a, b):
    """(lower, upper) of I_x(a, b) at a Double x, with 1 - x as an exact pair.
    The smaller tail is computed directly."""
    if x <= 0:
        return 0.0, 1.0
    y = 1.0 - x
    if y <= 0:
        return 1.0, 0.0
    if a + b >= forward.C['PROB_IBETA_LOADER_MIN_SHAPE']:
        log_bt = forward.beta_log_pdf(x, y, a, b, math.log(x), math.log(y)) + math.log(x) + math.log(y)
    else:
        log_bt = a * math.log(x) + b * math.log(y) - p3.log_beta(a, b)
    bt = math.exp(log_bt)
    saved_model, saved_eps = pair.model, forward.C['PROB_NUM_EPS']
    try:
        pair.model = p3
        forward.C['PROB_NUM_EPS'] = STOP
        if x < (a + 1.0) / (a + b + 2.0):
            cf, _ = pair.cf_pair(a, b, DD(x))
            lower = min(max(bt * cf / a, 0.0), 1.0)
            return lower, 1.0 - lower
        cf, _ = pair.cf_pair(b, a, DD(1.0) - DD(x))
        upper = min(max(bt * cf / b, 0.0), 1.0)
        return 1.0 - upper, upper
    finally:
        pair.model = saved_model
        forward.C['PROB_NUM_EPS'] = saved_eps


def inverse_pair(p, complement, a, b):
    """(x, y, trace) with I_x(a, b) = p."""
    if p <= 0:
        return 0.0, 1.0, []
    if complement <= 0:
        return 1.0, 0.0, []
    direct = p <= tails(0.5, a, b)[0]
    sa, sb = (a, b) if direct else (b, a)
    # Tails of the solved member u under (sa, sb)
    target_lower, target_upper = (p, complement) if direct else (complement, p)
    match_lower = target_lower <= target_upper

    def residual_at(u):
        lower, upper = tails(u, sa, sb)
        return lower - target_lower if match_lower else target_upper - upper

    lb = p3.log_beta(sa, sb)
    u = p3.seed(sa, sb, target_lower, lb)
    low, high, trace = 0.0, 1.0, []
    for iteration in range(1, 201):
        residual = residual_at(u)
        if residual < 0:
            low = max(low, u)
        else:
            high = min(high, u)
        density = 0.0
        if 0 < u < 1:
            density = p3.try_exp((sa - 1.0) * math.log(u) + (sb - 1.0) * math.log(1.0 - u) - lb) or 0.0
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
            fallback = True
        settled = abs(new - u) <= 2.0 ** -52 * abs(new) or new == u
        trace.append(dict(iteration=iteration, direct=direct, match_lower=match_lower, u=u,
                          residual=residual, density=density, low=low, high=high, new=new,
                          fallback=fallback, settled=settled, guarded_stagnation=stagnation))
        u = new
        if settled:
            x, y = (u, 1.0 - u) if direct else (1.0 - u, u)
            if a < 1 and b < 1:
                round_trip = tails(x, a, b)[0]
                if abs(round_trip - p) > 0.000001 * min(p, complement):
                    raise ValueError("No representable interior quantile")
            return x, y, trace
    raise ValueError("Incomplete beta inverse failed to converge")


def evaluate(p, d1, d2):
    x, y, trace = inverse_pair(p, 1.0 - p, d1 / 2.0, d2 / 2.0)
    result = p3.try_exp(math.log(d2) - math.log(d1) + math.log(x) - math.log(y))
    if result is None:
        raise ValueError("F transform overflow")
    return result, x, y, trace
