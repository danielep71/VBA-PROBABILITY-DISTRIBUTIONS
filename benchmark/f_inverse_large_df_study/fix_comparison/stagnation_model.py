"""Study-only P3: P2 with neighboring-float guarded Newton stagnation.

Copied production-order model; original vba_order_model.py remains frozen.
"""
import importlib.util
import math
from pathlib import Path

ROOT = Path(__file__).resolve().parents[3]


def load_model(name, path):
    spec = importlib.util.spec_from_file_location(name, path)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


forward = load_model("f35_forward", ROOT / "benchmark/beta_balanced_survival/beta_route_model.py")
normal = load_model("f35_normal", ROOT / "benchmark/normal_tail_study/kernel_model.py")
forward.X = float
_original_log_beta = forward.log_beta


def log_gamma_delta(z, increment):
    p = forward.SOURCE_LANCZOS
    zm1 = z - 1.0
    az = p[0]
    for k in range(1, 9):
        az = az + p[k] / (zm1 + k)
    da = p[1] / ((zm1 + 1.0) * (zm1 + 1.0 + increment))
    for k in range(2, 9):
        da = da + p[k] / ((zm1 + k) * (zm1 + k + increment))
    da = -increment * da
    t = z + forward.C["PROB_LANCZOS_G"] - 0.5
    return (increment * math.log(t)
            + (z + increment - 0.5) * forward.log1p(increment / t)
            - increment + forward.log1p(da / az))


def log_beta(a, b, lanczos=forward.SOURCE_LANCZOS):
    if (b == 0.5 and a >= 1) or (a == 0.5 and b >= 1):
        return _original_log_beta(a, b, lanczos)
    large, small = max(a, b), min(a, b)
    if large >= 1 and small / large < forward.C["PROB_LOGBETA_STABLE_RATIO"]:
        return forward.log_gamma(small, lanczos) - log_gamma_delta(large, small)
    return _original_log_beta(a, b, lanczos)


forward.log_beta = log_beta


def try_exp(x):
    if not math.isfinite(x):
        return None
    try:
        return math.exp(x)
    except OverflowError:
        return None


def seed(sa, sb, target, lb):
    if sa > 1 and sb > 1:
        z = normal.inv_cdf_raw(target)
        r = (z * z - 3.0) / 6.0
        s1, s2 = 1.0 / (2.0 * sa - 1.0), 1.0 / (2.0 * sb - 1.0)
        hh = 2.0 / (s1 + s2)
        w = z * math.sqrt(hh + r) / hh - (s2 - s1) * (r + 5.0 / 6.0 - 2.0 / (3.0 * hh))
        factor = try_exp(2.0 * w)
        u = sa / (sa + sb * factor) if factor is not None else 0.0
    else:
        u = try_exp((math.log(target) + math.log(sa) + lb) / sa)
        if u is None:
            u = 0.5
    if u <= 0 or u >= 1 or not math.isfinite(u):
        u = 0.5 * target + 0.25
    return u


def guarded_stagnation(u, proposal, residual, density, sa, sb, target):
    if proposal != u or not (0.0 < u < 1.0) or not math.isfinite(density) or density <= 0:
        return False
    below, above = math.nextafter(u, 0.0), math.nextafter(u, 1.0)
    if not (0.0 < below < u < above < 1.0):
        return False
    lower = forward.beta_regularized(below, 1.0-below, sa, sb)-target
    upper = forward.beta_regularized(above, 1.0-above, sa, sb)-target
    return (all(math.isfinite(v) for v in (lower, residual, upper))
            and lower <= residual <= upper and lower <= 0.0 <= upper
            and abs(residual) <= min(abs(lower), abs(upper)))


def inverse_pair(p, complement, a, b, candidate="F-0"):
    if p <= 0:
        return 0.0, 1.0, []
    if complement <= 0:
        return 1.0, 0.0, []
    direct = p <= complement
    if candidate == "F-A":
        direct = p <= forward.beta_regularized(0.5, 0.5, a, b)
    elif candidate == "F-B":
        direct = seed(a, b, p, log_beta(a, b)) <= 0.5
    elif candidate != "F-0":
        raise ValueError("Unknown candidate")
    sa, sb, target = (a, b, p) if direct else (b, a, complement)
    lb = log_beta(sa, sb)
    u = seed(sa, sb, target, lb)
    low, high, trace = 0.0, 1.0, []
    for iteration in range(1, 201):
        parts = {}
        value = forward.beta_regularized(u, 1.0 - u, sa, sb, trace=parts)
        residual = value - target
        if residual < 0:
            low = max(low, u)
        else:
            high = min(high, u)
        density = 0.0
        if 0 < u < 1:
            density = try_exp((sa - 1.0) * math.log(u) + (sb - 1.0) * math.log(1.0 - u) - lb) or 0.0
        fallback = density <= 0
        new = 0.5 * (low + high) if fallback else u - residual / density
        stagnation = not fallback and guarded_stagnation(u, new, residual, density, sa, sb, target)
        if not fallback and not stagnation and (new <= low or new >= high):
            new = 0.5 * (low + high)
            fallback = True
        settled = abs(new - u) <= 2.0**-52 * abs(new) or new == u
        trace.append(dict(iteration=iteration, direct=direct, target=target, u=u,
                          residual=residual, density=density, low=low, high=high,
                          new=new, fallback=fallback, settled=settled, guarded_stagnation=stagnation, forward=parts))
        u = new
        if settled:
            x, y = (u, 1.0 - u) if direct else (1.0 - u, u)
            if a < 1 and b < 1:
                rt = forward.beta_regularized(x, y, a, b)
                if abs(rt - p) > 0.000001 * min(p, complement):
                    raise ValueError("No representable interior quantile")
            return x, y, trace
    raise ValueError("Incomplete beta inverse failed to converge")


def f_inverse(p, d1, d2, candidate="F-0"):
    x, y, trace = inverse_pair(p, 1.0 - p, d1 / 2.0, d2 / 2.0, candidate)
    result = try_exp(math.log(d2) - math.log(d1) + math.log(x) - math.log(y))
    if result is None:
        raise ValueError("F transform overflow")
    return result, x, y, trace


def evaluate(p, d1, d2):
    import pair_cf_screen as pair
    original_forward, original_eps = forward.beta_regularized, forward.C['PROB_NUM_EPS']
    pair_forward_model = pair.model
    try:
        pair.model = __import__(__name__)
        forward.C['PROB_NUM_EPS'] = 2.0**-100
        forward.beta_regularized = pair.forward_pair
        return f_inverse(p, d1, d2, 'F-A')
    finally:
        forward.beta_regularized = original_forward
        forward.C['PROB_NUM_EPS'] = original_eps
        pair.model = pair_forward_model
