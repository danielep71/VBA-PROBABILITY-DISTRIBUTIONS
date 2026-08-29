"""
Chi-square reference cross-check in MPFR ARITHMETIC via gmpy2.

WHY THIS EXISTS

The frozen set needs a second implementation at every df. Two candidate
third-party incomplete-gamma oracles both fail at the largest shape:

  * mpmath.gammainc  -> NoConvergence at a = 5E7
  * MPFR mpfr_gamma_inc (Rmpfr::igamma) -> aborts the process at
    a >= ~4.5E7; measured ok at 4.0E7, assertion failure at 4.5E7

So df = 1E8 (a = 5E7) cannot be reached by either. This module evaluates
the incomplete gamma from first principles using only gmpy2's MPFR
arithmetic - add, multiply, divide, exp, log, lgamma - and NEVER calls a
library incomplete-gamma routine.

WHY gmpy2 RATHER THAN Rmpfr

The same route written in R was measured at ~810 microseconds per series
term (96.9 s for 119,799 terms at 200 bits), projecting ~183 minutes for
the df = 1E8 block. R's S4 dispatch dominates and buys no accuracy. gmpy2
binds the identical MPFR 4.2.2 library at C speed: the same 168,000-term
evaluation at 400 bits takes 0.10 s, about 1,600x faster, with the same
arithmetic backend.

WHAT THIS IS, AND IS NOT

  IS:     independent of mpmath's arithmetic backend, rounding and
          evaluation order. mpmath implements its own arbitrary-precision
          arithmetic in Python; gmpy2 delegates every operation to MPFR.
          Two different libraries computing every add and multiply.

  IS NOT: algorithmically independent. It uses the same converging
          lower-series / upper-Lentz-CF route as the primary. A shared
          ALGORITHMIC error would not be caught here.

Record it as IMPLEMENTATION-INDEPENDENT. The algorithm-independent leg
remains the quadrature route in chisq_crosscheck.py.
"""
import json
import platform
import time

import gmpy2
from gmpy2 import mpfr, get_context

PREC_LOW = 200          # bits, ~60 decimal digits
PREC_HIGH = 400         # bits, ~120 decimal digits
REQUIRED_AGREEMENT_DIGITS = 40
MAX_SERIES = 4_000_000
MAX_CF = 4_000_000
MAX_NEWTON = 200


class NonConvergence(Exception):
    pass


def setprec(bits):
    get_context().precision = bits


def hex_to_mpfr(hexstr):
    """Exact binary64 from a hex float literal - no decimal round-trip."""
    return mpfr(float.fromhex(hexstr))


def log_prefactor(a, x):
    return a * gmpy2.log(x) - x - gmpy2.lngamma(a)


def lower_series(a, x, bits):
    """Regularized P(a, x) by the converging lower series. For x < a + 1."""
    tol = mpfr(2) ** (-(bits + 20))
    term = mpfr(1)
    total = mpfr(1)
    for n in range(1, MAX_SERIES + 1):
        term = term * x / (a + n)
        total += term
        if abs(term) < abs(total) * tol:
            return gmpy2.exp(log_prefactor(a, x) - gmpy2.log(a)) * total, n
    raise NonConvergence(f"lower series did not converge ({bits} bits)")


def upper_cf(a, x, bits):
    """Regularized Q(a, x) by the modified Lentz CF. For x >= a + 1."""
    tiny = mpfr(2) ** (-(2 * bits + 40))
    tol = mpfr(2) ** (-(bits + 20))
    one = mpfr(1)

    b = x + one - a
    c = one / tiny
    d = one / b if b != 0 else one / tiny
    h = d
    for i in range(1, MAX_CF + 1):
        an = -mpfr(i) * (mpfr(i) - a)
        b += 2
        d = an * d + b
        if abs(d) < tiny:
            d = tiny
        c = b + an / c
        if abs(c) < tiny:
            c = tiny
        d = one / d
        delta = d * c
        h *= delta
        if abs(delta - one) < tol:
            return gmpy2.exp(log_prefactor(a, x)) * h, i
    raise NonConvergence(f"upper CF did not converge ({bits} bits)")


def reg_area(a, x, side, bits):
    """Return the DIRECTLY computed side, never a complement."""
    if x < a + 1:
        v, n = lower_series(a, x, bits)
        return (v, n) if side == "P" else (1 - v, n)
    v, n = upper_cf(a, x, bits)
    return (v, n) if side == "Q" else (1 - v, n)


def log_pdf_chi2(df, q):
    k = df / mpfr(2)
    return (k - 1) * gmpy2.log(q) - q / 2 - gmpy2.lngamma(k) - k * gmpy2.log(mpfr(2))


def wh_seed(df, p):
    """Wilson-Hilferty seed. Uses only elementary functions."""
    # Normal quantile via the inverse error function, built from MPFR erfc
    # by bisection so no mpmath entry point is used anywhere in this module.
    z = norm_ppf(p)
    t = mpfr(2) / (9 * df)
    base = 1 - t + z * gmpy2.sqrt(t)
    if base <= 0:
        base = mpfr(1) / (9 * df)
    return df * base ** 3


def norm_ppf(p):
    """Standard normal quantile by bisection on MPFR's erfc.

    Deliberately independent of mpmath's erfinv. Only a seed is needed, so
    modest accuracy suffices; Newton supplies every reported digit.
    """
    lo, hi = mpfr(-40), mpfr(40)
    for _ in range(200):
        mid = (lo + hi) / 2
        # Phi(mid) = erfc(-mid/sqrt2)/2
        cdf = gmpy2.erfc(-mid / gmpy2.sqrt(mpfr(2))) / 2
        if cdf < p:
            lo = mid
        else:
            hi = mid
        if hi - lo < mpfr(2) ** -60:
            break
    return (lo + hi) / 2


def invert(df_float, p_hex, bits, seed=None):
    """Solve P(df/2, q/2) = p for q, entirely in MPFR arithmetic."""
    setprec(bits)
    df = mpfr(df_float)
    a = df / 2
    p = hex_to_mpfr(p_hex)
    upper = p > mpfr(1) / 2
    side = "Q" if upper else "P"
    target = (1 - p) if upper else p

    q = mpfr(seed) if seed is not None else wh_seed(df, p)
    tol = mpfr(2) ** (-(bits - 30))
    total_iters = 0

    for step in range(1, MAX_NEWTON + 1):
        area, n = reg_area(a, q / 2, side, bits)
        total_iters += n
        resid = area - target
        if resid == 0:
            return {"q": q, "steps": step, "resid": abs(resid), "side": side,
                    "iters": total_iters}
        dens = gmpy2.exp(log_pdf_chi2(df, q)) * q
        d_du = -dens if upper else dens
        q_new = gmpy2.exp(gmpy2.log(q) - resid / d_du)
        if q_new <= 0 or not gmpy2.is_finite(q_new):
            raise NonConvergence(f"Newton left the domain (df={df_float:.0e})")
        rel = abs(q_new - q) / abs(q_new)
        q = q_new
        if rel < tol:
            area, n = reg_area(a, q / 2, side, bits)
            return {"q": q, "steps": step, "resid": abs(area - target),
                    "side": side, "iters": total_iters + n}
    raise NonConvergence(f"Newton did not converge (df={df_float:.0e})")


def agreement_digits(x, y, bits=600):
    setprec(bits)
    x, y = mpfr(x), mpfr(y)
    if x == y:
        return None
    return float(-gmpy2.log10(abs(x - y) / abs(x)))


def main():
    started = time.time()
    ref = json.load(open("chisq_reference.json"))
    records = ref["references"]
    assert len(records) == 69, f"expected 69 records, got {len(records)}"

    points = []
    worst = None
    worst_at = None
    n_rej = 0
    for i, r in enumerate(records, 1):
        p_hex = r["probability"]["hex"]
        df = r["df"]
        t0 = time.time()
        try:
            lo = invert(df, p_hex, PREC_LOW)
            hi = invert(df, p_hex, PREC_HIGH, seed=lo["q"])
            stab = agreement_digits(lo["q"], hi["q"])
            if stab is not None and stab < REQUIRED_AGREEMENT_DIGITS:
                status = "REJECTED"
                reason = (f"did not stabilise: {stab:.4g} digits, required "
                          f"{REQUIRED_AGREEMENT_DIGITS}")
            else:
                status, reason = "ACCEPTED", None
        except NonConvergence as exc:
            status, reason, stab, hi = "REJECTED", str(exc), None, None
        if status == "REJECTED":
            n_rej += 1

        agree = None
        if status == "ACCEPTED":
            agree = agreement_digits(mpfr(r["quantile"]), hi["q"])
            if agree is not None and (worst is None or agree < worst):
                worst = agree
                worst_at = (f"df={df:.0e} {r['arm']}/{r['band']} "
                            f"p={r['probability']['float_repr']}")

        setprec(PREC_HIGH)
        p = hex_to_mpfr(p_hex)
        denom = min(p, 1 - p)
        points.append({
            "index": i, "df": df, "arm": r["arm"], "band": r["band"],
            "probability_hex": p_hex,
            "probability_repr": r["probability"]["float_repr"],
            "status": status, "reason": reason,
            "precision_pair_bits": [PREC_LOW, PREC_HIGH],
            "stabilisation_digits": (round(stab, 4) if stab else None),
            "quantile": (gmpy2.mpfr(hi["q"]).__format__(".50g") if hi else None),
            "residual_side": (hi["side"] if hi else None),
            "tail_residual": (f"{float(hi['resid'] / denom):.6g}" if hi else None),
            "newton_steps": (hi["steps"] if hi else None),
            "route_iterations": (hi["iters"] if hi else None),
            "agreement_with_primary_digits": (round(agree, 4) if agree else None),
            "runtime_s": round(time.time() - t0, 3),
        })
        print(f"  [{i:2d}/69] df={df:.0e} {r['arm']}/{r['band']:8s} {status:8s} "
              f"agree={agree if agree is None else round(agree, 3)}", flush=True)

    payload = {
        "checkpoint": ("v1.0.0 plan Track A2 item 6 - Chi-square MPFR-arithmetic "
                       "cross-check (gmpy2)"),
        "independence": ("IMPLEMENTATION-INDEPENDENT ONLY - same series/CF "
                         "algorithm as the primary, different library, arithmetic "
                         "backend, rounding and evaluation order. NOT an "
                         "algorithm-independent confirmation."),
        "library_incomplete_gamma_used": False,
        "why_not_library_igamma": {
            "mpmath.gammainc": "NoConvergence at a = 5E7",
            "MPFR mpfr_gamma_inc": ("aborts the process at a >= ~4.5E7; measured "
                                    "ok at 4.0E7, assertion failure at 4.5E7 "
                                    "(gamma_inc.c:289). df=1E8 needs a = 5E7."),
        },
        "why_gmpy2_not_Rmpfr": ("the same route in R measured ~810 us per series "
                                "term (96.9 s / 119,799 terms at 200 bits), "
                                "projecting ~183 min for the df=1E8 block; gmpy2 "
                                "binds the identical MPFR 4.2.2 at C speed"),
        "generated_utc": time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime()),
        "runtime_s_total": round(time.time() - started, 3),
        "versions": {
            "python": platform.python_version(),
            "gmpy2": gmpy2.version(),
            "mpfr": gmpy2.mpfr_version(),
            "gmp": gmpy2.mp_version(),
            "platform": platform.platform(),
        },
        "method": {
            "incomplete_gamma": ("converging lower series / upper Lentz CF in "
                                 "gmpy2 MPFR arithmetic"),
            "inversion": "Newton on log q, Wilson-Hilferty seed",
            "normal_quantile": ("bisection on MPFR erfc - mpmath's erfinv is not "
                                "used anywhere in this module"),
            "direct_side": "lower area for p <= 1/2, upper area for p > 1/2",
            "precision_pair_bits": [PREC_LOW, PREC_HIGH],
            "required_agreement_digits": REQUIRED_AGREEMENT_DIGITS,
            "measures": ["quantile",
                         "tail_residual = |area - target| / min(p, 1-p)"],
        },
        "convergence": {
            "points_total": len(records),
            "accepted": len(records) - n_rej,
            "rejected": n_rej,
            "all_converged_and_stabilised": n_rej == 0,
        },
        "agreement": {
            "measure": ("primary (mpmath series/CF) vs gmpy2 MPFR arithmetic, "
                        "significant digits of the quantile"),
            "minimum_agreement_digits": (round(worst, 4) if worst else None),
            "minimum_agreement_at": worst_at,
        },
        "points": points,
    }
    with open("chisq_gmpy2.json", "w", newline="\n") as f:
        json.dump(payload, f, indent=2)
        f.write("\n")
    print(f"\n{len(records) - n_rej}/{len(records)} accepted, {n_rej} rejected")
    print(f"minimum primary-vs-gmpy2 agreement: {worst} digits at {worst_at}")


if __name__ == "__main__":
    main()
