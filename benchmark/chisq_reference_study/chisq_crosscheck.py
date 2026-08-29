"""
Independent cross-check of the frozen Chi-square reference set.

THREE LEGS, with deliberately different standing:

  1. INDEPENDENT (evidential). The quantile is re-derived from scratch by a
     route sharing no code path with the primary: the density is integrated by
     tanh-sinh quadrature under a variance-stabilizing substitution, and the
     quantile is found by Newton seeded independently from Wilson-Hilferty. The
     primary instead sums a converging power series / Lentz continued fraction.
     Different operation, different error mechanism.

     The small side is always integrated DIRECTLY - lower area for p <= 1/2,
     upper area for p > 1/2 - so no complement subtraction enters the compared
     quantity.

     mpmath's own mp.gammainc was tried first as the independent leg and is NOT
     usable here: it raises NoConvergence at a = 5E7, i.e. exactly the shapes
     this checkpoint targets. That failure is recorded, not hidden, because it
     is the reason the independent leg is quadrature.

  2. COARSE SANITY (not evidential). SciPy chi2.ppf in double precision, and
     SymPy used to confirm each probability's exact binary64 rational value.
     Neither can resolve 40+ digits; they catch a gross error such as a wrong
     df convention or a factor of two. No agreement figure is derived from
     them. SymPy is deliberately NOT asked to evaluate lowergamma at these
     shapes - it does not return in usable time.

  3. Rmpfr - NOT RUN HERE. See rmpfr_crosscheck.R. Until it runs and is
     verified by the maintainer, the checkpoint stays PROVISIONAL and no final
     agreement figure is recorded.
"""
import glob
import json
import platform
import time

import mpmath
from mpmath import mp, mpf

from chisq_reference import (DPS_LOW, DPS_HIGH, REQUIRED_AGREEMENT_DIGITS,
                             DF_VALUES, ARMS, exact_mpf, agreement_digits,
                             wilson_hilferty_start, log_pdf_chi2,
                             NonConvergence)

QUAD_DPS = 110
QUAD_SIGMA_WINDOW = 70
MAX_NEWTON_STEPS = 100


def _integrand(a, sa):
    def f(s):
        t = a + s * sa
        if t <= 0:
            return mpf(0)
        return mp.e ** ((a - 1) * mp.log(t) - t - mp.loggamma(a)) * sa
    return f


def quadrature_area(a, x, side):
    """Regularized lower ('P') or upper ('Q') area by tanh-sinh quadrature.

    Substitutes t = a + s*sqrt(a) so the peak sits near s = 0 and the integrand
    is O(1)-scaled, which is what lets the rule converge at large a.

    The window is +-70 sqrt(a). Every frozen point lies within ~7 sqrt(a) of a,
    and the neglected tail beyond the window is of order exp(-70^2/2) ~ 1E-1064,
    far below anything reportable at this precision.
    """
    sa = mp.sqrt(a)
    f = _integrand(a, sa)
    hi = (x - a) / sa
    W = mpf(QUAD_SIGMA_WINDOW)
    if side == "P":
        lo = -W
        if hi <= lo:
            return mpf(0)
        pts = sorted({lo, mpf(-20), mpf(0), hi} if hi > 0 else {lo, mpf(-20), hi})
    else:
        if hi >= W:
            return mpf(0)
        pts = sorted({hi, mpf(0), mpf(20), W} if hi < 0 else {hi, mpf(20), W})
    return mp.quad(f, pts)


def independent_invert(df, p_exact, dps=QUAD_DPS):
    """Re-derive the quantile by the quadrature route, seeded independently."""
    mp.dps = dps
    a = mpf(df) / 2
    upper = p_exact > mpf(1) / 2
    side = "Q" if upper else "P"
    target = (1 - p_exact) if upper else p_exact

    q = wilson_hilferty_start(mpf(df), p_exact)
    tol = mpf(10) ** (-(dps - 8))
    for step in range(1, MAX_NEWTON_STEPS + 1):
        area = quadrature_area(a, q / 2, side)
        resid = area - target
        if resid == 0:
            return q, step, area
        dens = mp.e ** log_pdf_chi2(mpf(df), q) * q
        d_du = -dens if upper else dens
        q_new = mp.e ** (mp.log(q) - resid / d_du)
        if q_new <= 0 or not mp.isfinite(q_new):
            raise NonConvergence(
                f"independent Newton left the domain at df={df:.0e}")
        rel = abs(q_new - q) / abs(q_new)
        q = q_new
        if rel < tol:
            return q, step, quadrature_area(a, q / 2, side)
    raise NonConvergence(
        f"independent Newton did not converge in {MAX_NEWTON_STEPS} steps "
        f"(df={df:.0e})")


def coarse_sanity(df, p, q_ref):
    """Cheap, non-evidential. Gross-error detection only."""
    out = {"standing": "coarse only - not evidential"}
    try:
        from scipy import stats
        q_sp = float(stats.chi2.ppf(p, df))
        out["scipy"] = {"quantile": q_sp,
                        "rel_diff": abs(q_sp - float(q_ref)) / abs(float(q_ref))}
    except Exception as exc:
        out["scipy"] = {"error": f"{type(exc).__name__}: {exc}"}
    try:
        import sympy
        r = sympy.Rational(p)
        out["sympy"] = {"exact_binary64_rational": f"{r.p}/{r.q}",
                        "is_dyadic": bool((r.q & (r.q - 1)) == 0)}
    except Exception as exc:
        out["sympy"] = {"error": f"{type(exc).__name__}: {exc}"}
    return out


def builtin_probe():
    """Record that mpmath's own gammainc cannot serve as the independent leg."""
    results = {}
    for df in DF_VALUES:
        mp.dps = 60
        a = mpf(df) / 2
        t0 = time.time()
        try:
            v = mp.gammainc(a, 0, a, regularized=True)
            results[f"{df:.0e}"] = {"status": "ok", "P": mp.nstr(v, 20),
                                    "runtime_s": round(time.time() - t0, 3)}
        except Exception as exc:
            results[f"{df:.0e}"] = {
                "status": "FAILED",
                "error": f"{type(exc).__name__}: {str(exc)[:120]}",
                "runtime_s": round(time.time() - t0, 3)}
    return results


def load_records():
    """Load the frozen references.

    Prefers the merged chisq_reference.json written by chisq_reference.py.
    Falls back to part_*.json chunks, which exist only when generation was run
    one df at a time to fit a runtime limit.
    """
    records = []
    if glob.glob("chisq_reference.json"):
        payload = json.load(open("chisq_reference.json"))
        records = payload["references"]
    else:
        for path in sorted(glob.glob("part_*.json")):
            records.extend(json.load(open(path)))
    if not records:
        raise SystemExit(
            "no references found: run `python chisq_reference.py` first to "
            "write chisq_reference.json")
    order, i = {}, 0
    for df in DF_VALUES:
        for arm, band, probs in ARMS:
            for p in probs:
                order[(df, float.hex(p), arm, band)] = i
                i += 1
    records.sort(key=lambda r: order[(r["df"], r["probability"]["hex"],
                                      r["arm"], r["band"])])
    return records


def main():
    started = time.time()
    records = load_records()
    assert len(records) == 69, f"expected 69 records, got {len(records)}"

    comparisons = []
    worst = None
    worst_at = None
    worst_scipy = 0.0
    for n, rec in enumerate(records, 1):
        if rec["status"] != "ACCEPTED":
            comparisons.append({"df": rec["df"], "arm": rec["arm"],
                                "band": rec["band"],
                                "probability_hex": rec["probability"]["hex"],
                                "status": "SKIPPED - primary rejected"})
            continue
        df = rec["df"]
        p_float = float.fromhex(rec["probability"]["hex"])
        mp.dps = QUAD_DPS
        p_exact = exact_mpf(p_float)
        q_prim = mpf(rec["quantile"])

        t0 = time.time()
        try:
            q_ind, steps, area = independent_invert(df, p_exact)
        except NonConvergence as exc:
            comparisons.append({"df": df, "arm": rec["arm"], "band": rec["band"],
                                "probability_hex": rec["probability"]["hex"],
                                "status": "INDEPENDENT_FAILED",
                                "error": str(exc)})
            print(f"  [{n:2d}/69] df={df:.0e} INDEPENDENT_FAILED", flush=True)
            continue
        runtime = time.time() - t0

        digits = agreement_digits(q_prim, q_ind)
        mp.dps = QUAD_DPS
        d = None if digits == mp.inf else float(mp.nstr(digits, 6))
        if d is not None and (worst is None or d < worst):
            worst = d
            worst_at = f"df={df:.0e} {rec['arm']}/{rec['band']} p={p_float!r}"

        sanity = coarse_sanity(df, p_float, q_prim)
        rd = (sanity.get("scipy") or {}).get("rel_diff")
        if rd:
            worst_scipy = max(worst_scipy, rd)

        comparisons.append({
            "df": df, "arm": rec["arm"], "band": rec["band"],
            "probability_hex": rec["probability"]["hex"],
            "probability_repr": rec["probability"]["float_repr"],
            "status": "COMPARED",
            "compared_measure": "quantile",
            "compared_side": "Q" if p_exact > mpf(1) / 2 else "P",
            "primary_quantile": mp.nstr(q_prim, 40),
            "independent_quantile": mp.nstr(q_ind, 40),
            "agreement_significant_digits": d,
            "independent_newton_steps": steps,
            "independent_runtime_s": round(runtime, 3),
            "coarse_sanity": sanity,
        })
        print(f"  [{n:2d}/69] df={df:.0e} {rec['arm']}/{rec['band']:8s} "
              f"agree={d}", flush=True)

    payload = {
        "checkpoint": "v1.0.0 plan Track A2 item 6 - Chi-square reference feasibility",
        "status": "PROVISIONAL - INCOMPLETE",
        "completion_blocked_on": (
            "the independent Rmpfr leg (rmpfr_crosscheck.R) has not been run or "
            "verified by the maintainer. No final agreement figure is recorded "
            "and this checkpoint must NOT be marked complete."),
        "generated_utc": time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime()),
        "runtime_s_total": round(time.time() - started, 3),
        "versions": {"python": platform.python_version(),
                     "mpmath": mpmath.__version__,
                     "platform": platform.platform()},
        "legs": {
            "primary": {
                "method": "converging lower series / upper Lentz continued fraction",
                "precision_pair": [DPS_LOW, DPS_HIGH],
                "required_agreement_digits": REQUIRED_AGREEMENT_DIGITS,
                "standing": "evidential",
            },
            "independent": {
                "method": ("tanh-sinh quadrature of the density under "
                           "t = a + s*sqrt(a), independent Wilson-Hilferty seed, "
                           "Newton on log q, small side integrated directly"),
                "precision_dps": QUAD_DPS,
                "standing": "evidential - algorithmically independent",
            },
            "mpmath_builtin_probe": {
                "method": "mp.gammainc regularized",
                "standing": ("REJECTED as independent leg - does not converge at "
                             "these shapes"),
                "results": builtin_probe(),
            },
            "coarse_sanity": {
                "method": ("scipy.stats.chi2.ppf (double) for the quantile; sympy "
                           "for the exact binary64 rational of each probability"),
                "standing": "NOT evidential - gross-error detection only",
                "worst_scipy_relative_difference": worst_scipy,
            },
            "rmpfr": {
                "method": "Rmpfr igamma - see rmpfr_crosscheck.R",
                "standing": "REQUIRED - NOT YET RUN",
                "result": None,
            },
        },
        "summary": {
            "points_total": len(records),
            "points_compared": sum(1 for c in comparisons
                                   if c["status"] == "COMPARED"),
            "primary_vs_independent_min_agreement_digits": worst,
            "primary_vs_independent_min_agreement_at": worst_at,
            "final_agreement_figure": None,
            "final_agreement_figure_note": (
                "deliberately null: the final figure is primary-vs-Rmpfr "
                "agreement, which cannot be computed until the Rmpfr leg runs"),
        },
        "comparisons": comparisons,
    }
    with open("chisq_crosscheck.json", "w", newline="\n") as f:
        json.dump(payload, f, indent=2)
        f.write("\n")
    print(f"\ncompared {payload['summary']['points_compared']}/69")
    print(f"primary vs independent quadrature: min agreement {worst} digits "
          f"(at {worst_at})")
    print(f"coarse scipy worst relative difference: {worst_scipy:.3e}")
    print("Rmpfr leg NOT run - checkpoint remains PROVISIONAL.")


if __name__ == "__main__":
    main()
