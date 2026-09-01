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

  3. Rmpfr - RUN AND RECORDED, in chisq_rmpfr_1e6.json and
     chisq_rmpfr_1e7.json. Native igamma covers df = 1E6 and 1E7 only, because
     mpfr_gamma_inc aborts above shape ~4.5E7 and df = 1E8 needs 5E7. The
     df = 1E8 block is covered by the accepted gmpy2/MPFR substitution
     (chisq_gmpy2.json) plus the quadrature leg above.

This module also derives the COMPOSITE AUTHORITY from the committed leg
records. Every figure is recomputed from those files on each run - none is
hand-written - so the decision record cannot be silently invalidated by
re-running the script. The loader fails closed on missing or malformed
inputs, rejected or unstabilised points, duplicate or foreign identities,
unexpected degrees of freedom, inconsistent precision pairs, and incomplete
coverage, so a partial record can never be reported as a complete one.
"""
import glob
import json
import os
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


# ---------------------------------------------------------------------------
# Composite authority
#
# The authority is DERIVED from the committed leg records, never hand-written.
# A decision record a script can silently overwrite is not a record, so the
# figures below are recomputed on every run and the loader fails closed rather
# than reporting a partial or mismatched picture.
# ---------------------------------------------------------------------------

# Leg record files, and the identity each is required to have. The expected df
# coverage is asserted, not inferred, so a truncated or mislabelled record is a
# hard failure instead of a quietly smaller number.
LEG_SPECS = (
    ("gmpy2", "chisq_gmpy2.json", None,
     "agreement_with_primary_digits", "implementation-independent"),
    ("rmpfr_1e6", "chisq_rmpfr_1e6.json", 1e6,
     "agreement_with_python_primary_digits", "third-party incomplete gamma"),
    ("rmpfr_1e7", "chisq_rmpfr_1e7.json", 1e7,
     "agreement_with_python_primary_digits", "third-party incomplete gamma"),
)

# Measured by rmpfr_probe.R: igamma succeeds at shape 4.0E7 and aborts at
# 4.5E7. This is the one figure not derived from a machine-readable record,
# because an MPFR abort kills the process before anything can be written. It is
# a recorded measurement, not an estimate.
RMPFR_CEILING = {
    "highest_shape_ok": 4.0e7,
    "lowest_shape_aborting": 4.5e7,
    "df_required_for_1e8": 5.0e7,
    "failure": "gamma_inc.c:289 MPFR assertion failed - aborts the process",
    "source": "rmpfr_probe.R (bisection; an abort cannot be caught or logged)",
}


class LegError(Exception):
    """A leg record is missing, malformed, or does not match the frozen set."""


def _reference_identities(records):
    """(df, probability_hex) for every frozen reference, as the identity key."""
    ids = {}
    for r in records:
        key = (float(r["df"]), r["probability"]["hex"])
        if key in ids:
            raise LegError(f"duplicate identity in the reference set: {key}")
        ids[key] = r
    return ids


def load_leg(name, path, expected_df, agree_key, standing, ref_ids):
    """Load one leg record, failing closed on anything that would make the
    composite authority overstate the evidence."""
    if not os.path.exists(path):
        raise LegError(f"{name}: leg record {path!r} is missing")
    try:
        with open(path, encoding="utf-8") as f:
            doc = json.load(f)
    except (ValueError, OSError) as exc:
        raise LegError(f"{name}: {path!r} is malformed: {exc}")

    points = doc.get("points")
    if not isinstance(points, list) or not points:
        raise LegError(f"{name}: {path!r} has no points array")

    seen = set()
    agreements = []
    pairs = set()
    for i, p in enumerate(points):
        for field in ("df", "probability_hex", "status", "precision_pair_bits"):
            if field not in p:
                raise LegError(f"{name}: point {i} is missing {field!r}")
        if p["status"] != "ACCEPTED":
            raise LegError(
                f"{name}: point {i} ({p['probability_hex']}) has status "
                f"{p['status']!r}: {p.get('reason')}")
        # R writes df as an integer, Python as a float; normalise before
        # comparing so an identity match is not defeated by JSON typing.
        df = float(p["df"])
        if expected_df is not None and df != expected_df:
            raise LegError(
                f"{name}: point {i} has df={df:.0e}, expected "
                f"{expected_df:.0e}")
        key = (df, p["probability_hex"])
        if key in seen:
            raise LegError(f"{name}: duplicate point identity {key}")
        if key not in ref_ids:
            raise LegError(
                f"{name}: point {key} does not correspond to any frozen "
                "reference")
        seen.add(key)
        pairs.add(tuple(p["precision_pair_bits"]))
        a = p.get(agree_key)
        if a is None:
            raise LegError(
                f"{name}: point {i} carries no agreement figure under "
                f"{agree_key!r}")
        agreements.append(float(a))

    if len(pairs) != 1:
        raise LegError(f"{name}: inconsistent precision pairs across points: "
                       f"{sorted(pairs)}")
    pair = pairs.pop()

    expected_ids = {k for k in ref_ids
                    if expected_df is None or k[0] == expected_df}
    missing = expected_ids - seen
    if missing:
        raise LegError(
            f"{name}: incomplete coverage - {len(missing)} frozen point(s) "
            f"absent, e.g. {sorted(missing)[:2]}")

    conv = doc.get("convergence") or {}
    if conv.get("rejected"):
        raise LegError(f"{name}: record reports {conv['rejected']} rejected point(s)")

    runtime = doc.get("runtime_s_total")
    if runtime is None:
        raise LegError(f"{name}: record carries no runtime_s_total")

    return {
        "record": os.path.basename(path),
        "standing": standing,
        "points": len(seen),
        "df_values": sorted({k[0] for k in seen}),
        "precision_pair_bits": list(pair),
        "min_agreement_digits": round(min(agreements), 4),
        "max_agreement_digits": round(max(agreements), 4),
        "runtime_s": runtime,
        "versions": doc.get("versions"),
    }


def build_authority(records, quadrature, quad_runtime):
    """Derive the composite authority from the committed leg records."""
    ref_ids = _reference_identities(records)
    legs = {}
    for name, path, exp_df, agree_key, standing in LEG_SPECS:
        legs[name] = load_leg(name, path, exp_df, agree_key, standing, ref_ids)

    # Quadrature comes from this run rather than a committed record. Points
    # whose agreement is beyond measurement (an exact match at working
    # precision) carry None and are excluded from the minimum rather than
    # being counted as zero.
    quad_vals = [v for v in quadrature if v is not None]
    if not quad_vals:
        raise LegError("quadrature leg produced no measurable agreement")

    all_mins = [legs[n]["min_agreement_digits"] for n in legs]
    all_mins.append(round(min(quad_vals), 4))
    conservative = min(all_mins)

    rmpfr_points = legs["rmpfr_1e6"]["points"] + legs["rmpfr_1e7"]["points"]
    rmpfr_dfs = sorted(set(legs["rmpfr_1e6"]["df_values"] +
                           legs["rmpfr_1e7"]["df_values"]))
    covered_by_rmpfr = {k for k in ref_ids if k[0] in set(rmpfr_dfs)}
    not_covered = sorted({k[0] for k in ref_ids} - set(rmpfr_dfs))

    return {
        "status": "COMPLETE",
        "decision": (
            "df = 1E8 is covered by the accepted gmpy2/MPFR substitution plus "
            "algorithm-independent quadrature. Rmpfr remains the third-party "
            "check wherever its incomplete-gamma routine can execute."),
        "conservative_final_agreement_digits": conservative,
        "conservative_basis": (
            "the minimum agreement across every leg and every point; the "
            "maximum would flatter the result"),
        "legs": {
            "primary": {
                "record": "chisq_reference.json",
                "standing": "route under test",
                "points": len(records),
                "precision_pair_dps": [DPS_LOW, DPS_HIGH],
                "df_values": sorted({float(r["df"]) for r in records}),
            },
            "quadrature": {
                "record": "chisq_crosscheck.json (this file)",
                "standing": "algorithm-independent",
                "points": len(quadrature),
                "points_measurable": len(quad_vals),
                "points_beyond_measurement": len(quadrature) - len(quad_vals),
                "precision_dps": QUAD_DPS,
                "min_agreement_digits": round(min(quad_vals), 4),
                "max_agreement_digits": round(max(quad_vals), 4),
                "runtime_s": round(quad_runtime, 3),
            },
            "gmpy2_mpfr": legs["gmpy2"],
            "rmpfr_igamma": {
                "records": [legs["rmpfr_1e6"]["record"],
                            legs["rmpfr_1e7"]["record"]],
                "standing": "third-party incomplete gamma",
                "points": rmpfr_points,
                "points_feasible_of_total": f"{rmpfr_points}/{len(ref_ids)}",
                "df_values": rmpfr_dfs,
                "df_not_covered": not_covered,
                "precision_pair_bits": legs["rmpfr_1e6"]["precision_pair_bits"],
                "min_agreement_digits": min(
                    legs["rmpfr_1e6"]["min_agreement_digits"],
                    legs["rmpfr_1e7"]["min_agreement_digits"]),
                "max_agreement_digits": max(
                    legs["rmpfr_1e6"]["max_agreement_digits"],
                    legs["rmpfr_1e7"]["max_agreement_digits"]),
                "runtime_s_by_df": {
                    f"{legs['rmpfr_1e6']['df_values'][0]:.0e}":
                        legs["rmpfr_1e6"]["runtime_s"],
                    f"{legs['rmpfr_1e7']['df_values'][0]:.0e}":
                        legs["rmpfr_1e7"]["runtime_s"],
                },
                "runtime_s_combined": round(
                    legs["rmpfr_1e6"]["runtime_s"] +
                    legs["rmpfr_1e7"]["runtime_s"], 3),
                "coverage_limit": RMPFR_CEILING,
            },
        },
        "rmpfr_coverage_check": {
            "frozen_points": len(ref_ids),
            "covered_by_rmpfr": len(covered_by_rmpfr),
            "uncovered_df": not_covered,
        },
    }


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
    # Track measurements and failures separately. An absent SciPy must never
    # be reported as a worst-case difference of zero: "not measured" and
    # "agreed perfectly" are opposite claims, and the initial 0.0 renders the
    # first as the second.
    scipy_diffs = []
    scipy_errors = []
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
        sc = sanity.get("scipy") or {}
        if "rel_diff" in sc:
            scipy_diffs.append(sc["rel_diff"])
        else:
            scipy_errors.append(sc.get("error", "unknown"))

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

    quad_runtime = time.time() - started
    quad_agreements = [c.get("agreement_significant_digits")
                       for c in comparisons if c["status"] == "COMPARED"]
    authority = build_authority(records, quad_agreements, quad_runtime)

    payload = {
        "checkpoint": "v1.0.0 plan Track A2 item 6 - Chi-square reference feasibility",
        "status": "COMPLETE",
        "composite_authority": authority,
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
                "scipy_available": bool(scipy_diffs),
                "scipy_points_measured": len(scipy_diffs),
                "scipy_points_failed": len(scipy_errors),
                "scipy_first_error": (scipy_errors[0] if scipy_errors else None),
                "worst_scipy_relative_difference": (
                    max(scipy_diffs) if scipy_diffs else None),
                "worst_scipy_note": (
                    None if scipy_diffs else
                    "SciPy was not available; the leg did NOT run. This is null "
                    "rather than 0.0 because not measured and agreed perfectly "
                    "are opposite claims."),
            },
            "rmpfr": {
                "method": "Rmpfr igamma - see rmpfr_crosscheck.R",
                "standing": "third-party incomplete gamma",
                "result": ("run and recorded in chisq_rmpfr_1e6.json and "
                           "chisq_rmpfr_1e7.json; see composite_authority"),
            },
            "gmpy2_mpfr": {
                "method": ("converging series / Lentz CF in gmpy2 MPFR "
                           "arithmetic - see chisq_gmpy2.py"),
                "standing": ("implementation-independent: same algorithm as the "
                             "primary, different library, arithmetic backend, "
                             "rounding and evaluation order"),
                "result": "recorded in chisq_gmpy2.json; see composite_authority",
            },
        },
        "summary": {
            "points_total": len(records),
            "points_compared": sum(1 for c in comparisons
                                   if c["status"] == "COMPARED"),
            "primary_vs_independent_min_agreement_digits": worst,
            "primary_vs_independent_min_agreement_at": worst_at,
            "final_agreement_figure": authority[
                "conservative_final_agreement_digits"],
            "final_agreement_figure_note": (
                "conservative: the minimum across every leg and every point, "
                "derived from the committed leg records"),
        },
        "comparisons": comparisons,
    }
    with open("chisq_crosscheck.json", "w", newline="\n") as f:
        json.dump(payload, f, indent=2)
        f.write("\n")
    print(f"\ncompared {payload['summary']['points_compared']}/69")
    print(f"primary vs independent quadrature: min agreement {worst} digits "
          f"(at {worst_at})")
    if scipy_diffs:
        print(f"coarse scipy worst relative difference: {max(scipy_diffs):.3e} "
              f"({len(scipy_diffs)}/69 measured)")
    else:
        print(f"coarse scipy: NOT MEASURED - SciPy unavailable "
              f"({scipy_errors[0] if scipy_errors else 'unknown'}); "
              f"recorded as null, not zero")
    a = authority
    print(f"\ncomposite authority: {a['status']}")
    for key in ("primary", "quadrature", "gmpy2_mpfr", "rmpfr_igamma"):
        leg = a["legs"][key]
        print(f"  {key:14s} points={leg.get('points')} "
              f"min={leg.get('min_agreement_digits')}")
    print(f"conservative final agreement: "
          f"{a['conservative_final_agreement_digits']} significant digits")


if __name__ == "__main__":
    main()
