"""
Measure the balanced Beta_Survival route against binary64-argument references
(#163), and write results.json.

Sections of results.json:

  binding       SHA-256 of the VBA procedures beta_route_model.py ports.
  semantics     PROB_LogGamma observations reproduced under plain binary64
                and under 64-bit intermediates; the model uses the second.
  fidelity      the model against every committed Excel observation it can
                evaluate: LogGamma, LogGammaHalfDiff, PROB_LogBeta, LogBeta,
                Beta_Survival and Beta_Cumulative, in every grid.
  observations  every committed Beta_Survival and Beta_Cumulative observation
                with min(a, b) / max(a, b) >= 0.1 and both shapes below 100,
                with its error against the committed reference (the gate's
                arithmetic) and against a reference recomputed here at the
                binary64 argument. Rows outside that range are counted, by
                reason, under "excluded".
  sweeps        the model on deterministic point sets, against certified
                references, in three variants:
                  source     the route as the source has it;
                  exact_lb   PROB_LogBeta replaced by the correctly rounded
                             Log(Beta), every other statement unchanged;
                  candidate  PROB_LogGamma with full-precision Lanczos
                             coefficients (CANDIDATE_LANCZOS), the repair
                             candidate.
                Points: "local" around (0.82, 2.7, 3.4), "ulp_neighbours" of
                x = 0.82, "domain" across shapes 0.5 to 19, and "large" with
                both shapes 20 to 100 at the mean and +-1, 2, 3 sd.

References. Survival: _survival_oracle.certify (I_{1-x}(b, a) at the exact
binary64 arguments, two working precisions agreeing to 32 digits).
Cumulative: I_x(a, b) by the same construction (cdf_reference below).

    python measure_balanced_survival.py           write results.json
    python measure_balanced_survival.py --check   recompute; fail if
                                                  results.json differs
"""
import argparse
import csv
import json
import math
import os
import sys
from decimal import Decimal, getcontext

import mpmath as mp

HERE = os.path.dirname(os.path.abspath(__file__))
BENCHMARK = os.path.dirname(HERE)
sys.path.insert(0, BENCHMARK)
sys.path.insert(0, HERE)
import _survival_oracle as O                                      # noqa: E402
import beta_route_model as M                                      # noqa: E402

RESULTS = os.path.join(HERE, "results.json")
THRESHOLD = 5e-15                  # Beta_Survival.balanced.output_rel
CDF_THRESHOLD = 2e-14              # Beta_Cumulative.balanced.output_rel
BANDS = (5e-15, 1e-14, 2e-14, 5e-14)
FIDELITY_LIMIT_ULPS = 2

# The g = 7, n = 9 Lanczos coefficients at 17 significant digits. Each rounds
# to the 15-digit literal the source declares; the source therefore carries
# truncations of up to 2.7E-12 (PROB_LANCZOS_P2).
CANDIDATE_LANCZOS = (
    0.99999999999980993, 676.5203681218851, -1259.1392167224028,
    771.32342877765313, -176.61502916214059, 12.507343278686905,
    -0.13857109526572012, 9.9843695780195716e-6, 1.5056327351493116e-7,
)

GRIDS = (
    "probability_accuracy_grid.csv",
    "holdout/holdout_grid.csv",
    "holdout_older/holdout_older_grid.csv",
    "beta_f_unbalanced/beta_f_unbalanced_grid.csv",
    "cdf_large_shape/cdf_large_shape_grid.csv",
    "logbeta_study/logbeta_switch_grid.csv",
    "loggamma_regimes_study/loggamma_regimes_grid.csv",
    "density_large_shape/density_large_shape_grid.csv",
)
EXCEL_COMPARISON = "excel_comparison/excel_comparison_grid.csv"

getcontext().prec = 60


# --------------------------------------------------------------------------
# helpers


def ordered(v):
    """Ordered-integer image of a Double, monotone across the sign."""
    import struct
    i = struct.unpack(">q", struct.pack(">d", v))[0]
    return i if i >= 0 else -(i & 0x7FFFFFFFFFFFFFFF)


def ulp_distance(a, b):
    return abs(ordered(a) - ordered(b))


def parse_observed(raw):
    return sum((Decimal(p.strip()) for p in raw.split(";")), Decimal(0))


def sig(v, digits=3):
    """A float as a short decimal string, for a stable results file."""
    return float(f"{float(v):.{digits}e}")


def cdf_reference(x, a, b):
    """I_x(a, b) at exact binary64 arguments, certified like the survival."""
    args = O.exact_args([repr(x), repr(a), repr(b)])
    extra = O.magnitude_digits(args)

    def at(dps):
        with mp.workdps(dps):
            xv, av, bv = (+v for v in args)
            if min(av, bv) >= O.QUAD_MIN_SHAPE:
                return +O.beta_lower_quad(xv, av, bv)
            return +mp.betainc(av, bv, 0, xv, regularized=True)

    lo, hi = at(O.DPS_LOW + extra), at(O.DPS_HIGH + extra)
    with mp.workdps(O.DPS_HIGH + extra):
        if hi != 0 and abs(hi - lo) / abs(hi) > mp.mpf(10) ** -O.CERTIFY_DIGITS:
            raise O.OracleError(f"cdf reference unstable at {(x, a, b)}")
    return hi


def survival_reference(x, a, b):
    return O.certify("Beta_Survival", [repr(x), repr(a), repr(b)])[0]


def rel(value, reference):
    with mp.workdps(60):
        return float(abs(mp.mpf(value) - reference) / abs(reference))


def read_rows(rel_path):
    with open(os.path.join(BENCHMARK, rel_path), newline="", encoding="utf-8") as f:
        return list(csv.DictReader(f))


def excel_comparison_rows():
    """K_STATS_Beta_* rows of the pipe-delimited excel_comparison grid."""
    path = os.path.join(BENCHMARK, EXCEL_COMPARISON)
    with open(path, newline="", encoding="utf-8") as f:
        rows = list(csv.DictReader(f, delimiter="|"))
    out = []
    for r in rows:
        call = r["kstats_call"]
        for fn in ("Beta_Survival", "Beta_Cumulative"):
            if call.startswith(f"K_STATS_{fn}("):
                args = call[len(f"K_STATS_{fn}("):-1].split(",")
                out.append({"function": fn, "arg1": args[0], "arg2": args[1],
                            "arg3": args[2], "reference": r["reference"],
                            "observed_vba": r["observed_kstats"],
                            "regime": "", "label": r["label"]})
    return out


# --------------------------------------------------------------------------
# sections


def semantics():
    """Which arithmetic the LogGamma observations follow."""
    out = {}
    rows = [r for g in GRIDS for r in read_rows(g)
            if (r.get("function") or r.get("quantity")) == "LogGamma"
            and (r.get("observed_vba") or "").strip()
            and float(r["arg1"]) >= 0.5]
    for bits in (53, 64):
        exact = 0
        with M.intermediate_precision(bits):
            for r in rows:
                obs = float(parse_observed(r["observed_vba"]))
                exact += M.log_gamma(float(r["arg1"])) == obs
        out[f"{bits}_bit_intermediates"] = {"observations": len(rows), "exact": exact}
    return out


FIDELITY_FUNCTIONS = {
    "LogGamma": lambda r: M.log_gamma(float(r["arg1"])),
    "LogGammaHalfDiff": lambda r: M.log_gamma_half_diff(float(r["arg1"])),
    "PROB_LogBeta": lambda r: M.log_beta(float(r["arg1"]), float(r["arg2"])),
    "LogBeta": lambda r: M.log_beta(float(r["arg1"]), float(r["arg2"])),
    "Beta_Survival": lambda r: M.beta_survival(*(float(r[k]) for k in ("arg1", "arg2", "arg3"))),
    "Beta_Cumulative": lambda r: M.beta_cumulative(*(float(r[k]) for k in ("arg1", "arg2", "arg3"))),
    "Beta_Density": lambda r: M.beta_density(*(float(r[k]) for k in ("arg1", "arg2", "arg3"))),
    "StirlingError": lambda r: M.stirling_error(float(r["arg1"])),
}


# Committed observations exported before a change to a procedure they reach.
# The model ports today's source, so it is not compared with them; they are
# counted under "fidelity.predates_source". Dates are those of the commits.
PREDATES = {
    "beta_f_unbalanced": "exported 2026-07-19, before the Loader dispatch of "
                         "PROB_TryBetaRegularized (2026-07-29) and the later changes "
                         "to PROB_TryBetaLogPdf, PROB_TryDeviancePart and "
                         "PROB_StirlingError",
}
# holdout_older was re-exported on 2026-11-06 (#81, run 42210983927) and is
# compared like every other grid since.


def reaches_loader(q, r):
    """True when the row reaches PROB_TryBetaLogPdf or PROB_StirlingError."""
    if q in ("Beta_Density", "StirlingError"):
        return True
    if q in ("Beta_Survival", "Beta_Cumulative"):
        return float(r["arg2"]) + float(r["arg3"]) >= M.C["PROB_IBETA_LOADER_MIN_SHAPE"]
    return False


def predates(grid, q, r):
    if not reaches_loader(q, r):
        return None
    if grid.startswith("beta_f_unbalanced/") or r.get("evidence_set") == "beta_f_unbalanced":
        return "beta_f_unbalanced"
    return None


def fidelity():
    per, old = {}, {}
    refused = 0
    rows = [(g, r) for g in GRIDS for r in read_rows(g)]
    rows += [(EXCEL_COMPARISON, r) for r in excel_comparison_rows()]
    for g, r in rows:
        q = r.get("function") or r.get("quantity")
        raw = (r.get("observed_vba") or "").strip()
        if q not in FIDELITY_FUNCTIONS or not raw or "ERR" in raw.upper():
            continue
        try:
            modelled = FIDELITY_FUNCTIONS[q](r)
        except M.NotModelled:
            refused += 1
            continue
        d = ulp_distance(modelled, float(parse_observed(raw)))
        reason = predates(g, q, r)
        if reason:
            o = old.setdefault(reason, {"why": PREDATES[reason], "observations": 0,
                                        "exact": 0, "max_ulps": 0})
            o["observations"] += 1
            o["exact"] += d == 0
            o["max_ulps"] = max(o["max_ulps"], d)
            continue
        s = per.setdefault(q, {"observations": 0, "exact": 0, "max_ulps": 0})
        s["observations"] += 1
        s["exact"] += d == 0
        s["max_ulps"] = max(s["max_ulps"], d)
    return {"limit_ulps": FIDELITY_LIMIT_ULPS,
            "observations": sum(s["observations"] for s in per.values()),
            "exact": sum(s["exact"] for s in per.values()),
            "max_ulps": max(s["max_ulps"] for s in per.values()),
            "refused_not_modelled": refused, "by_function": dict(sorted(per.items())),
            "predates_source": dict(sorted(old.items()))}


def observations():
    rows, excluded = [], {}
    sources = [(g, r) for g in GRIDS[:5] for r in read_rows(g)]
    sources += [(EXCEL_COMPARISON, r) for r in excel_comparison_rows()]
    for grid, r in sources:
        fn = r.get("function")
        raw = (r.get("observed_vba") or "").strip()
        if fn not in ("Beta_Survival", "Beta_Cumulative") or not raw:
            continue
        x, a, b = (float(r[k]) for k in ("arg1", "arg2", "arg3"))
        if min(a, b) / max(a, b) < 0.1:
            reason = "ratio below 0.1"
        elif max(a, b) >= 100:
            reason = "a shape at or above 100 (large_shape and extreme_shape contracts)"
        elif max(a, b) < 0.5:
            reason = "both shapes below 0.5 (tiny_unbalanced contract)"
        else:
            reason = None
        if reason:
            excluded[reason] = excluded.get(reason, 0) + 1
            continue
        obs = parse_observed(raw)
        ref_cell = Decimal(r["reference"].strip())
        survival = fn == "Beta_Survival"
        reference = survival_reference(x, a, b) if survival else cdf_reference(x, a, b)
        t = {}
        (M.beta_survival if survival else M.beta_cumulative)(x, a, b, trace=t)
        rows.append({
            "grid": grid, "function": fn, "x": r["arg1"], "a": r["arg2"], "b": r["arg3"],
            "regime": r.get("regime") or "", "contract_threshold":
                (THRESHOLD if survival else CDF_THRESHOLD) if r.get("regime") == "balanced" else None,
            "value": sig(obs, 6),
            "gate_error": sig(abs(obs - ref_cell) / abs(ref_cell)),
            "binary64_error": rel(float(obs), reference),
            "branch": t["branch"],
        })
    for row in rows:
        row["binary64_error"] = sig(row["binary64_error"])
    return {"rows": rows, "excluded": dict(sorted(excluded.items()))}


def local_points():
    pts = []
    for i in range(26):
        x = round(0.70 + 0.01 * i, 2)
        for j in range(5):
            a = round(2.5 + 0.1 * j, 1)
            for k in range(5):
                pts.append((x, a, round(3.2 + 0.1 * k, 1)))
    return pts


def ulp_neighbour_points():
    pts = []
    for k in range(-8, 9):
        if k:
            x = 0.82
            step = math.inf if k > 0 else -math.inf
            for _ in range(abs(k)):
                x = math.nextafter(x, step)
            pts.append((x, 2.7, 3.4))
    return pts


DOMAIN_SHAPES = (0.5, 0.75, 1.0, 1.5, 2.0, 2.7, 3.4, 4.3, 5.0, 7.0, 10.0, 14.0, 19.0)
DOMAIN_X = (0.01, 0.05, 0.1, 0.2, 0.3, 0.4, 0.5, 0.6, 0.7, 0.8, 0.9, 0.95, 0.99)
LARGE_SHAPES = (20.0, 30.0, 50.0, 100.0)


def domain_points():
    return [(x, a, b) for a in DOMAIN_SHAPES for b in DOMAIN_SHAPES
            if min(a, b) / max(a, b) >= 0.1 for x in DOMAIN_X]


def large_points():
    pts = []
    for a in LARGE_SHAPES:
        for b in LARGE_SHAPES:
            mean = a / (a + b)
            sd = math.sqrt(a * b / ((a + b) ** 2 * (a + b + 1)))
            for k in (-3, -2, -1, 0, 1, 2, 3):
                pts.append((float(f"{mean + k * sd:.6f}"), a, b))
    return pts


VARIANTS = {
    "source": {},
    "exact_lb": {"exact_lb": True},
    "lanczos17": {"lanczos": CANDIDATE_LANCZOS},
    "loader_prefactor": {"prefactor": "loader"},
}
SHAPE_BANDS = ((0.5, 5.0), (5.0, 10.0), (10.0, 20.0), (20.0, 100.0))


def band(p):
    top = max(p[1], p[2])
    for lo, hi in SHAPE_BANDS:
        if top <= hi:
            return f"max shape <= {hi:g}"
    raise ValueError(p)


def summarize(errors, points, extra):
    order = sorted(range(len(errors)), key=lambda i: -errors[i])
    srt = sorted(errors)

    def q(p):
        return sig(srt[min(len(srt) - 1, int(p * len(srt)))])

    by_shape = {}
    for e, p in zip(errors, points):
        by_shape.setdefault(band(p), []).append(e)
    return {
        "points": len(errors),
        "max": sig(srt[-1]),
        "p50": q(0.50), "p90": q(0.90), "p99": q(0.99),
        "above": {f"{t:.0e}": sum(e > t for e in errors) for t in BANDS},
        "by_max_shape": {k: {"points": len(v), "max": sig(max(v)),
                             "above_5e-15": sum(e > THRESHOLD for e in v)}
                         for k, v in sorted(by_shape.items())},
        "worst": [dict(x=points[i][0], a=points[i][1], b=points[i][2],
                       error=sig(errors[i]), **extra[i]) for i in order[:5]],
    }


def modelled(p):
    try:
        M.beta_survival(*p)
        M.beta_cumulative(*p)
        return True
    except M.NotModelled:
        return False


def sweep(points):
    requested = len(points)
    points = [p for p in points if modelled(p)]
    refs_s = [survival_reference(*p) for p in points]
    refs_c = [cdf_reference(*p) for p in points]
    out = {"refused_not_modelled": requested - len(points)}
    for fn, model, refs in (("Beta_Survival", M.beta_survival, refs_s),
                            ("Beta_Cumulative", M.beta_cumulative, refs_c)):
        res = {}
        for name, kw in VARIANTS.items():
            errors, extra = [], []
            for p, ref in zip(points, refs):
                t = {}
                v = model(*p, trace=t, **kw)
                errors.append(rel(v, ref))
                extra.append({"value": sig(ref, 4), "branch": t["branch"]})
            res[name] = summarize(errors, points, extra)
            if name == "source":
                by = {}
                for e, x in zip(errors, extra):
                    by.setdefault(x["branch"], []).append(e)
                res[name]["by_branch"] = {
                    k: {"points": len(v), "max": sig(max(v)),
                        "above_5e-15": sum(e > THRESHOLD for e in v)}
                    for k, v in sorted(by.items())}
        out[fn] = res
    return out


def build():
    return {
        "issue": 163,
        "threshold": {"Beta_Survival.balanced.output_rel": THRESHOLD,
                      "Beta_Cumulative.balanced.output_rel": CDF_THRESHOLD},
        "binding": M.source_binding(),
        "semantics": semantics(),
        "fidelity": fidelity(),
        "observations": observations(),
        "sweeps": {
            "local": sweep(local_points()),
            "ulp_neighbours": sweep(ulp_neighbour_points()),
            "domain": sweep(domain_points()),
            "large": sweep(large_points()),
        },
    }


def main():
    ap = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    ap.add_argument("--check", action="store_true",
                    help="recompute and compare with results.json; write nothing")
    args = ap.parse_args()
    text = json.dumps(build(), indent=1, sort_keys=True) + "\n"
    if args.check:
        with open(RESULTS, encoding="utf-8") as f:
            if f.read() != text:
                print("FAIL: results.json differs from a fresh measurement")
                return 1
        print("results.json reproduces")
        return 0
    with open(RESULTS, "w", encoding="utf-8", newline="\n") as f:
        f.write(text)
    print(f"wrote {os.path.relpath(RESULTS, BENCHMARK)}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
