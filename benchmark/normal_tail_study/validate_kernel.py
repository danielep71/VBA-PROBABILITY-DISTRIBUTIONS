"""
Apply PREREGISTRATION.md (#72) to the fitted kernel and write results.json.

The steps follow the preregistration:
  - §3 oracles: 60- and 100-dps agreement at every point, and quadrature at
    every 25th grid point;
  - A1 to A5 and A8 on the validation set V and the approximant set A;
  - the seams;
  - A7, the #34 route;
  - §7 predictions for every committed affected row outside the independent
    holdout;
  - last, A6 on the holdout H.

H is evaluated once (§5). results.json records that exposure in its
`holdout_exposure` block, and holdout_record.py (#80) enforces it: every mode
reads and checks the existing record before it computes or writes anything,
and refuses a missing, malformed or inconsistent one.

    python validate_kernel.py --verify
        Read-only replay. Recomputes everything except H and compares it
        with results.json. Never evaluates H, never writes.
    python validate_kernel.py --regenerate --reason TEXT
        Recomputes everything except H and rewrites results.json. The
        recorded holdout_exposure and acceptance.A6 are carried over
        unchanged; the reason is appended to `regenerations`.
    python validate_kernel.py --first-evaluation
        Evaluates H once. Refused unless holdout_exposure.status is
        NOT_EVALUATED; for this study it is EVALUATED (#75).
    python validate_kernel.py --init
        New study only: writes V with status NOT_EVALUATED when results.json
        does not exist.

Exit status: 0 success, 1 study failure or mismatch, 2 usage, 3 refused (the
output is left unchanged).
"""
import argparse
import csv
import hashlib
import json
import math
import os
import platform
import random
import struct
import sys
from decimal import Decimal

import mpmath as mp

import holdout_record as hr
import kernel_model as km

HERE = os.path.dirname(os.path.abspath(__file__))
BENCH = os.path.dirname(HERE)
OUT = os.path.join(HERE, "results.json")
PREREGISTRATION = os.path.join(HERE, "PREREGISTRATION.md")
sys.path.insert(0, os.path.join(BENCH, "student_t_large_df_study"))
import binary64_route_model as rm  # noqa: E402

B = km.BOUND
DBL_MIN = 2.0 ** -1022
BANDS = ((0.0, 3.0), (3.0, 7.07106781186547), (7.07106781186547, 37.5))
GRIDS = (
    "probability_accuracy_grid.csv",
    "survival_boundary/survival_boundary_grid.csv",
    "holdout_older/holdout_older_grid.csv",
)
DIRECT = ("NormalStandard_Survival", "NormalStandard_Cumulative", "NormalStandard_Density")
STANDARDISED = ("Normal_Cumulative", "Normal_Survival", "Lognormal_Cumulative", "Lognormal_Survival")
INTERVALS = ("NormalStandard_IntervalProbability", "Normal_IntervalProbability")
INVERSES = tuple(f"{fam}_{s}" for fam in ("NormalStandard", "Normal", "Lognormal")
                 for s in ("InverseCumulative", "InverseSurvival"))
UNCHANGED = ("Normal_Density", "Lognormal_Density", "NormalStandard_InverseCumulativeFast")


# ---------------------------------------------------------------------------
# Oracles (§3)
# ---------------------------------------------------------------------------

def q_ref(z, dps=60):
    with mp.workdps(dps):
        return mp.erfc(mp.mpf(z) / mp.sqrt(2)) / 2


def cdf_ref(z, dps=60):
    with mp.workdps(dps):
        return mp.erfc(-mp.mpf(z) / mp.sqrt(2)) / 2


def pdf_ref(z, dps=60):
    with mp.workdps(dps):
        z = mp.mpf(z)
        return mp.exp(-z * z / 2) / mp.sqrt(2 * mp.pi)


def q_quadrature(z):
    """Q(z) by tanh-sinh quadrature, independent of erfc.

    With t = z + s, Q(z) = phi(z) * integral_0^inf exp(-z*s - s^2/2) ds. The
    integrand decays on the scale 1/(z + 1), so the breakpoints follow it;
    integrating exp(-t^2/2) directly over [z, inf) loses all precision in the
    far tail.
    """
    with mp.workdps(100):
        z = mp.mpf(z)
        h = 1 / (z + 1)
        f = lambda s: mp.exp(-z * s - s * s / 2)  # noqa: E731
        integral = mp.quad(f, [0, h, 4 * h, 16 * h, 64 * h, mp.inf])
        return mp.exp(-z * z / 2) / mp.sqrt(2 * mp.pi) * integral


def digits(a, b):
    """Significant digits of agreement between two high-precision values."""
    with mp.workdps(120):
        if a == b:
            return 100.0
        return float(-mp.log10(abs((mp.mpf(a) - mp.mpf(b)) / mp.mpf(b))))


def rel(value, reference):
    with mp.workdps(60):
        reference = mp.mpf(str(reference))
        if reference == 0:
            return 0.0 if value == 0 else math.inf
        return float(abs((mp.mpf(value) - reference) / reference))


def hexd(x):
    return struct.pack(">d", x).hex()


# ---------------------------------------------------------------------------
# Sets (§5)
# ---------------------------------------------------------------------------

def validation_set():
    grid = [k / 100 for k in range(0, 3751)]
    rnd = random.Random(72001)
    return grid, [rnd.uniform(0, 37.5) for _ in range(20000)]


# Holdout H (§5). Its points are bound by SHA-256 in holdout_exposure.
HOLDOUT_DEFINITION = [
    {"seed": 72002, "count": 10000, "low": 0, "high": 37.5},
    {"seed": 72003, "count": 5000, "low": 0, "high": 8},
]


def holdout_set():
    return hr.holdout_points(HOLDOUT_DEFINITION)


def approximant_set():
    rnd = random.Random(72004)
    return [rnd.uniform(0, 3) for _ in range(20000)] + [k / 1000 for k in range(0, 3001)]


def around(x):
    return [math.nextafter(x, -math.inf), x, math.nextafter(x, math.inf)]


def last_normal_q():
    """Largest Double z with Q(z) >= 2^-1022 (characterisation only)."""
    lo, hi = 37.0, 38.5
    for _ in range(200):
        mid = (lo + hi) / 2
        if q_ref(mid) >= DBL_MIN:
            lo = mid
        else:
            hi = mid
        if math.nextafter(lo, math.inf) >= hi:
            break
    return lo


def seam_set():
    seams = [("split", z) for z in around(3.0)]
    for k in range(1, 9856, 97):
        seams += [("zh_step", z) for z in around(k / 256)]
    seams += [("last_normal_q", z) for z in around(last_normal_q())]
    seams += [("tail_underflow", z) for z in around(38.5)]
    return seams


# ---------------------------------------------------------------------------
# Kernel measurement
# ---------------------------------------------------------------------------

def measure_point(z, kernel="candidate", stability=True):
    q60 = q_ref(z)
    out = {"z": z}
    if stability:
        out["stability_digits"] = min(digits(q60, q_ref(z, 100)), digits(pdf_ref(z), pdf_ref(z, 100)))
    upper, pdf = km.KERNELS[kernel]
    if q60 >= DBL_MIN:
        out["q"] = rel(upper(z), q60)
        out["cdf"] = rel(km.cdf(z, kernel), cdf_ref(z))
    p60 = pdf_ref(z)
    if p60 >= DBL_MIN:
        out["pdf"] = rel(pdf(z), p60)
    return out


def worst_of(rows, key):
    candidates = [r for r in rows if key in r]
    if not candidates:
        return {"value": None, "z": None}
    w = max(candidates, key=lambda r: r[key])
    return {"value": w[key], "z": w["z"]}


def band_worst(rows, key="q"):
    out = []
    for lo, hi in BANDS:
        inside = [r for r in rows if lo <= r["z"] < hi and key in r]
        w = max(inside, key=lambda r: r[key])
        out.append({"band": [lo, hi], "worst": w[key], "z": w["z"], "points": len(inside)})
    return out


def kernel_pass(rows):
    return all(r.get(k, 0.0) <= B for r in rows for k in ("q", "cdf", "pdf"))


# ---------------------------------------------------------------------------
# A1, A4: exact-arithmetic checks
# ---------------------------------------------------------------------------

def approximant_check(raw):
    with mp.workdps(60):
        p = [mp.mpf(r["value"]) for r in raw["numerator"]]
        d = [mp.mpf(r["value"]) for r in raw["denominator"]]
        worst = (0.0, None)
        for z in approximant_set():
            zm = mp.mpf(z)
            approx = mp.polyval(p[::-1], zm) / mp.polyval(d[::-1], zm)
            exact = mp.erfc(zm / mp.sqrt(2)) / 2 * mp.exp(zm * zm / 2)
            e = float(abs(approx / exact - 1))
            if e > worst[0]:
                worst = (e, z)
    return {"pass": worst[0] <= 2.0 ** -55, "worst": worst[0], "z": worst[1], "limit": 2.0 ** -55,
            "points": 23001}


def cancellation_check(k):
    """A2' (Amendment 1): Horner's cancellation factor for P and D on set A."""
    limit = 1 + 2.0 ** -20
    out = {"limit": limit}
    with mp.workdps(40):
        for name in ("numerator", "denominator"):
            c = [mp.mpf(x) for x in k[name]]
            worst = (mp.mpf(1), None)
            for z in approximant_set():
                zm = mp.mpf(z)
                powers = [zm ** j for j in range(len(c))]
                factor = sum(abs(a) * p for a, p in zip(c, powers)) / abs(sum(a * p for a, p in zip(c, powers)))
                if factor > worst[0]:
                    worst = (factor, z)
            out[name] = {"worst_factor_minus_one": float(worst[0] - 1), "z": worst[1]}
    out["pass"] = all(out[n]["worst_factor_minus_one"] <= 2.0 ** -20 for n in ("numerator", "denominator"))
    return out


def truncation_check(levels):
    rows = []
    with mp.workdps(100):
        for z in [3 + i / 2 for i in range(0, 15)]:
            zm = mp.mpf(z)
            d = zm
            for k in range(levels, 0, -1):
                d = zm + k / d
            mills = mp.erfc(zm / mp.sqrt(2)) / 2 / (mp.exp(-zm * zm / 2) / mp.sqrt(2 * mp.pi))
            rows.append({"z": z, "truncation": float(abs((1 / d) / mills - 1))})
    monotone = all(b["truncation"] <= a["truncation"] for a, b in zip(rows, rows[1:]))
    return {"pass": rows[0]["truncation"] <= 2.0 ** -56 and monotone, "limit_at_3": 2.0 ** -56,
            "non_increasing": monotone, "points": rows}


# ---------------------------------------------------------------------------
# A7: the #34 route
# ---------------------------------------------------------------------------

def route_candidate(t, df):
    series = 0.0
    for k in range(rm.ORDER, 0, -1):
        g = 0.0
        for c in rm.COEFFICIENTS[k - 1]:
            g = g * t + c
        series = series / df + g
    series = series / df
    return km.candidate_upper_tail(t) + km.candidate_pdf(t) * series


def route_check():
    with open(os.path.join(BENCH, "student_t_large_df_study", "results.json"), encoding="utf-8") as f:
        study = json.load(f)
    fitting = [(p["t"], p["df"]) for p in study["forward_fitting"]
               if p["df"] >= rm.DF_MIN and p["t"] <= rm.T_MAX]
    holdout = [(p["t"], p["df"]) for p in study["forward_holdout"] if p["inside_region"]]
    out = {}
    for name, points in (("fitting", fitting), ("holdout", holdout)):
        rows = [{"t": t, "df": df, "error": rel(route_candidate(t, df), rm.oracle_tail(t, df))}
                for t, df in points]
        out[name] = {"points": len(rows), "worst": max(r["error"] for r in rows),
                     "rows": rows}
    out["pass"] = all(out[n]["worst"] <= B for n in ("fitting", "holdout"))
    return out


# ---------------------------------------------------------------------------
# §7 predictions
# ---------------------------------------------------------------------------

def parse_observed(raw):
    t = (raw or "").strip()
    if t == "" or t.upper() == "ERROR":
        return None
    total = sum((Decimal(p.strip()) for p in t.split(";")), Decimal(0))
    return float(total)


def ulps_apart(a, b):
    if a == b:
        return 0
    ia = struct.unpack(">q", struct.pack(">d", a))[0]
    ib = struct.unpack(">q", struct.pack(">d", b))[0]
    if (ia < 0) != (ib < 0):
        return abs(ia) + abs(ib)
    return abs(ia - ib)


def contract_thresholds():
    out = {}
    with open(os.path.join(BENCH, "accuracy_contracts.csv"), encoding="utf-8-sig") as f:
        for r in csv.DictReader(f):
            if r["status"] == "active" and r["measure"] == "output_error" and r["metric"] == "relative":
                out[(r["function"], r["regime"])] = float(r["threshold"])
    return out


def args_of(row):
    def cell(k):
        t = (row.get(k) or "").strip()
        return float(t) if t else None
    args = [cell(f"arg{i}") for i in range(1, 5)]
    fn = row["function"]
    if fn.startswith("Normal_"):
        if fn == "Normal_IntervalProbability":
            # The density-helper exporter calls (A1, A2, 0#, A3): arg3 is the
            # standard deviation and the mean is zero
            # (benchmark/density_helpers/M_STATS_PROBDIST_DENSHELP.bas).
            args = [args[0], args[1], 0.0, 1.0 if args[2] is None else args[2]]
        else:
            args[1] = 0.0 if args[1] is None else args[1]
            args[2] = 1.0 if args[2] is None else args[2]
    return args


def exact_z(fn, args):
    with mp.workdps(60):
        if fn.startswith("Lognormal"):
            return (mp.log(mp.mpf(args[0])) - mp.mpf(args[1])) / mp.mpf(args[2])
        if fn.startswith("Normal_"):
            return (mp.mpf(args[0]) - mp.mpf(args[1])) / mp.mpf(args[2])
        return mp.mpf(args[0])


# P2 conditioning, as restated by Amendment 2 (#78). F is the function the row
# evaluates at its exact standardised argument z: the survival S(z) = Q(z) or
# the cumulative Phi(z), for either sign of z.
#
#   kappa = |z| phi(z) / F    the relative condition number |z F'(z) / F(z)|;
#                             nonnegative for both signs (the frozen survival
#                             expression z phi / Q is negative for z < 0)
#   eta   >= |z~ - z|         first-order bound on the absolute error of the
#                             computed z~
#   bound = B + phi(z) eta / F
#
# Wherever z != 0, eta = delta |z| with the frozen delta of §7 P2, so
# phi eta / F = kappa delta exactly. The absolute form also holds at z = 0,
# where delta is 0/0 for Lognormal.

U52 = 2 ** -52                          # two roundings of 2^-53 each, or one ulp of VBA Log


def density(z):
    with mp.workdps(60):
        return mp.exp(-z * z / 2) / mp.sqrt(2 * mp.pi)


def condition(z, value):
    """kappa = |z| phi(z) / F, the relative condition number of F at z (Amendment 2)."""
    with mp.workdps(60):
        return float(abs(z) * density(z) / mp.mpf(str(value)))


def z_error(fn, args, z):
    """eta: first-order bound on |z~ - z| for the computed standardised argument (Amendment 2).

    NormalStandard: z is the argument itself, eta = 0.
    Normal: z~ = fl(fl(x + (-mu)) / sigma), two roundings, eta = 2^-52 |z|.
    Lognormal: z~ = fl(fl(L~ + (-mu)) / sigma) with |L~ - ln x| <= 2^-52 |ln x|
    (one ulp of VBA Log), eta = 2^-52 (|z| + |ln x| / sigma).
    """
    with mp.workdps(60):
        if fn.startswith("Lognormal"):
            return U52 * (abs(z) + abs(mp.log(mp.mpf(args[0]))) / mp.mpf(args[2]))
        if fn.startswith("Normal_"):
            return U52 * abs(z)
        return mp.mpf(0)


def standardisation_term(z, value, eta):
    """phi(z) eta / F: first-order relative error of F caused by an error eta in z."""
    with mp.workdps(60):
        return float(density(z) * eta / mp.mpf(str(value)))


def interval_bound(fn, args, reference):
    with mp.workdps(60):
        if fn == "NormalStandard_IntervalProbability":
            lo, hi = args[0], args[1]
            zl, zu = mp.mpf(lo), mp.mpf(hi)
        else:
            lo = km.standardize(args[0], args[2], args[3])
            hi = km.standardize(args[1], args[2], args[3])
            zl = (mp.mpf(args[0]) - mp.mpf(args[2])) / mp.mpf(args[3])
            zu = (mp.mpf(args[1]) - mp.mpf(args[2])) / mp.mpf(args[3])
        _, branch, _ = km.interval(lo, hi)
        if branch == "upper":
            tails = [(zl, q_ref(zl)), (zu, q_ref(zu))]
        elif branch == "lower":
            tails = [(zu, q_ref(-zu)), (zl, q_ref(-zl))]
        else:
            tails = [(zl, q_ref(-zl)), (zu, q_ref(zu))]
        p = mp.mpf(str(reference))
        total = 0.0
        for z, tail in tails:
            b = B + standardisation_term(z, tail, z_error(fn, args, z))
            total += float(tail * b / p)
        if branch == "crossing":
            total += float(2.0 ** -53 / p)
    return total, branch


def predictions():
    thresholds = contract_thresholds()
    rows, unchanged = [], 0
    for grid in GRIDS:
        with open(os.path.join(BENCH, grid), encoding="utf-8-sig") as f:
            for row in csv.DictReader(f):
                fn = row["function"]
                if fn in UNCHANGED:
                    unchanged += 1
                    continue
                if fn not in DIRECT + STANDARDISED + INTERVALS + INVERSES:
                    continue
                args = args_of(row)
                reference = Decimal(row["reference"].strip())
                observed = parse_observed(row["observed_vba"])
                now, _ = km.public(fn, args, "current")
                new, _ = km.public(fn, args, "candidate")
                rec = {
                    "grid": grid, "function": fn, "regime": row["regime"],
                    "args": [a for a in args if a is not None],
                    "predicted": repr(new), "predicted_hex": hexd(new),
                    "predicted_error": rel(new, reference),
                    "current_model_error": rel(now, reference),
                    "model_vs_committed_ulps": None if observed is None else ulps_apart(now, observed),
                }
                if fn in DIRECT:
                    rec["prediction"] = "P1"
                    rec["bound"] = B
                elif fn in STANDARDISED:
                    rec["prediction"] = "P2"
                    z = exact_z(fn, args)
                    eta = z_error(fn, args, z)
                    rec["z"] = float(z)
                    rec["condition"] = condition(z, reference)
                    rec["z_error"] = float(eta)
                    rec["bound"] = B + standardisation_term(z, reference, eta)
                elif fn in INTERVALS:
                    rec["prediction"] = "P3"
                    rec["bound"], rec["branch"] = interval_bound(fn, args, reference)
                else:
                    rec["prediction"] = "P4"
                    limit = thresholds.get((fn, row["regime"]))
                    rec["contract"] = limit
                    rec["predicted_breach"] = limit is not None and rec["predicted_error"] > limit
                    if row["regime"] == "split_boundary":
                        rec["bound"] = 1e-14
                if "bound" in rec:
                    rec["within_bound"] = rec["predicted_error"] <= rec["bound"]
                rows.append(rec)
    summary = {}
    for tag in ("P1", "P2", "P3", "P4"):
        mine = [r for r in rows if r["prediction"] == tag]
        bounded = [r for r in mine if "bound" in r]
        summary[tag] = {
            "rows": len(mine),
            "worst_predicted_error": max(r["predicted_error"] for r in mine),
            "outside_bound": [r["function"] + " " + str(r["args"]) for r in bounded if not r["within_bound"]],
            "bounds_finite_nonnegative": all(math.isfinite(r["bound"]) and r["bound"] >= 0 for r in bounded),
            "max_model_vs_committed_ulps": max((r["model_vs_committed_ulps"] or 0) for r in mine),
        }
    summary["P4"]["predicted_breaches"] = [r["function"] + " " + str(r["args"])
                                           for r in rows if r.get("predicted_breach")]
    p2 = [r for r in rows if r["prediction"] == "P2"]
    summary["P2"]["negative_z_rows"] = [r["function"] + " " + str(r["args"]) for r in p2 if r["z"] < 0]
    summary["P2"]["centre_rows"] = [r["function"] + " " + str(r["args"]) for r in p2 if r["z"] == 0]
    summary["unchanged_rows"] = unchanged
    return {"grids": list(GRIDS), "summary": summary, "rows": rows}


# ---------------------------------------------------------------------------

def compute():
    """Everything except H: the oracles, A1-A5, A7, A8, the seams and §7."""
    with open(km.COEFFICIENTS_PATH, "rb") as f:
        coefficients_sha = hashlib.sha256(f.read()).hexdigest()
    raw = km.COEFFICIENTS["raw"]
    problems = []

    a1 = approximant_check(raw)
    a2_frozen = {"pass": all(mp.mpf(r["value"]) > 0 for r in raw["numerator"] + raw["denominator"]),
                 "negative": [f"{name}[{i}] = {r['value']}" for name in ("numerator", "denominator")
                              for i, r in enumerate(raw[name]) if mp.mpf(r["value"]) <= 0]}
    a2 = cancellation_check(km.COEFFICIENTS)
    literals = raw["numerator"] + raw["denominator"] + [raw["inv_sqrt_two_pi"]]
    a3 = {"pass": all(r["pair_is_exact"] and float(r["hi"]) + float(r["lo"]) == float(r["double"])
                      for r in literals), "constants": len(literals)}
    a4 = truncation_check(km.COEFFICIENTS["levels"])

    grid, uniform = validation_set()
    v_rows = [measure_point(z) for z in grid + uniform]
    seams = []
    for name, z in seam_set():
        m = measure_point(z)
        m["seam"] = name
        seams.append(m)
    current_rows = [measure_point(z, "current", stability=False) for z in grid + uniform]

    stability = min(r["stability_digits"] for r in v_rows + seams)
    quadrature = []
    for z in grid[::25]:
        if q_ref(z) >= DBL_MIN:
            quadrature.append({"z": z, "agreement_digits": digits(q_quadrature(z), q_ref(z, 100))})
    independence = min(r["agreement_digits"] for r in quadrature)

    a5 = {"pass": kernel_pass(v_rows) and kernel_pass(seams),
          "points": len(v_rows), "seam_probes": len(seams),
          "worst": {k: worst_of(v_rows, k) for k in ("q", "cdf", "pdf")},
          "worst_seam": {k: worst_of(seams, k) for k in ("q", "cdf", "pdf")}}
    meets_now = [i for i, r in enumerate(current_rows) if r.get("q", math.inf) <= B]
    a8 = {"pass": all(v_rows[i].get("q", math.inf) <= B for i in meets_now),
          "current_meets_bound": len(meets_now), "points": len(current_rows)}
    a7 = route_check()

    if stability < 50:
        problems.append(f"oracle stability {stability:.1f} < 50 digits")
    if independence < 40:
        problems.append(f"quadrature agreement {independence:.1f} < 40 digits")
    for name, check in (("A1", a1), ("A2", a2), ("A3", a3), ("A4", a4), ("A5", a5), ("A7", a7), ("A8", a8)):
        if not check["pass"]:
            problems.append(f"{name} failed")

    print(f"computed without the holdout: kernel worst Q {a5['worst']['q']['value']:.3e}, "
          f"pdf {a5['worst']['pdf']['value']:.3e}; route worst "
          f"{max(a7['fitting']['worst'], a7['holdout']['worst']):.3e}; {len(problems)} problem(s)")
    return {
        "study": "normal upper-tail kernel (#72)",
        "preregistration": "PREREGISTRATION.md",
        "coefficients_sha256": coefficients_sha,
        "problems": problems,
        "versions": {"python": platform.python_version(), "mpmath": mp.__version__},
        "bound": B,
        "oracle": {"stability_min_digits": stability, "quadrature_min_digits": independence,
                   "quadrature_points": len(quadrature)},
        "amendments": ["Amendment 1: A2 (strict positivity) replaced by A2' (cancellation factor "
                       "<= 1 + 2^-20 on set A), recorded after the derivation run failed A2 only",
                       "Amendment 2 (#78): P2 conditioning restated as kappa = |z| phi(z) / F for survival "
                       "and cumulative alike, and kappa * delta as phi(z) * eta / F with eta = 2^-52 |z| "
                       "(Normal) or 2^-52 (|z| + |ln x| / sigma) (Lognormal), defined at z = 0; recorded "
                       "after A6 was observed, nothing refitted, no threshold changed"],
        "acceptance": {"A1": a1, "A2": a2, "A2_as_frozen": a2_frozen, "A3": a3, "A4": a4, "A5": a5,
                       "A7": a7, "A8": a8},
        "bands": {"current": band_worst(current_rows), "candidate": band_worst(v_rows)},
        "seams": [s for s in seams if s["seam"] in ("split", "last_normal_q", "tail_underflow")],
        "predictions": predictions(),
        "holdout_559_inspected": False,
    }


def evaluate_holdout(points):
    """A6. Called only by holdout_record.first_evaluation, once per study."""
    h_rows = [measure_point(z) for z in points]
    return {"pass": kernel_pass(h_rows), "points": len(h_rows), "evaluated_once": True,
            "worst": {k: worst_of(h_rows, k) for k in ("q", "cdf", "pdf")},
            "stability_min_digits": min(r["stability_digits"] for r in h_rows)}


def make_study(output=OUT, log=None):
    return hr.Study(output=output, preregistration=PREREGISTRATION, coefficients=km.COEFFICIENTS_PATH,
                    holdout_definition=HOLDOUT_DEFINITION, compute=compute,
                    evaluate_holdout=evaluate_holdout, log=log or print)


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    mode = parser.add_mutually_exclusive_group(required=True)
    mode.add_argument("--verify", action="store_const", dest="mode", const="verify",
                      help="read-only replay of everything except the holdout")
    mode.add_argument("--regenerate", action="store_const", dest="mode", const="regenerate",
                      help="rewrite everything except the holdout; keep the recorded exposure and A6")
    mode.add_argument("--first-evaluation", action="store_const", dest="mode", const="first_evaluation",
                      help="evaluate the holdout once; refused once an exposure is recorded")
    mode.add_argument("--init", action="store_const", dest="mode", const="init",
                      help="new study only: write V with status NOT_EVALUATED")
    parser.add_argument("--reason", help="why --regenerate is run; recorded in results.json")
    parser.add_argument("--output", default=OUT, help="the results file (default: results.json here)")
    args = parser.parse_args(argv)
    if args.reason is not None and args.mode != "regenerate":
        parser.error("--reason applies to --regenerate only")
    return hr.run(make_study(os.path.abspath(args.output)), args.mode, args.reason)


if __name__ == "__main__":
    sys.exit(main())
