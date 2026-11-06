"""
Model the Student-t inverse polish of INVERSE_PREREGISTRATION.md (#34) and
record its predictions for the Excel wave, before any VBA is written. Writes
inverse_predictions.json.

The polish runs after PROB_TryStudentTInvTail's existing loop has converged,
when DegreesFreedom >= 1E5 and the target tail is at most 0.4:

    For Step = 1 To 3
        If X > 8 Then Exit For
        TailAtX = PROB_TryStudentTTail(X, DegreesFreedom)      ' the large-df route
        Density = PROB_TryStudentTPDF(X, DegreesFreedom)
        If Density <= 0 Then Exit For
        XNew = X + (TailAtX - Tail) / Density
        If XNew = X Then Exit For
        X = XNew
    Next

polish() is that loop in binary64, with production_route_model.route_tail as
TailAtX. outcomes() is the set of results the polish can return for a row:
  - every start within +-START_ULPS Doubles of the correctly rounded quantile;
  - the route's tail shifted by up to +-TAIL_ULPS Doubles, chosen
    independently at every step;
  - the density scaled by 1 + d, for d in DENSITY_SHIFTS, also chosen at every
    step.
The existing loop's stopping rule leaves X within about 6 Doubles of the
quantile (EXCEL_WAVE_RESULT.md). The tail shift covers the model's 1-ulp
disagreement with Excel (PT1 of the first wave), and the density shift covers
any last-bit difference in PROB_TryStudentTPDF.

    predictions  one record per committed observation in the main grid,
                 tchi_large_df and the holdout:
                   PI   StudentT_InverseCumulative rows the polish reaches:
                        the outcome set, and the bound
                        (2^-50 + 2^-52) * min(p, 1-p) / (|q| f(q)) + 2^-51;
                   PU   every other Student-t row, the density rows and every
                        F row: bit-identical to the committed observation.
    validation   the outcome set and the bound on a deterministic set of
                 (p, df), with the worst error as a fraction of its bound.

The holdout rows this file reads are the ones the first wave already observed.
The new holdout of INVERSE_PREREGISTRATION.md §6 is skipped: its rows are
recognised by redrawing their arguments from the seed, never by reading them,
and check_inverse_wave.py checks them (PH) after the export.

    python predict_inverse.py            write inverse_predictions.json
    python predict_inverse.py --check    recompute; fail if the file differs
"""
import argparse
import json
import math
import os
import sys
import time

import mpmath as mp

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import predict_production as PP  # noqa: E402
import production_route_model as P  # noqa: E402
import validate_expansion as V  # noqa: E402

OUT = os.path.join(HERE, "inverse_predictions.json")
GRIDS = PP.GRIDS + ("holdout/holdout_grid.csv",)
STEPS = 3                   # PROB_T_LDF_POLISH_STEPS
MAX_TAIL = 0.4              # PROB_T_LDF_POLISH_MAX_TAIL
START_ULPS = 32
TAIL_ULPS = 2
DENSITY_SHIFTS = (0.0, 1e-9, -1e-9)
SEAM_ULPS = 64              # rows whose quantile is this close to t = 8 are refused

V_P = (0.6, 0.65, 0.75, 0.9, 0.95, 0.975, 0.99, 0.999, 0.9999, 1 - 1e-6, 1 - 1e-9, 1 - 1e-12)
V_DF = (1e5, 2e5, 5e5, 1e6, 1e7, 1e8)


def tail_of(p):
    """The target tail PROB_TryStudentTInvTail receives (1 - p is exact for p > 0.5)."""
    return 1.0 - p if p > 0.5 else p


def reaches_polish(p, df):
    return p != 0.5 and P.DF_MIN <= df <= PP.T_MAX_DF and tail_of(p) <= MAX_TAIL


def density(x, df):
    with mp.workdps(40):
        x, df = mp.mpf(x), mp.mpf(df)
        return float(mp.e ** (mp.loggamma((df + 1) / 2) - mp.loggamma(df / 2)
                              - mp.log(df * mp.pi) / 2 - (df + 1) / 2 * mp.log1p(x * x / df)))


def polish(x, tail, df, tail_shift=0, density_shift=0.0):
    for _ in range(STEPS):
        if x > P.T_MAX:
            break
        t = P.route_tail(x, df)
        if tail_shift:
            t = t + tail_shift * math.ulp(t)
        d = density(x, df) * (1.0 + density_shift)
        if d <= 0.0:
            break
        x_new = x + (t - tail) / d
        if x_new == x:
            break
        x = x_new
    return x


def reachable(starts, tail, df):
    """Every result of the polish from these starts, with the tail shift and the
    density scale chosen independently at every step: the states are expanded
    step by step, so each Double is evaluated once per step."""
    frontier, done = set(starts), set()
    for _ in range(STEPS):
        following = set()
        for x in frontier:
            if x > P.T_MAX:
                done.add(x)
                continue
            t0, d0 = P.route_tail(x, df), density(x, df)
            for e in range(-TAIL_ULPS, TAIL_ULPS + 1):
                t = t0 + e * math.ulp(t0) if e else t0
                for shift in DENSITY_SHIFTS:
                    d = d0 * (1.0 + shift)
                    if d <= 0.0:
                        done.add(x)
                        continue
                    x_new = x + (t - tail) / d
                    (done if x_new == x else following).add(x_new)
        frontier = following
    return done | frontier


def exact_quantile(tail, df):
    with mp.workdps(60):
        seed = mp.sqrt(2) * mp.erfinv(1 - 2 * mp.mpf(tail))
        return mp.findroot(lambda x: V.oracle_tail(x, mp.mpf(df)) - mp.mpf(tail), seed)


def outcomes(p, df):
    """(correctly rounded |q|, sorted outcome set of |q|, worst relative error, bound)."""
    tail = tail_of(p)
    q = exact_quantile(tail, df)
    qd = float(q)
    if abs(qd - P.T_MAX) <= SEAM_ULPS * math.ulp(P.T_MAX):
        raise SystemExit(f"p={p!r}, df={df!r}: quantile within {SEAM_ULPS} ulps of t = 8")
    found = reachable({qd + k * math.ulp(qd) for k in range(-START_ULPS, START_ULPS + 1)}, tail, df)
    with mp.workdps(60):
        worst = max(float(abs((mp.mpf(x) - q) / q)) for x in found)
    return qd, sorted(found), worst, PP.quantile_bound(p, qd, df)


def validation():
    rows = []
    for df in V_DF:
        for p in V_P:
            qd, found, worst, bound = outcomes(p, df)
            if qd > P.T_MAX:
                continue
            rows.append({"p": p, "df": df, "q": qd, "outcomes": len(found),
                         "spread_ulps": round((found[-1] - found[0]) / math.ulp(qd)),
                         "worst": worst, "bound": bound, "fraction": worst / bound})
    worst = max(rows, key=lambda r: r["fraction"])
    return {"points": len(rows), "worst_fraction_of_bound": worst["fraction"],
            "worst_at": {"p": worst["p"], "df": worst["df"]},
            "within_bound": all(r["worst"] <= r["bound"] for r in rows), "rows": rows}


def section6_rows():
    """(function, arg1, arg2) of the untouched holdout, redrawn from its seed."""
    sys.path.insert(0, os.path.join(PP.BENCH, "holdout"))
    import generate_holdout  # noqa: E402
    return {(fn, a1, a2) for fn, a1, a2 in generate_holdout.student_t_inverse_draws()}


def predictions():
    records, unchanged = [], {}
    fresh = section6_rows()
    for grid in GRIDS:
        for row in PP.read(grid):
            fn = row.get("function", "")
            if not (fn.startswith("StudentT_") or fn.startswith("F_")):
                continue
            if grid.startswith("holdout/") and row.get("arg2", "").strip() and \
                    (fn, float(row["arg1"]), float(row["arg2"])) in fresh:
                continue
            key = {"grid": grid, "function": fn, "args": [row["arg1"], row["arg2"]],
                   "regime": row.get("regime", "")}
            if fn == "StudentT_InverseCumulative" and row.get("arg2", "").strip():
                p, df = float(row["arg1"]), float(row["arg2"])
                if reaches_polish(p, df):
                    qd, found, worst, bound = outcomes(p, df)
                    if qd <= P.T_MAX:
                        sign = 1.0 if p > 0.5 else -1.0
                        records.append({**key, "prediction": "PI",
                                        "outcomes": [repr(sign * x) for x in found],
                                        "worst_error": worst, "bound": bound})
                        continue
            unchanged[grid] = unchanged.get(grid, 0) + 1
    return records, unchanged


def build():
    start = time.time()
    records, unchanged = predictions()
    return {
        "issue": "#34",
        "design": "INVERSE_PREREGISTRATION.md",
        "polish": {"steps": STEPS, "max_tail": MAX_TAIL, "df_min": P.DF_MIN, "t_max": P.T_MAX,
                   "start_ulps": START_ULPS, "tail_ulps": TAIL_ULPS,
                   "density_shifts": list(DENSITY_SHIFTS)},
        "validation": validation(),
        "predictions": {"grids": list(GRIDS),
                        "PI": len(records),
                        "PI_within_bound": all(r["worst_error"] <= r["bound"] for r in records),
                        "PU_unchanged_rows": unchanged,
                        "rows": records},
        "new_holdout_read": False,
        "runtime_s": round(time.time() - start, 1),
    }


def comparable(result):
    out = dict(result)
    out.pop("runtime_s", None)
    return out


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    ap.add_argument("--check", action="store_true")
    a = ap.parse_args(argv)
    result = build()
    if a.check:
        with open(OUT, encoding="utf-8") as f:
            recorded = json.load(f)
        if comparable(recorded) != comparable(json.loads(json.dumps(result))):
            print("FAIL: inverse_predictions.json differs from a recomputation")
            return 1
        print("PASS: inverse_predictions.json reproduces")
        return 0
    with open(OUT, "w", encoding="utf-8", newline="\n") as f:
        json.dump(result, f, indent=1)
        f.write("\n")
    v, p = result["validation"], result["predictions"]
    print(f"validation: {v['points']} points, worst {v['worst_fraction_of_bound']:.3f} of the bound "
          f"at p={v['worst_at']['p']}, df={v['worst_at']['df']:.0e}; within bound: {v['within_bound']}")
    print(f"predictions: PI {p['PI']} (within bound: {p['PI_within_bound']}), "
          f"PU unchanged {p['PU_unchanged_rows']}")
    return 0 if v["within_bound"] and p["PI_within_bound"] else 1


if __name__ == "__main__":
    sys.exit(main())
