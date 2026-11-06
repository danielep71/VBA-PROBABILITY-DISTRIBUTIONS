"""
Measure the production Student-t large-df route and record its predictions for
the Excel wave (#34), as PRODUCTION_PREREGISTRATION.md requires, before any VBA
is written. Writes production_predictions.json.

  measurement   production_route_model.route_tail against the study's primary
                oracle, I_z(df/2, 1/2) / 2 at 60 digits, on the validation set
                V of PRODUCTION_PREREGISTRATION.md §4, with the worst point
                re-checked at 100 digits.
  seams         at df = DF_MIN and |t| = T_MAX, the route and the faithful
                binary64 continued fraction of the current code on both sides.
  predictions   one record per committed Student-t observation the route can
                reach, in the main grid and in tchi_large_df:
                  PT1  StudentT_Survival / StudentT_Cumulative in the region:
                       the predicted Double, and the bound 2^-50;
                  PT2  StudentT_InverseCumulative whose quantile is in the
                       region: the bound (2^-50 + 2^-52) * min(p, 1-p) /
                       (|q| f(q)) + 2^-51 on the quantile's relative error;
                  PT3  every other Student-t row, the density rows and every F
                       row: bit-identical to the baseline export.

The independent holdout is not read.

    python predict_production.py            write production_predictions.json
    python predict_production.py --check    recompute; fail if the file differs
"""
import argparse
import csv
import hashlib
import json
import math
import os
import sys
import time

import mpmath as mp

HERE = os.path.dirname(os.path.abspath(__file__))
BENCH = os.path.dirname(HERE)
sys.path.insert(0, HERE)
import production_route_model as P  # noqa: E402
import validate_expansion as V  # noqa: E402

OUT = os.path.join(HERE, "production_predictions.json")
GRIDS = ("probability_accuracy_grid.csv", "tchi_large_df/tchi_large_df_grid.csv")
FORWARD = ("StudentT_Survival", "StudentT_Cumulative")
INVERSE = ("StudentT_InverseCumulative",)
B = P.BOUND
U52 = 2.0 ** -52
ULPS = 4

# §4 validation set V: deterministic, disjoint from nothing it is compared with.
V_DF = (1e5, 2e5, 5e5, 1e6, 2e6, 5e6, 1e7, 2e7, 5e7, 1e8)
V_T = tuple([1e-4, 1e-3, 1e-2] + [round(0.05 * i, 2) for i in range(1, 161)]
            + [math.nextafter(3.0, 0.0), math.nextafter(3.0, 9.0),
               math.nextafter(7.07106781186547, 0.0), 7.07106781186547])


def oracle(t, df, dps=60):
    with mp.workdps(dps):
        return V.oracle_tail(mp.mpf(t), mp.mpf(df))


def rel(value, reference):
    with mp.workdps(60):
        return float(abs((mp.mpf(value) - reference) / reference))


def cf(t, df):
    """The current route's binary64 continued fraction (validate_expansion), or None."""
    saved = mp.mp.dps
    try:
        value = V.cf_double_tail(t, df)[0]
    finally:
        mp.mp.dps = saved
    return value


def rel_or_none(value, reference):
    return None if value is None else rel(value, reference)


def central_mass(abs_t, df):
    """PROB_TF_TryStudentTCentralMass takes the point (curvature <= 1E-8)."""
    return 2.0 * math.log(abs_t) + math.log1p(df) - math.log(df) <= math.log(1e-8)


T_MAX_DF = 1e8              # PROB_T_MAX_DF: the public functions reject larger df


def reaches_route(abs_t, df):
    """The public function returns a value from the new route."""
    return (abs_t > 0.0 and df <= T_MAX_DF and not central_mass(abs_t, df)
            and P.in_region(abs_t, df))


def measurement():
    rows = []
    for df in V_DF:
        for t in V_T:
            if not reaches_route(t, df):
                continue
            rows.append((rel(P.route_tail(t, df), oracle(t, df)), t, df))
    rows.sort(reverse=True)
    worst, t, df = rows[0]
    check = rel(P.route_tail(t, df), oracle(t, df, 100))
    return {"points": len(rows), "worst": worst, "worst_at": {"t": t, "df": df},
            "worst_at_100_digits": check, "within_bound": worst <= B,
            "top": [{"t": t, "df": df, "error": e} for e, t, df in rows[:10]]}


def improvement():
    """§8 of PREREGISTRATION.md, on the production arithmetic: at every study
    fitting point in the region where the current continued fraction misses the
    bound, the route must be more accurate than it."""
    with open(os.path.join(HERE, "results.json"), encoding="utf-8") as f:
        fitting = json.load(f)["forward_fitting"]
    rows, worse = [], []
    for point in fitting:
        t, df, current = point["t"], point["df"], float(point["C"])
        if not reaches_route(t, df):
            continue
        route = rel(P.route_tail(t, df), oracle(t, df))
        rows.append({"t": t, "df": df, "current_cf": current, "route": route})
        if current > B and not route < current:
            worse.append(f"t={t}, df={df}")
    return {"points": len(rows), "route_worse_than_current": worse,
            "route_worst": max(r["route"] for r in rows), "rows": rows}


def seams():
    out = []
    for t in (1.0, 2.0, 4.0, 8.0):
        below = math.nextafter(P.DF_MIN, 0.0)
        out.append({"seam": "df", "t": t, "df_below": below,
                    "cf_below": rel_or_none(cf(t, below), oracle(t, below)),
                    "route_at": rel(P.route_tail(t, P.DF_MIN), oracle(t, P.DF_MIN))})
    for df in (1e5, 1e6, 1e8):
        above = math.nextafter(P.T_MAX, 9.0)
        out.append({"seam": "t", "df": df, "t_above": above,
                    "route_at": rel(P.route_tail(P.T_MAX, df), oracle(P.T_MAX, df)),
                    "cf_above": rel_or_none(cf(above, df), oracle(above, df))})
    return out


def read(grid):
    with open(os.path.join(BENCH, grid), encoding="utf-8-sig") as f:
        return list(csv.DictReader(f))


def observed(raw):
    raw = (raw or "").strip()
    if not raw or "ERR" in raw.upper():
        return None
    return float(sum(mp.mpf(p) for p in raw.split(";")))


def quantile_bound(p, q, df):
    """(B + 2^-52) * min(p, 1-p) / (|q| f(q)) + 2^-51."""
    with mp.workdps(60):
        q, df, p = mp.mpf(q), mp.mpf(df), mp.mpf(p)
        log_f = (mp.loggamma((df + 1) / 2) - mp.loggamma(df / 2) - mp.log(df * mp.pi) / 2
                 - (df + 1) / 2 * mp.log1p(q * q / df))
        return float((B + U52) * min(p, 1 - p) / (abs(q) * mp.e ** log_f) + 2 * U52)


def predictions():
    records, unchanged = [], {}
    for grid in GRIDS:
        for row in read(grid):
            fn = row.get("function", "")
            if not (fn.startswith("StudentT_") or fn.startswith("F_")):
                continue
            a1, a2 = float(row["arg1"]), float(row["arg2"]) if row.get("arg2", "").strip() else None
            key = {"grid": grid, "function": fn, "args": [row["arg1"], row["arg2"]],
                   "regime": row.get("regime", "")}
            if fn in FORWARD and a2 is not None and reaches_route(abs(a1), a2):
                value = P.survival(a1, a2) if fn == "StudentT_Survival" else P.cumulative(a1, a2)
                with mp.workdps(60):
                    reference = mp.mpf(row["reference"])
                records.append({**key, "prediction": "PT1", "predicted": repr(value),
                                "predicted_error": rel(value, reference), "bound": B,
                                "ulps_allowed": ULPS})
                continue
            if fn in INVERSE and a2 is not None:
                q = float(row["reference"])
                if q != 0.0 and reaches_route(abs(q), a2):
                    records.append({**key, "prediction": "PT2",
                                    "bound": quantile_bound(a1, q, a2)})
                    continue
            unchanged[grid] = unchanged.get(grid, 0) + 1
    return records, unchanged


def build():
    start = time.time()
    with open(os.path.join(HERE, "coefficients.json"), "rb") as f:
        coefficients_sha = hashlib.sha256(f.read()).hexdigest()
    records, unchanged = predictions()
    result = {
        "issue": "#34",
        "design": "PRODUCTION_PREREGISTRATION.md",
        "route": {"order": P.ORDER, "df_min": P.DF_MIN, "t_max": P.T_MAX,
                  "bound": B, "coefficients_sha256": coefficients_sha,
                  "literals": [[[hi, lo] for _, _, hi, lo in row] for row in P.COEFFICIENTS]},
        "measurement": measurement(),
        "improvement": improvement(),
        "seams": seams(),
        "predictions": {
            "grids": list(GRIDS),
            "PT1": sum(r["prediction"] == "PT1" for r in records),
            "PT2": sum(r["prediction"] == "PT2" for r in records),
            "PT3_unchanged_rows": unchanged,
            "rows": records,
        },
        "holdout_read": False,
        "runtime_s": round(time.time() - start, 1),
    }
    return result


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
            print("FAIL: production_predictions.json differs from a recomputation")
            return 1
        print("PASS: production_predictions.json reproduces")
        return 0
    with open(OUT, "w", encoding="utf-8", newline="\n") as f:
        json.dump(result, f, indent=1)
        f.write("\n")
    m = result["measurement"]
    print(f"V: {m['points']} points, worst {m['worst']:.3e} at t={m['worst_at']['t']}, "
          f"df={m['worst_at']['df']:.0e} ({m['worst_at_100_digits']:.3e} at 100 digits); "
          f"within 2^-50: {m['within_bound']}")
    p = result["predictions"]
    print(f"predictions: PT1 {p['PT1']}, PT2 {p['PT2']}, PT3 unchanged {p['PT3_unchanged_rows']}")
    i = result["improvement"]
    print(f"improvement: {i['points']} study fitting points, route worst {i['route_worst']:.3e}, "
          f"worse than the current route at {i['route_worse_than_current'] or 'none'}")
    return 0 if m["within_bound"] and not i["route_worse_than_current"] else 1


if __name__ == "__main__":
    sys.exit(main())
