"""
Student-t behaviour across the large-df seams (#210).

The large-df route serves df >= 1E5 and |t| <= 8; the incomplete-beta continued
fraction serves the rest. This study measures, in Excel, what happens at the
Doubles on both sides of each seam:

  monotonicity  survival must not increase, and the cumulative must not
                decrease, as t grows (t seam) or, for t > 0, as df grows (df
                seam); for t < 0 the df direction reverses;
  complement    survival(t) + cumulative(t) - 1, in units of 2^-53;
  symmetry      survival(t) against cumulative(-t), bit for bit;
  round trips   inverse(p) against t, where p is the correctly rounded
                cumulative at t < 0 (the lower tail, so p carries the tail
                exactly), and against the exact quantile of p.

The rows are appended to tchi_large_df/tchi_large_df_grid.csv with regime
"seam", exported once by Export_TChiLargeDf, and read back here.

    python seam_study.py --append   add the seam rows (blank observations), once
    python seam_study.py --check    analyse the exported rows; write seam_results.json
"""
import argparse
import csv
import io
import json
import math
import os
import sys

import mpmath as mp

HERE = os.path.dirname(os.path.abspath(__file__))
BENCH = os.path.dirname(HERE)
sys.path.insert(0, HERE)
import predict_production as PP  # noqa: E402

GRID = os.path.join(BENCH, "tchi_large_df", "tchi_large_df_grid.csv")
OUT = os.path.join(HERE, "seam_results.json")
REGIME = "seam"
DF_SEAM = 1e5
T_SEAM = 8.0
DF_SEAM_T = (0.5, 1.0, 2.0, 4.0, 8.0)
T_SEAM_DF = (1e5, 1e6, 1e7, 1e8)
U = 2.0 ** -53


def df_neighbours():
    return (math.nextafter(DF_SEAM, 0.0), DF_SEAM, math.nextafter(DF_SEAM, math.inf))


def t_neighbours():
    return (math.nextafter(T_SEAM, 0.0), T_SEAM, math.nextafter(T_SEAM, math.inf))


def points():
    """(t, df) with t > 0, in the order the rows are written."""
    out = [(t, df) for t in DF_SEAM_T for df in df_neighbours()]
    out += [(t, df) for df in T_SEAM_DF for t in t_neighbours() if (t, df) not in out]
    return out


def upper_tail(t, df):
    """I_z(df/2, 1/2) / 2 at the binary64 arguments, the study's primary oracle."""
    with mp.workdps(80):
        t, df = mp.mpf(t), mp.mpf(df)
        return mp.betainc(df / 2, mp.mpf(1) / 2, 0, df / (df + t * t), regularized=True) / 2


def quantile(p, df):
    """The exact quantile of the binary64 probability p < 0.5 (negative)."""
    with mp.workdps(80):
        target = mp.mpf(p)
        seed = -mp.sqrt(2) * mp.erfinv(1 - 2 * target)
        return mp.findroot(lambda x: upper_tail(-x, df) - target, seed)


def rows():
    out = []

    def add(fn, a1, a2, ref):
        out.append({"function": fn, "vba_kernel": "K_STATS_" + fn, "claim": "characterization",
                    "metric": "rel", "arg1": repr(a1), "arg2": repr(a2), "arg3": "",
                    "reference": mp.nstr(ref, 40), "observed_vba": "",
                    "regime": REGIME, "evidence_set": "tchi_large_df"})

    with mp.workdps(80):
        for t, df in points():
            tail = upper_tail(t, df)
            for sign in (1.0, -1.0):
                x = sign * t
                add("StudentT_Survival", x, df, tail if sign > 0 else 1 - tail)
                add("StudentT_Cumulative", x, df, 1 - tail if sign > 0 else tail)
        for t, df in points():
            p = float(upper_tail(t, df))          # correctly rounded cumulative at -t
            add("StudentT_InverseCumulative", p, df, quantile(p, df))
    return out


def append():
    with open(GRID, newline="", encoding="utf-8") as f:
        text = f.read()
    if any(r["regime"] == REGIME for r in csv.DictReader(io.StringIO(text))):
        raise SystemExit("refusing: the grid already holds seam rows")
    fields = text.split("\n", 1)[0].strip().split(",")
    buf = io.StringIO()
    w = csv.DictWriter(buf, fieldnames=fields, lineterminator="\n")
    new = rows()
    for r in new:
        w.writerow({k: r.get(k, "") for k in fields})
    with open(GRID, "w", newline="", encoding="utf-8") as f:
        f.write(text if text.endswith("\n") else text + "\n")
        f.write(buf.getvalue())
    print(f"appended {len(new)} seam rows to {GRID}")


def observed_grid():
    found = {}
    with open(GRID, encoding="utf-8-sig") as f:
        for r in csv.DictReader(f):
            if r["regime"] == REGIME:
                found[(r["function"], float(r["arg1"]), float(r["arg2"]))] = r
    return found


def value(row):
    v = PP.observed(row["observed_vba"])
    if v is None:
        raise SystemExit(f"no observation for {row['function']}({row['arg1']}, {row['arg2']}): "
                         "export the grid first")
    return v


def rel_error(v, reference):
    with mp.workdps(60):
        reference = mp.mpf(reference)
        return float(abs((mp.mpf(v) - reference) / reference))


def monotone(series, increasing):
    """The largest relative step against the required direction, 0 when monotone."""
    worst = 0.0
    for (_, a), (_, b) in zip(series, series[1:]):
        step = (a - b) if increasing else (b - a)        # > 0 means a violation
        if step > 0:
            worst = max(worst, step / max(abs(a), abs(b)))
    return worst


def check():
    g = observed_grid()

    def S(t, df):
        return value(g[("StudentT_Survival", t, df)])

    def C(t, df):
        return value(g[("StudentT_Cumulative", t, df)])

    result = {"issue": "#210", "rows": len(g), "df_seam": [], "t_seam": [],
              "complement": {}, "symmetry": {}, "round_trips": [], "errors": {}}
    for t in DF_SEAM_T:
        for sign in (1.0, -1.0):
            x = sign * t
            dfs = df_neighbours()
            s = [(df, S(x, df)) for df in dfs]
            c = [(df, C(x, df)) for df in dfs]
            # For t > 0 the survival falls and the cumulative rises as df grows.
            result["df_seam"].append({
                "t": x, "survival": [v for _, v in s], "cumulative": [v for _, v in c],
                "survival_violation": monotone(s, increasing=sign < 0),
                "cumulative_violation": monotone(c, increasing=sign > 0)})
    for df in T_SEAM_DF:
        for sign in (1.0, -1.0):
            ts = [sign * t for t in t_neighbours()]
            ts.sort()
            s = [(x, S(x, df)) for x in ts]
            c = [(x, C(x, df)) for x in ts]
            result["t_seam"].append({
                "df": df, "t": ts, "survival": [v for _, v in s], "cumulative": [v for _, v in c],
                "survival_violation": monotone(s, increasing=False),
                "cumulative_violation": monotone(c, increasing=True)})
    complement, symmetry = [], []
    for t, df in points():
        for x in (t, -t):
            with mp.workdps(60):
                complement.append(float(abs(mp.mpf(S(x, df)) + mp.mpf(C(x, df)) - 1) / U))
            symmetry.append(S(x, df) == C(-x, df))
    result["complement"] = {"pairs": len(complement), "worst_units_of_2^-53": max(complement)}
    result["symmetry"] = {"pairs": len(symmetry), "bit_identical": sum(symmetry)}
    for t, df in points():
        p = float(upper_tail(t, df))
        row = g[("StudentT_InverseCumulative", p, df)]
        q = value(row)
        result["round_trips"].append({
            "t": -t, "df": df, "p": p, "observed": q,
            "round_trip_rel": abs(q + t) / t,
            "inverse_rel": rel_error(q, row["reference"])})
    forward = {}
    for (fn, a1, a2), row in g.items():
        if fn != "StudentT_InverseCumulative":
            forward.setdefault(fn, []).append(rel_error(value(row), row["reference"]))
    result["errors"] = {fn: max(v) for fn, v in forward.items()}
    result["worst_monotonicity_violation"] = max(
        max(r["survival_violation"], r["cumulative_violation"])
        for r in result["df_seam"] + result["t_seam"])
    with open(OUT, "w", encoding="utf-8", newline="\n") as f:
        json.dump(result, f, indent=1)
        f.write("\n")
    summary = {k: result[k] for k in ("rows", "complement", "symmetry", "errors",
                                      "worst_monotonicity_violation")}
    summary["worst_round_trip_rel"] = max(r["round_trip_rel"] for r in result["round_trips"])
    summary["worst_inverse_rel"] = max(r["inverse_rel"] for r in result["round_trips"])
    print(json.dumps(summary, indent=1))
    return 0


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    mode = ap.add_mutually_exclusive_group(required=True)
    mode.add_argument("--append", action="store_true")
    mode.add_argument("--check", action="store_true")
    a = ap.parse_args(argv)
    if a.append:
        append()
        return 0
    return check()


if __name__ == "__main__":
    sys.exit(main())
