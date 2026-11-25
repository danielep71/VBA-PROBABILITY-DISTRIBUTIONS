"""
#365: the Student-t survival below df = 1E5.

Below df = 1E5 the large-df expansion does not apply, and until #365 the
survival came from the incomplete-beta continued fraction at
x = df / (df + t^2) given as one Double. The study has two parts.

The map (step 1). StudentT_Survival(t, df) rows, regime "below1e5", on

  DFS   24 log-spaced df from 100 to 10^4.875 and the Double just below 1E5;
  TS    t in (0.5, 0.75, 1, 1.5, 2, 3, 4, 6, 8, 12, 16, 24, 40);

leaving out a point whose reference is below the smallest normal Double.
Exported from the source before the rule, they are the map of
below1e5_results.json; below1e5_before.json keeps their observations.

The rule (step 2, BELOW1E5_PREREGISTRATION.md). Rows added at registration:

  holdout   regime "below1e5_holdout": the survival at 12 df x 12 t inside
            the rule's region, the cumulative at -t for three t, and the
            inverse at the correctly rounded lower-tail probability of two t;
  seam      regime "below1e5_seam": the survival at the Doubles on either
            side of |t| = 16 and of df = 1999, where the rule begins.

Rules, on the exported grid (--check):

  G1  forward     every map and holdout forward row the rule serves is within
                  FORWARD of the reference;
  G2  unchanged   every map row the rule does not serve holds the observation
                  of below1e5_before.json, bit for bit;
  G3  df seam     across df = 1E5, at t = 2, 4 and 8 and both signs (seam rows
                  of #210), the survival and the cumulative move against the
                  required direction by at most SEAM;
  G4  inverse     every holdout inverse is within INVERSE of the reference;
  G5  numbers     every row of the three regimes holds a number.

Reported, not gating: the steps across |t| = 16 and df = 1999.

The comparison with Excel (CONTRIBUTING.md, "Excel as the floor"). Rows of
regime "below1e5_excel" hold, at 12 integer df (Excel truncates df) and 12 t,
the library's StudentT_Survival and Excel's own T.DIST.RT, written in the
grid as the function Excel_T_DIST_RT, against the same reference. --compare
lists every point where Excel is closer to the reference than the library.

    python below1e5_study.py --append            add the map rows, once (step 1)
    python below1e5_study.py --append-holdout    add the holdout and seam rows, once
    python below1e5_study.py --map               report the map rows against the model
    python below1e5_study.py --check             judge the exported rows; exit 0 when G1-G5 hold
    python below1e5_study.py --append-excel      add the comparison rows, once
    python below1e5_study.py --compare           library against Excel; exit 0 when Excel is nowhere closer
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
sys.path.insert(0, HERE)
import above8_study as above8  # noqa: E402
import below1e5_model as model  # noqa: E402
import predict_production as PP  # noqa: E402
import seam_study as seam  # noqa: E402

GRID = seam.GRID
OUT = os.path.join(HERE, "below1e5_results.json")
REGIME = "below1e5"
DFS = tuple(round(10.0 ** (2.0 + k / 8.0), 6) for k in range(24)) + (math.nextafter(1e5, 0.0),)
TS = (0.5, 0.75, 1.0, 1.5, 2.0, 3.0, 4.0, 6.0, 8.0, 12.0, 16.0, 24.0, 40.0)
SMALLEST_NORMAL = mp.mpf(2) ** -1022
BOUND = 1e-13            # the level the map reports against

BEFORE = os.path.join(HERE, "below1e5_before.json")
HOLDOUT = "below1e5_holdout"
SEAM = "below1e5_seam"
HOLDOUT_DF = (2048.0, 2999.5, 4500.25, 6789.0, 9999.0, 15000.5, 22222.0, 33333.25,
              50000.0, 65536.0, 88888.5, 99999.0)
HOLDOUT_T = (1.74, 1.9, 2.25, 2.75, 3.3, 4.7, 6.2, 9.1, 11.5, 13.6, 15.9, 16.0)
CUMULATIVE_T = (2.25, 6.2, 13.6)
INVERSE_T = (2.25, 9.1)
INVERSE_DF = HOLDOUT_DF[::2]
T_SEAM = model.T_MAX
T_SEAM_DF = (2048.0, 9999.0, 50000.0)
DF_EDGE = 1999.0         # a + b = 1000, where PROB_TryBetaLowerNearOne begins
DF_EDGE_T = (4.0, 8.0)
DF_SEAM_T = (2.0, 4.0, 8.0)
EXCEL = "below1e5_excel"
EXCEL_FUNCTION = "Excel_T_DIST_RT"
EXCEL_OUT = os.path.join(HERE, "below1e5_excel.json")
EXCEL_DF = (100.0, 316.0, 562.0, 1000.0, 1778.0, 1998.0, 1999.0, 2371.0, 5623.0, 10000.0,
            42170.0, 99999.0)
EXCEL_T = (0.5, 1.0, 1.5, 2.0, 3.0, 4.0, 8.0, 12.0, 16.0, 17.0, 24.0, 40.0)
ULP = 2.0 ** -52         # one unit in the last place, relative: the rounding of a comparison
FORWARD = 1e-13          # G1
SEAM_STEP = 1e-13        # G3
INVERSE = 1e-12          # G4, the inverse guarantee of #327


def points():
    """(t, df, reference) in the order the rows are written."""
    out = []
    with mp.workdps(80):
        for df in DFS:
            for t in TS:
                ref = seam.upper_tail(t, df)
                if ref >= SMALLEST_NORMAL:
                    out.append((t, df, ref))
    return out


def _row(fn, a1, a2, ref, regime):
    return {"function": fn, "vba_kernel": "K_STATS_" + fn, "claim": "characterization",
            "metric": "rel", "arg1": repr(a1), "arg2": repr(a2), "arg3": "",
            "reference": mp.nstr(ref, 40), "observed_vba": "",
            "regime": regime, "evidence_set": "tchi_large_df"}


def map_rows():
    return [_row("StudentT_Survival", t, df, ref, REGIME) for t, df, ref in points()]


def t_neighbours():
    return (math.nextafter(T_SEAM, 0.0), T_SEAM, math.nextafter(T_SEAM, math.inf))


def df_neighbours():
    return (math.nextafter(DF_EDGE, 0.0), DF_EDGE, math.nextafter(DF_EDGE, math.inf))


def holdout_rows():
    out = []
    with mp.workdps(80):
        for df in HOLDOUT_DF:
            for t in HOLDOUT_T:
                out.append(_row("StudentT_Survival", t, df, seam.upper_tail(t, df), HOLDOUT))
        for df in HOLDOUT_DF:
            for t in CUMULATIVE_T:
                out.append(_row("StudentT_Cumulative", -t, df, seam.upper_tail(t, df), HOLDOUT))
        for df in INVERSE_DF:
            for t in INVERSE_T:
                p = float(seam.upper_tail(t, df))
                out.append(_row("StudentT_InverseCumulative", p, df,
                                above8.quantile(p, df, t), HOLDOUT))
    return out


def seam_rows():
    out = []
    with mp.workdps(80):
        for df in T_SEAM_DF:
            for t in t_neighbours():
                out.append(_row("StudentT_Survival", t, df, seam.upper_tail(t, df), SEAM))
        for t in DF_EDGE_T:
            for df in df_neighbours():
                out.append(_row("StudentT_Survival", t, df, seam.upper_tail(t, df), SEAM))
    return out


def excel_rows():
    """The library and Excel at the same integer-df points, in pairs."""
    out = []
    with mp.workdps(80):
        for df in EXCEL_DF:
            for t in EXCEL_T:
                ref = seam.upper_tail(t, df)
                if ref < SMALLEST_NORMAL:
                    continue
                out.append(_row("StudentT_Survival", t, df, ref, EXCEL))
                excel = _row(EXCEL_FUNCTION, t, df, ref, EXCEL)
                excel["vba_kernel"] = "T.DIST.RT"
                out.append(excel)
    return out


def rows():
    """Every row the study owns, in the order of the grid."""
    return map_rows() + holdout_rows() + seam_rows() + excel_rows()


def append(new, regimes):
    with open(GRID, newline="", encoding="utf-8") as f:
        text = f.read()
    if any(r["regime"] in regimes for r in csv.DictReader(io.StringIO(text))):
        raise SystemExit(f"refusing: the grid already holds {' or '.join(regimes)} rows")
    fields = text.split("\n", 1)[0].strip().split(",")
    buf = io.StringIO()
    w = csv.DictWriter(buf, fieldnames=fields, lineterminator="\n")
    for r in new:
        w.writerow({k: r.get(k, "") for k in fields})
    with open(GRID, "w", newline="", encoding="utf-8") as f:
        f.write(text if text.endswith("\n") else text + "\n")
        f.write(buf.getvalue())
    print(f"appended {len(new)} {' and '.join(regimes)} rows to {GRID}")


def grid_rows(regime):
    with open(GRID, encoding="utf-8-sig") as f:
        return [r for r in csv.DictReader(f) if r["regime"] == regime]


def key(row):
    return f"{row['function']}|{row['arg1']}|{row['arg2']}"


def freeze_before():
    """Keep the map rows' observations, exported before the rule, for G2."""
    found = grid_rows(REGIME)
    if any(PP.observed(r["observed_vba"]) is None for r in found):
        raise SystemExit("refusing: a map row holds no number")
    with open(BEFORE, "w", encoding="utf-8", newline="\n") as f:
        json.dump({"issue": "#365", "source": "tchi_large_df export at 35c3b78, before the rule",
                   "observed_vba": {key(r): r["observed_vba"] for r in found}}, f, indent=1)
        f.write("\n")
    print(f"kept {len(found)} observations in {BEFORE}")


def judge(map_found, holdout_found, seam_found, seam210, before):
    """(failures, report) from the exported rows; before maps key(row) to an observation."""
    failures = []
    report = {"served": [], "inverses": [], "df_seam": [], "t_seam": [], "df_edge": []}

    def number(row, rule):
        v = PP.observed(row["observed_vba"])
        if v is None:
            failures.append(f"{rule} {row['function']}({row['arg1']}, {row['arg2']}): no number")
        return v

    for r in map_found + holdout_found:
        v = number(r, "G5")
        if v is None:
            continue
        a1, df = float(r["arg1"]), float(r["arg2"])
        e = seam.rel_error(v, r["reference"])
        if r["function"] == "StudentT_InverseCumulative":
            report["inverses"].append((a1, df, e))
            if e > INVERSE:
                failures.append(f"G4 inverse({a1!r}, {df!r}): {e:.3g} above {INVERSE:g}")
        elif model.served(abs(a1), df):
            report["served"].append((a1, df, e))
            if e > FORWARD:
                failures.append(f"G1 {r['function']}({a1!r}, {df!r}): {e:.3g} above {FORWARD:g}")
        elif r["regime"] == REGIME:
            if PP.observed(before.get(key(r), "")) != v:
                failures.append(f"G2 survival({a1!r}, {df!r}): the observation changed")
        else:
            failures.append(f"a holdout row the rule does not serve: {key(r)}")
    table = {(r["function"], float(r["arg1"]), float(r["arg2"])): r for r in seam210}
    for t in DF_SEAM_T:
        for sign in (1.0, -1.0):
            x = sign * t
            for fn, increasing in (("StudentT_Survival", sign < 0), ("StudentT_Cumulative", sign > 0)):
                series = [(df, PP.observed(table[(fn, x, df)]["observed_vba"]))
                          for df in seam.df_neighbours()]
                if any(v is None for _, v in series):
                    failures.append(f"G3 {fn}({x!r}) across df = 1E5: no number")
                    continue
                worst = seam.monotone(series, increasing)
                report["df_seam"].append((fn, x, worst))
                if worst > SEAM_STEP:
                    failures.append(f"G3 {fn}({x!r}) across df = 1E5: {worst:.3g} "
                                    f"above {SEAM_STEP:g}")
    probes = {(float(r["arg1"]), float(r["arg2"])): number(r, "G5") for r in seam_found}
    if None not in probes.values():
        for df in T_SEAM_DF:
            series = [(t, probes[(t, df)]) for t in t_neighbours()]
            report["t_seam"].append((df, seam.monotone(series, increasing=False)))
        for t in DF_EDGE_T:
            series = [(df, probes[(t, df)]) for df in df_neighbours()]
            report["df_edge"].append((t, seam.monotone(series, increasing=False)))
    return failures, report


def check():
    holdout_found, seam_found = grid_rows(HOLDOUT), grid_rows(SEAM)
    if not holdout_found or not seam_found:
        print("FAIL: the grid holds no holdout or seam rows; run --append-holdout")
        return 1
    with open(BEFORE, encoding="utf-8") as f:
        before = json.load(f)["observed_vba"]
    failures, report = judge(grid_rows(REGIME), holdout_found, seam_found,
                             grid_rows(seam.REGIME), before)

    def worst(items):
        return max((e for *_, e in items), default=0.0)

    print(f"  forward rows the rule serves           {len(report['served'])}, "
          f"worst {worst(report['served']):.3g}")
    print(f"  holdout inverses                       {len(report['inverses'])}, "
          f"worst {worst(report['inverses']):.3g}")
    print(f"  across df = 1E5, worst step            {worst(report['df_seam']):.3g}")
    print(f"  across |t| = 16, worst step (reported) {worst(report['t_seam']):.3g}")
    print(f"  across df = 1999, worst step (reported) {worst(report['df_edge']):.3g}")
    for text in failures:
        print(f"FAIL: {text}")
    if failures:
        print(f"REJECTED: {len(failures)} finding(s)")
        return 1
    print("ACCEPTED: G1 to G5 hold")
    return 0


def compare():
    """Library against Excel on the comparison rows; exit 0 when Excel is nowhere closer."""
    found = grid_rows(EXCEL)
    if not found:
        print("FAIL: the grid holds no comparison rows; run --append-excel")
        return 1
    pairs = {}
    for r in found:
        side = "excel" if r["function"] == EXCEL_FUNCTION else "library"
        pairs.setdefault((float(r["arg1"]), float(r["arg2"])), {})[side] = r
    points, unusable, blank = [], [], []
    for (t, df), pair in sorted(pairs.items(), key=lambda kv: (kv[0][1], kv[0][0])):
        lib = PP.observed(pair["library"]["observed_vba"])
        raw = pair["excel"]["observed_vba"].strip()
        if lib is None or not raw:
            blank.append((t, df))
            continue
        xl = PP.observed(raw)
        if xl is None:
            unusable.append({"t": t, "df": df, "excel": raw})
            continue
        ref = pair["library"]["reference"]
        points.append({"t": t, "df": df, "library_rel": seam.rel_error(lib, ref),
                       "excel_rel": seam.rel_error(xl, ref), "served": model.served(t, df)})
    if blank:
        print(f"FAIL: {len(blank)} pair(s) hold no observation, first {blank[0]}: export the grid first")
        return 1
    closer = [p for p in points if p["excel_rel"] < p["library_rel"]]
    beyond = [p for p in closer if p["library_rel"] - p["excel_rel"] > ULP]
    result = {"issue": "#365", "pairs": len(pairs), "compared": len(points),
              "excel_unusable": unusable, "ulp": ULP,
              "excel_closer": len(closer), "excel_closer_beyond_one_ulp": len(beyond),
              "library_worst": max(p["library_rel"] for p in points),
              "excel_worst": max(p["excel_rel"] for p in points), "points": points}
    with open(EXCEL_OUT, "w", encoding="utf-8", newline="\n") as f:
        json.dump(result, f, indent=1)
        f.write("\n")
    print(f"  pairs                               {len(pairs)}, compared {len(points)}, "
          f"Excel unusable {len(unusable)}")
    print(f"  library worst                       {result['library_worst']:.3g}")
    print(f"  Excel worst                         {result['excel_worst']:.3g}")
    print(f"  Excel closer                        {len(closer)}, by more than one ulp {len(beyond)}")
    for p in beyond:
        print(f"    df = {p['df']:<8g} t = {p['t']:<5g} library {p['library_rel']:.3g}  "
              f"Excel {p['excel_rel']:.3g}  served {p['served']}")
    if beyond:
        print(f"EXCEL CLOSER at {len(beyond)} point(s)")
        return 1
    print("PASS: Excel is nowhere closer than the library by more than one ulp")
    return 0


def report_map():
    import above8_model as M
    found = grid_rows(REGIME)
    if not found:
        print(f"FAIL: the grid holds no {REGIME} rows; run --append")
        return 1
    points_out, missing = [], []
    for r in found:
        t, df = float(r["arg1"]), float(r["arg2"])
        v = PP.observed(r["observed_vba"])
        if v is None:
            missing.append((t, df))
            continue
        with mp.workdps(60):
            ref = mp.mpf(r["reference"])
            excel = float(abs((mp.mpf(v) - ref) / ref))
            model = float(abs((mp.mpf(M.current_tail(t, df)) - ref) / ref))
        points_out.append({"t": t, "df": df, "excel_rel": excel, "model_rel": model,
                           "excel_equals_model": v == M.current_tail(t, df)})
    if missing:
        print(f"FAIL: {len(missing)} row(s) hold no number, first {missing[0]}: export the grid first")
        return 1
    over = [p for p in points_out if p["excel_rel"] > BOUND]
    worst = max(points_out, key=lambda p: p["excel_rel"])
    result = {
        "issue": "#365", "rows": len(points_out), "bound": BOUND,
        "worst": worst, "above_bound": len(over),
        "smallest_df_above_bound": min((p["df"] for p in over), default=None),
        "bit_identical_to_model": sum(p["excel_equals_model"] for p in points_out),
        "points": points_out,
    }
    print(f"  rows                         {len(points_out)}")
    print(f"  worst                        {worst['excel_rel']:.3g} at t = {worst['t']:g}, "
          f"df = {worst['df']!r} (model {worst['model_rel']:.3g})")
    print(f"  above {BOUND:g}                 {len(over)}, smallest df "
          f"{result['smallest_df_above_bound']!r}")
    print(f"  bit-identical to the model   {result['bit_identical_to_model']}")
    for p in sorted(over, key=lambda p: (p["df"], p["t"])):
        print(f"    df = {p['df']:<14.10g} t = {p['t']:<5g} excel {p['excel_rel']:.3g}  "
              f"model {p['model_rel']:.3g}")
    return 0


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    mode = ap.add_mutually_exclusive_group(required=True)
    mode.add_argument("--append", action="store_true")
    mode.add_argument("--append-holdout", action="store_true")
    mode.add_argument("--freeze-before", action="store_true")
    mode.add_argument("--map", action="store_true")
    mode.add_argument("--check", action="store_true")
    mode.add_argument("--append-excel", action="store_true")
    mode.add_argument("--compare", action="store_true")
    a = ap.parse_args(argv)
    if a.append:
        append(map_rows(), (REGIME,))
        return 0
    if a.append_holdout:
        append(holdout_rows() + seam_rows(), (HOLDOUT, SEAM))
        return 0
    if a.freeze_before:
        freeze_before()
        return 0
    if a.append_excel:
        append(excel_rows(), (EXCEL,))
        return 0
    if a.compare:
        return compare()
    return report_map() if a.map else check()


if __name__ == "__main__":
    sys.exit(main())
