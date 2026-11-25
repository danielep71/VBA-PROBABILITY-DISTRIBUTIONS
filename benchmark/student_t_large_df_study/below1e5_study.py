"""
#365: the Student-t survival below df = 1E5, mapped in Excel.

Below df = 1E5 the large-df expansion and the near-one route of #213 do not
apply, so the survival comes from the incomplete-beta continued fraction at
x = df / (df + t^2). The binary64 route model puts its error above 1E-13
from df of about 560 up to the seam, worst near t = 2 just below 1E5. Excel
measured 3.16E-12 there (PRODUCTION_AMENDMENT_1.md), where the model gives
6.97E-12: the model is not bit-faithful, so the region has to be measured in
Excel before any rule is chosen.

This study adds StudentT_Survival(t, df) rows, regime "below1e5", on

  DFS   24 log-spaced df from 100 to 10^4.875 and the Double just below 1E5;
  TS    t in (0.5, 0.75, 1, 1.5, 2, 3, 4, 6, 8, 12, 16, 24, 40);

leaving out a point whose reference is below the smallest normal Double.
The references come from the oracle of the #210 seam study at 80 digits.
The rows carry no rule: --check reports the measured error, the binary64
model's error at the same point and the region above 1E-13, for the
preregistration of the fix.

    python below1e5_study.py --append   add the rows (blank observations), once
    python below1e5_study.py --check    report on the exported rows; write below1e5_results.json
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
import predict_production as PP  # noqa: E402
import seam_study as seam  # noqa: E402

GRID = seam.GRID
OUT = os.path.join(HERE, "below1e5_results.json")
REGIME = "below1e5"
DFS = tuple(round(10.0 ** (2.0 + k / 8.0), 6) for k in range(24)) + (math.nextafter(1e5, 0.0),)
TS = (0.5, 0.75, 1.0, 1.5, 2.0, 3.0, 4.0, 6.0, 8.0, 12.0, 16.0, 24.0, 40.0)
SMALLEST_NORMAL = mp.mpf(2) ** -1022
BOUND = 1e-13            # forward guarantee of #327, reported


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


def rows():
    return [{"function": "StudentT_Survival", "vba_kernel": "K_STATS_StudentT_Survival",
             "claim": "characterization", "metric": "rel", "arg1": repr(t), "arg2": repr(df),
             "arg3": "", "reference": mp.nstr(ref, 40), "observed_vba": "",
             "regime": REGIME, "evidence_set": "tchi_large_df"}
            for t, df, ref in points()]


def append():
    with open(GRID, newline="", encoding="utf-8") as f:
        text = f.read()
    if any(r["regime"] == REGIME for r in csv.DictReader(io.StringIO(text))):
        raise SystemExit(f"refusing: the grid already holds {REGIME} rows")
    fields = text.split("\n", 1)[0].strip().split(",")
    buf = io.StringIO()
    w = csv.DictWriter(buf, fieldnames=fields, lineterminator="\n")
    new = rows()
    for r in new:
        w.writerow({k: r.get(k, "") for k in fields})
    with open(GRID, "w", newline="", encoding="utf-8") as f:
        f.write(text if text.endswith("\n") else text + "\n")
        f.write(buf.getvalue())
    print(f"appended {len(new)} {REGIME} rows to {GRID}")


def check():
    import above8_model as M
    with open(GRID, encoding="utf-8-sig") as f:
        found = [r for r in csv.DictReader(f) if r["regime"] == REGIME]
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
    with open(OUT, "w", encoding="utf-8", newline="\n") as f:
        json.dump(result, f, indent=1)
        f.write("\n")
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
    mode.add_argument("--check", action="store_true")
    a = ap.parse_args(argv)
    if a.append:
        append()
        return 0
    return check()


if __name__ == "__main__":
    sys.exit(main())
