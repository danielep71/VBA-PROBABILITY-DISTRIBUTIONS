"""
#213: the Student-t tail above |t| = 8 at large df (ABOVE8_PREREGISTRATION.md).

Adds rows with regime "above8" to tchi_large_df/tchi_large_df_grid.csv and
judges them, with the seam rows of #210, once Excel has exported the grid.

  probes        StudentT_Survival(t, df) at t in PROBE_T and df in DFS;
  inverses      StudentT_InverseCumulative(p, df) at the correctly rounded
                lower-tail probability p of t in INVERSE_T.

Rules, on the exported grid:

  F1  forward     the survival at 8 + 1 ulp (seam rows) and at every probe
                  with t <= GATED_T is within FORWARD of the reference;
  F2  monotone    across |t| = 8 the survival does not rise and the
                  cumulative does not fall, at every df and both signs;
  F3  inverse     every seam round trip with df >= 1E5 and every above8
                  inverse is within INVERSE of the reference;
  F4  numbers     every above8 row holds a number.

Reported, not gating: the probes beyond GATED_T, and the inverse against
its 5E-15 measured target.

    python above8_study.py --append   add the rows (blank observations), once
    python above8_study.py --check    judge the exported rows; exit 0 when F1-F4 hold
"""
import argparse
import csv
import io
import math
import os
import sys

import mpmath as mp

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import predict_production as PP  # noqa: E402
import seam_study as seam  # noqa: E402

GRID = seam.GRID
REGIME = "above8"
DFS = (1e5, 1e6, 1e7, 1e8)
PROBE_T = (8.25, 9.0, 10.0, 14.0, 20.0, 30.0)
INVERSE_T = (8.25, 10.0)
GATED_T = 14.0
FORWARD = 1e-13          # forward guarantee of #327
INVERSE = 1e-12          # inverse guarantee of #327
TARGET = 5e-15           # measured target of the inverse contract, reported


def rows():
    out = []

    def add(fn, a1, a2, ref):
        out.append({"function": fn, "vba_kernel": "K_STATS_" + fn, "claim": "characterization",
                    "metric": "rel", "arg1": repr(a1), "arg2": repr(a2), "arg3": "",
                    "reference": mp.nstr(ref, 40), "observed_vba": "",
                    "regime": REGIME, "evidence_set": "tchi_large_df"})

    with mp.workdps(80):
        for df in DFS:
            for t in PROBE_T:
                add("StudentT_Survival", t, df, seam.upper_tail(t, df))
        for df in DFS:
            for t in INVERSE_T:
                p = float(seam.upper_tail(t, df))
                add("StudentT_InverseCumulative", p, df, quantile(p, df, t))
    return out


def quantile(p, df, near):
    """The exact quantile (negative) of the binary64 lower-tail probability p."""
    with mp.workdps(80):
        target = mp.log(mp.mpf(p))
        root = mp.findroot(lambda x: mp.log(seam.upper_tail(x, df)) - target, mp.mpf(near),
                           tol=mp.mpf(10) ** -50)
        return -root


def append():
    with open(GRID, newline="", encoding="utf-8") as f:
        text = f.read()
    if any(r["regime"] == REGIME for r in csv.DictReader(io.StringIO(text))):
        raise SystemExit("refusing: the grid already holds above8 rows")
    fields = text.split("\n", 1)[0].strip().split(",")
    buf = io.StringIO()
    w = csv.DictWriter(buf, fieldnames=fields, lineterminator="\n")
    new = rows()
    for r in new:
        w.writerow({k: r.get(k, "") for k in fields})
    with open(GRID, "w", newline="", encoding="utf-8") as f:
        f.write(text if text.endswith("\n") else text + "\n")
        f.write(buf.getvalue())
    print(f"appended {len(new)} above8 rows to {GRID}")


def grid_rows(regime):
    with open(GRID, encoding="utf-8-sig") as f:
        return [r for r in csv.DictReader(f) if r["regime"] == regime]


def judge(seam_rows, probe_rows):
    """(failures, report) from the exported seam and above8 rows."""
    failures, report = [], {"probes": [], "inverses": [], "seam_inverses": [], "seam_forward": []}
    above = math.nextafter(seam.T_SEAM, math.inf)
    table = {(r["function"], float(r["arg1"]), float(r["arg2"])): r for r in seam_rows}

    def number(row, rule):
        v = PP.observed(row["observed_vba"])
        if v is None:
            failures.append(f"{rule} {row['function']}({row['arg1']}, {row['arg2']}): no number")
        return v

    def rel(v, row):
        return seam.rel_error(v, row["reference"])

    # F1 at the seam: the small side of each pair just above 8.
    for df in seam.T_SEAM_DF:
        for fn, x in (("StudentT_Survival", above), ("StudentT_Cumulative", -above)):
            row = table[(fn, x, df)]
            v = number(row, "F1")
            if v is None:
                continue
            e = rel(v, row)
            report["seam_forward"].append((fn, x, df, e))
            if e > FORWARD:
                failures.append(f"F1 {fn}({x!r}, {df:g}): {e:.3g} above {FORWARD:g}")
    # F2: monotone across |t| = 8.
    for df in seam.T_SEAM_DF:
        for sign in (1.0, -1.0):
            ts = sorted(sign * t for t in seam.t_neighbours())
            for fn, increasing in (("StudentT_Survival", False), ("StudentT_Cumulative", True)):
                series = [(x, PP.observed(table[(fn, x, df)]["observed_vba"])) for x in ts]
                if any(v is None for _, v in series):
                    failures.append(f"F2 {fn} at df {df:g}: no number")
                    continue
                worst = seam.monotone(series, increasing)
                if worst > 0:
                    failures.append(f"F2 {fn} at df {df:g}, sign {sign:+g}: "
                                    f"not monotone by {worst:.3g}")
    # F3 at the seam: the round trips at df >= 1E5.
    for r in seam_rows:
        if r["function"] != "StudentT_InverseCumulative" or float(r["arg2"]) < 1e5:
            continue
        v = number(r, "F3")
        if v is None:
            continue
        e = rel(v, r)
        report["seam_inverses"].append((float(r["arg1"]), float(r["arg2"]), e))
        if e > INVERSE:
            failures.append(f"F3 inverse({r['arg1']}, {r['arg2']}): {e:.3g} above {INVERSE:g}")
    # The above8 rows.
    for r in probe_rows:
        v = number(r, "F4")
        if v is None:
            continue
        e = rel(v, r)
        a1, df = float(r["arg1"]), float(r["arg2"])
        if r["function"] == "StudentT_Survival":
            report["probes"].append((a1, df, e))
            if a1 <= GATED_T and e > FORWARD:
                failures.append(f"F1 survival({a1!r}, {df:g}): {e:.3g} above {FORWARD:g}")
        else:
            report["inverses"].append((a1, df, e))
            if e > INVERSE:
                failures.append(f"F3 inverse({a1!r}, {df:g}): {e:.3g} above {INVERSE:g}")
    return failures, report


def check():
    probe_rows = grid_rows(REGIME)
    if not probe_rows:
        print("FAIL: the grid holds no above8 rows; run --append")
        return 1
    failures, report = judge(grid_rows(seam.REGIME), probe_rows)

    def worst(items):
        return max((e for *_, e in items), default=0.0)

    gated = [p for p in report["probes"] if p[0] <= GATED_T]
    beyond = [p for p in report["probes"] if p[0] > GATED_T]
    print(f"  forward at 8 + 1 ulp, worst            {worst(report['seam_forward']):.3g}")
    print(f"  probes with t <= {GATED_T:g}, worst           {worst(gated):.3g}")
    print(f"  probes beyond, worst (reported)        {worst(beyond):.3g}")
    print(f"  seam round trips, df >= 1E5, worst     {worst(report['seam_inverses']):.3g}")
    print(f"  above8 inverses, worst                 {worst(report['inverses']):.3g}")
    over = sum(1 for *_, e in report["seam_inverses"] + report["inverses"] if e > TARGET)
    print(f"  inverses above the {TARGET:g} target (reported): {over}")
    for text in failures:
        print(f"FAIL: {text}")
    if failures:
        print(f"REJECTED: {len(failures)} finding(s)")
        return 1
    print("ACCEPTED: F1 to F4 hold")
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
