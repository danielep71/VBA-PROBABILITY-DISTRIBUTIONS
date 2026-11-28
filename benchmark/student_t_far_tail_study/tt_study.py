"""
#365 (PREREGISTRATION.md): the holdout of the Student t tail beyond |t| = 16
below df = 1E5 carried as pairs of Doubles, and the judge of its rules.

The grid is benchmark/excel_function_map/student_t_far_tail_holdout_grid.csv,
in the format of the Excel floor grid, exported by
Export_StudentTFarTailHoldout as the study export xtt_tail. No point of it is
in an earlier holdout.

Classes of points:

  sv  K_STATS_StudentT_Survival where the rule answers;
  cu  K_STATS_StudentT_Cumulative at the negative of such an argument;
  in  K_STATS_StudentT_InverseCumulative at a probability whose quantile is
      beyond -16;
  ed  the survival where the rule does not answer: at |t| = 16 and at
      df = 1E5.

Rules, on the exported grids (--check), against the export of the source of
before (tt_before.json, written by --freeze-before):

  T1  accuracy    every sv and cu point is within ACCURACY of the reference;
  T2  floor       no sv or cu point is inferior to Excel, every row holds an
                  observation and no library value is ERROR;
  T3  the defect  each comparison row of the issue with t above 16 (the
                  below1e5_excel rows of the large-df grid) is within
                  ACCURACY;
  T4  inverse     every in point is within INVERSE and is not newly inferior
                  to Excel;
  T5  edges       no ed observation changes;
  T6  elsewhere   in the floor grid of #493 and in the six earlier holdouts no
                  point is newly inferior to Excel with a library error above
                  LOSS, and none has more inferior points than before.

    python tt_study.py --build           write the holdout grid (blank observations)
    python tt_study.py --build --check   fail if the committed grid's keys differ
    python tt_study.py --model           the rule against today on the holdout, in the model
    python tt_study.py --freeze-before   keep the verdicts of the source of before
    python tt_study.py --check           judge the exported grids; exit 0 when T1-T6 hold
"""
import argparse
import json
import math
import os
import sys

import mpmath as mp

HERE = os.path.dirname(os.path.abspath(__file__))
BENCH = os.path.dirname(HERE)
sys.path.insert(0, os.path.join(BENCH, "negative_binomial_pair_study"))
sys.path.insert(0, os.path.join(BENCH, "student_t_large_df_study"))
sys.path.insert(0, HERE)
import nb_study as S7  # noqa: E402
import below1e5_study as B5S  # noqa: E402
import tt_model as TT  # noqa: E402

S1, P = S7.S1, S7.P
GRID = os.path.join(S1.MAP, "student_t_far_tail_holdout_grid.csv")
BEFORE = os.path.join(HERE, "tt_before.json")
ISSUE = "#365"
ACCURACY = 1e-14         # T1, T3: the floor of CONTRIBUTING.md
INVERSE = 1e-12          # T4: the inverse guarantee of #327
LOSS = 3e-14             # T6: a newly inferior point above it is a real loss
ELSEWHERE = S7.ELSEWHERE + (("holdout of #504", S7.GRID),)
SURVIVAL_DF = (5.0, 23.0, 137.0, 420.0, 911.0, 1500.0, 1997.0, 2000.0, 3100.0, 7400.0,
               20000.0, 61000.0, 99998.0)
SURVIVAL_T = (16.03125, 17.5, 19.0, 22.25, 27.0, 33.0, 45.0, 62.0, 90.0, 150.0, 400.0,
              2500.0, 1e30, 1e50)
CUMULATIVE_T = (17.5, 27.0, 62.0)
INVERSE_P = (1e-70, 1e-150, 1e-270)
INVERSE_DF = (23.0, 420.0, 3100.0, 20000.0, 99998.0)
EDGE_T = 16.0
EDGE_T_DF = (137.0, 2000.0, 20000.0, 99998.0)
EDGE_DF = 100000.0
EDGE_DF_T = (17.5, 22.25, 27.0)
LOG_SMALLEST = -290.0 * math.log(10.0)


def _survival(t, df):
    return S1.REFERENCE["StudentT_Survival"](t, df)


def _quantile(p, df):
    """The quantile of the lower-tail probability p: negative.

    The root is taken on the logarithm of the tail. On the tail itself a
    residual of the size of p, far below the working precision, stops the
    search wherever it starts.
    """
    seed = -mp.sqrt(df * (mp.exp(-2 * mp.log(p) / df) - 1)) if df < 200 else \
        -mp.sqrt(-2 * mp.log(p))
    target = mp.log(p)
    root = mp.findroot(lambda x: mp.log(_survival(-x, df)) - target, seed)
    if abs(_survival(-root, df) / p - 1) > mp.mpf(10) ** -40:
        raise SystemExit(f"no quantile for p = {p}, df = {df}")
    return root


S1.REFERENCE.setdefault("StudentT_Cumulative", lambda t, df: _survival(-t, df))
S1.REFERENCE.setdefault("StudentT_InverseCumulative", _quantile)


def _representable(t, df):
    """A cheap bound that keeps the reference away from values far below 1E-290."""
    return -0.5 * df * math.log1p(t / df * t) > LOG_SMALLEST - 20.0


def design():
    """(class, function, arguments)."""
    out = []

    def add(cls, fn, *args):
        args = tuple(float(a) for a in args)
        with mp.workdps(S1.DPS):
            ref = S1.REFERENCE[fn](*[mp.mpf(a) for a in args])
        if mp.isfinite(ref) and abs(ref) > S1.SMALLEST and (cls, fn, args) not in out:
            out.append((cls, fn, args))

    for df in SURVIVAL_DF:
        for t in SURVIVAL_T:
            if _representable(t, df):
                add("sv", "StudentT_Survival", t, df)
                if t in CUMULATIVE_T:
                    add("cu", "StudentT_Cumulative", -t, df)
    for df in INVERSE_DF:
        for p in INVERSE_P:
            with mp.workdps(S1.DPS):
                beyond = _survival(mp.mpf(EDGE_T), mp.mpf(df)) > p
            if beyond:
                add("in", "StudentT_InverseCumulative", p, df)
    for df in EDGE_T_DF:
        add("ed", "StudentT_Survival", EDGE_T, df)
    for t in EDGE_DF_T:
        add("ed", "StudentT_Survival", t, EDGE_DF)
    return out


def _with_design(function, *args):
    saved = S1.design, S1.GRID
    S1.design, S1.GRID = design, GRID
    try:
        return function(*args)
    finally:
        S1.design, S1.GRID = saved


def overlap():
    """Points of this design that are in an earlier holdout: must be none."""
    earlier = set()
    for d in (S1.design, S7.S2.design, S7.S4.design, S7.S5.design, S7.S6.design, S7.design):
        earlier |= {(fn, args) for _, fn, args in d()}
    return [(fn, args) for _, fn, args in design() if (fn, args) in earlier]


def model():
    by_class = {}
    for cls, fn, args in design():
        if cls not in ("sv", "cu"):
            continue
        t, df = abs(args[0]), args[1]
        with mp.workdps(60):
            ref = _survival(mp.mpf(t), mp.mpf(df))
        s = by_class.setdefault(cls, [0, 0.0, 0.0, 0, None])
        today, rule = P.rel(TT.tail(t, df, rule=False), ref), P.rel(TT.tail(t, df), ref)
        s[0] += 1
        s[1] = max(s[1], today)
        if rule >= s[2]:
            s[2], s[4] = rule, (fn, args)
        s[3] += today > 1e-13
    for cls, (n, today, rule, above, where) in sorted(by_class.items()):
        print(f"  {cls}  {n:>3} points: today worst {today:.3g}, {above} above 1E-13; "
              f"rule worst {rule:.3g} at {where[0]}{where[1]}")
    return 0


def defect_points():
    """{(t, df): library relative error} on the comparison rows of the issue with t > 16."""
    out = {}
    for r in B5S.grid_rows(B5S.EXCEL):
        t, df = float(r["arg1"]), float(r["arg2"])
        if r["function"] == B5S.EXCEL_FUNCTION or t <= TT.T_MIN:
            continue
        value = B5S.PP.observed(r["observed_vba"])
        out[f"{t:g};{df:g}"] = None if value is None else B5S.seam.rel_error(value, r["reference"])
    return out


def freeze_before():
    data = {"issue": ISSUE, "holdout": S1._slim(S1.verdicts(GRID)),
            "defect": defect_points(),
            "elsewhere": {name: S1._slim(S1.verdicts(path)) for name, path in ELSEWHERE}}
    with open(BEFORE, "w", encoding="utf-8", newline="\n") as f:
        json.dump(data, f, indent=1)
        f.write("\n")
    print(f"kept {len(data['holdout'])} holdout verdicts, {len(data['defect'])} rows of the "
          f"issue and the verdicts of {len(data['elsewhere'])} other grids in {BEFORE}")
    return 0


def judge(holdout, defect, elsewhere, before):
    """(failures, report); elsewhere maps a grid's name to its verdicts after."""
    failures, report = [], {"classes": {}, "new": []}

    def inferior(points):
        return {k for k, p in points.items() if p["verdict"] == "inferior"}

    for k, p in sorted(holdout.items()):
        cls = k.split(":")[0]
        c = report["classes"].setdefault(cls, {
            "points": 0, "worst": 0.0, "worst_before": 0.0, "inferior": 0, "inferior_before": 0,
            "changed": 0})
        c["points"] += 1
        prior = before["holdout"][k]
        c["worst_before"] = max(c["worst_before"], prior["library_rel"] or 0.0)
        c["inferior_before"] += prior["verdict"] == "inferior"
        if p["verdict"] == "library error" or p.get("library_missing"):
            failures.append(f"T2 {p['kstats_call']}: the library returns an error")
            continue
        e = p.get("library_rel")
        if e is None:
            failures.append(f"T2 {p['kstats_call']}: no relative error")
            continue
        c["worst"] = max(c["worst"], e)
        c["inferior"] += p["verdict"] == "inferior"
        c["changed"] += e != prior["library_rel"]
        if cls == "ed":
            if e != prior["library_rel"]:
                failures.append(f"T5 {p['kstats_call']}: the observation changed where the "
                                "rule does not answer")
        elif cls == "in":
            if e > INVERSE:
                failures.append(f"T4 {p['kstats_call']}: {e:.3g} above {INVERSE:g}")
            if p["verdict"] == "inferior" and prior["verdict"] != "inferior":
                failures.append(f"T4 {p['kstats_call']}: newly inferior to Excel ({e:.3g})")
        else:
            if e > ACCURACY:
                failures.append(f"T1 {p['kstats_call']}: {e:.3g} above {ACCURACY:g}")
            if p["verdict"] == "inferior":
                failures.append(f"T2 {p['kstats_call']}: inferior to Excel ({e:.3g}, "
                                f"Excel {p.get('excel_rel', 0.0):.3g})")
    worst = 0.0
    for k in sorted(before["defect"]):
        e = defect.get(k)
        if e is None or e > ACCURACY:
            failures.append(f"T3 t;df = {k}: {e} above {ACCURACY:g}")
        worst = max(worst, 1.0 if e is None else e)
    report["defect"] = {"rows": len(before["defect"]), "worst": worst,
                        "worst_before": max(v or 0.0 for v in before["defect"].values())}
    for name, after in elsewhere.items():
        prior = before["elsewhere"][name]
        now, then = inferior(after), inferior(prior)
        report[name] = {"inferior_before": len(then), "inferior_after": len(now),
                        "changed": sum(after[k].get("library_rel") != prior[k]["library_rel"]
                                       for k in after if k in prior)}
        for k in sorted(now - then):
            p = after[k]
            e = p.get("library_rel", 1.0)
            report["new"].append((name, p["kstats_call"], prior[k]["library_rel"], e,
                                  p.get("excel_rel")))
            if e > LOSS:
                failures.append(f"T6 {p['kstats_call']}: newly inferior in the {name} at "
                                f"{e:.3g}, above {LOSS:g}")
        if len(now) > len(then):
            failures.append(f"T6 {len(now)} points of the {name} inferior, {len(then)} before")
    return failures, report


def check():
    with open(BEFORE, encoding="utf-8") as f:
        before = json.load(f)
    failures, report = judge(S1.verdicts(GRID), defect_points(),
                             {name: S1.verdicts(path) for name, path in ELSEWHERE}, before)
    for cls, c in sorted(report["classes"].items()):
        print(f"  {cls}  {c['points']:>3} points: worst {c['worst']:.3g} (before "
              f"{c['worst_before']:.3g}); inferior to Excel {c['inferior']} (before "
              f"{c['inferior_before']}); changed {c['changed']}")
    d = report["defect"]
    print(f"  the {d['rows']} rows of {ISSUE} with t above 16: worst {d['worst']:.3g} "
          f"(before {d['worst_before']:.3g})")
    for name, _ in ELSEWHERE:
        r = report[name]
        print(f"  {name}: inferior to Excel {r['inferior_after']} (before "
              f"{r['inferior_before']}); changed {r['changed']}")
    for name, call, was, now, excel in report["new"]:
        print(f"  newly inferior ({name}): {call}  {was:.3g} -> {now:.3g}  Excel {excel:.3g}")
    for text in failures:
        print(f"FAIL: {text}")
    if failures:
        print(f"REJECTED: {len(failures)} finding(s)")
        return 1
    print("ACCEPTED: T1 to T6 hold")
    return 0


def main(argv=None):
    ap = argparse.ArgumentParser(description=(__doc__ or "").split("\n\n")[0])
    ap.add_argument("--build", action="store_true")
    ap.add_argument("--freeze-before", action="store_true")
    ap.add_argument("--check", action="store_true")
    ap.add_argument("--model", action="store_true")
    a = ap.parse_args(argv)
    if a.model:
        return model()
    if a.build:
        return _with_design(S1.build, a.check)
    if a.freeze_before:
        return freeze_before()
    if a.check:
        return check()
    ap.error("choose --build, --freeze-before, --check or --model")
    return 2


if __name__ == "__main__":
    sys.exit(main())
