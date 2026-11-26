"""
#501 (PREREGISTRATION.md): the holdout of the Beta, F and Student t densities
carried as pairs of Doubles, and the judge of its rules.

The grid is benchmark/excel_function_map/density_pair_holdout_grid.csv, in the
format of the Excel floor grid, exported by Export_DensityPairHoldout as the
study export xdn_pair. No point of it is in a holdout of #175 or #498.

Rules, on the exported grids (--check), against the export of the source of
before (density_before.json, written by --freeze-before):

  L1  accuracy    every holdout point is within ACCURACY of the reference;
  L2  floor       no holdout point is inferior to Excel, every row holds an
                  observation and no library value is ERROR;
  L3  the defect  each of the ten points of the issue, rows of the two
                  holdouts of #175, is within ACCURACY;
  L4  elsewhere   in the floor grid of #493, in the two holdouts of #175 and
                  in the holdout of #498 no point is newly inferior to Excel
                  with a library error above LOSS, and none has more inferior
                  points than before.

Reported, not gating: the holdout before and after, by class.

    python density_study.py --build           write the holdout grid (blank observations)
    python density_study.py --build --check   fail if the committed grid's keys differ
    python density_study.py --model           the rule against today on the holdout, in the model
    python density_study.py --freeze-before   keep the verdicts of the source of before
    python density_study.py --check           judge the exported grids; exit 0 when L1-L4 hold
"""
import argparse
import json
import math
import os
import sys

import mpmath as mp

HERE = os.path.dirname(os.path.abspath(__file__))
BENCH = os.path.dirname(HERE)
sys.path.insert(0, os.path.join(BENCH, "deviance_pair_study"))
sys.path.insert(0, HERE)
import density_model as D  # noqa: E402
import pair_study as S4  # noqa: E402

S1, S2, P = S4.S1, S4.S2, S4.P
GRID = os.path.join(S1.MAP, "density_pair_holdout_grid.csv")
BEFORE = os.path.join(HERE, "density_before.json")
ISSUE = "#501"
ACCURACY = 1e-14         # L1, L3: the floor of CONTRIBUTING.md
LOSS = 3e-14             # L4: a newly inferior point above it is a real loss
ELSEWHERE = (("floor grid", S1.FB.GRID), ("first holdout of #175", S1.GRID),
             ("second holdout of #175", S2.GRID), ("holdout of #498", S4.GRID))
DEFECT_CALLS = (
    "K_STATS_F_Density(6;300;900)", "K_STATS_F_Density(5;500;700)",
    "K_STATS_StudentT_Density(20;900)", "K_STATS_F_Density(0.2;300;900)",
    "K_STATS_F_Density(3;300;900)", "K_STATS_StudentT_Density(16;700)",
    "K_STATS_StudentT_Density(9;80)", "K_STATS_Beta_Density(0.812111;150;90)",
    "K_STATS_StudentT_Density(20;80)", "K_STATS_Beta_Density(0.362112;480;500)")
DROPS = S4.DROPS


def design():
    """(class, function, arguments); on each side of the mode, the points where the
    logarithm of the density has fallen by each of DROPS."""
    out = []

    def add(cls, fn, *args):
        args = tuple(float(a) for a in args)
        with mp.workdps(S1.DPS):
            ref = S1.REFERENCE[fn](*[mp.mpf(a) for a in args])
        if mp.isfinite(ref) and ref > S1.SMALLEST and (cls, fn, args) not in out:
            out.append((cls, fn, args))

    def sides(log_f, centre, low, high, both=True):
        for drop in DROPS:
            for lo, hi in (((low, None), (None, high)) if both else ((None, high),)):
                x = S4._at_drop(log_f, centre, lo, hi, drop, False)
                if x is not None:
                    yield x

    for a, b in ((7.5, 19.0), (130.0, 70.0), (520.0, 460.0), (3500.0, 8000.0), (2.5, 640.0),
                 (64.0, 64.0)):
        def log_f(x, a=a, b=b):
            return (a - 1) * math.log(x) + (b - 1) * math.log1p(-x)
        for x in sides(log_f, (a - 1) / (a + b - 2), 1e-6, 1 - 1e-6):
            if 1e-6 < x < 1 - 1e-6:
                add("bd", "Beta_Density", x, a, b)
    for d1, d2 in ((8, 30), (250, 1100), (1400, 90), (46, 46), (3000, 5000)):
        def log_f(x, d1=d1, d2=d2):
            return (d1 / 2 - 1) * math.log(x) - (d1 + d2) / 2 * math.log1p(d1 * x / d2)
        centre = (d1 - 2) / d1 * d2 / (d2 + 2)
        for x in sides(log_f, centre, centre * 1e-4, centre * 1e4):
            add("fd", "F_Density", x, d1, d2)
    for df in (3, 11, 64, 450, 5000, 40000):
        def log_f(t, df=df):
            return -(df + 1) / 2 * math.log1p(t * t / df)
        for t in sides(log_f, 0.0, None, 1e12, both=False):
            add("td", "StudentT_Density", t, df)
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
    for d in (S1.design, S2.design, S4.design):
        earlier |= {(fn, args) for _, fn, args in d()}
    return [(fn, args) for _, fn, args in design() if (fn, args) in earlier]


MODEL = {"Beta_Density": D.beta_density, "F_Density": D.f_density,
         "StudentT_Density": D.student_t_density}


def model():
    by_class = {}
    for cls, fn, args in design():
        with mp.workdps(60):
            ref = S1.REFERENCE[fn](*[mp.mpf(a) for a in args])
        s = by_class.setdefault(cls, [0, 0.0, 0.0, 0, None])
        today, rule = P.rel(MODEL[fn](*args, rule=False), ref), P.rel(MODEL[fn](*args), ref)
        s[0] += 1
        s[1] = max(s[1], today)
        if rule > s[2]:
            s[2], s[4] = rule, (fn, args)
        s[3] += today > 1e-13
    for cls, (n, today, rule, above, where) in sorted(by_class.items()):
        print(f"  {cls}  {n:>3} points: today worst {today:.3g}, {above} above 1E-13; "
              f"rule worst {rule:.3g} at {where[0]}{where[1]}")
    return 0


def freeze_before():
    data = {"issue": ISSUE, "holdout": S1._slim(S1.verdicts(GRID)),
            "elsewhere": {name: S1._slim(S1.verdicts(path)) for name, path in ELSEWHERE}}
    with open(BEFORE, "w", encoding="utf-8", newline="\n") as f:
        json.dump(data, f, indent=1)
        f.write("\n")
    print(f"kept {len(data['holdout'])} holdout verdicts and those of "
          f"{len(data['elsewhere'])} other grids in {BEFORE}")
    return 0


def judge(holdout, elsewhere, before):
    """(failures, report); elsewhere maps a grid's name to its verdicts after."""
    failures, report = [], {"classes": {}, "new": []}

    def inferior(points):
        return {k for k, p in points.items() if p["verdict"] == "inferior"}

    for k, p in sorted(holdout.items()):
        c = report["classes"].setdefault(k.split(":")[0], {
            "points": 0, "worst": 0.0, "worst_before": 0.0, "inferior": 0, "inferior_before": 0})
        c["points"] += 1
        prior = before["holdout"][k]
        c["worst_before"] = max(c["worst_before"], prior["library_rel"] or 0.0)
        c["inferior_before"] += prior["verdict"] == "inferior"
        if p["verdict"] == "library error" or p.get("library_missing"):
            failures.append(f"L2 {p['kstats_call']}: the library returns an error")
            continue
        e = p.get("library_rel")
        if e is None:
            continue                                   # a count, not a relative error
        c["worst"] = max(c["worst"], e)
        if e > ACCURACY:
            failures.append(f"L1 {p['kstats_call']}: {e:.3g} above {ACCURACY:g}")
        if p["verdict"] == "inferior":
            c["inferior"] += 1
            failures.append(f"L2 {p['kstats_call']}: inferior to Excel ({e:.3g}, "
                            f"Excel {p.get('excel_rel', 0.0):.3g})")
    by_call = {}
    for name in ("first holdout of #175", "second holdout of #175"):
        by_call.update({p["kstats_call"]: p for p in elsewhere[name].values()})
    defect = []
    for call in DEFECT_CALLS:
        e = by_call[call].get("library_rel") if call in by_call else None
        defect.append(e)
        if e is None or e > ACCURACY:
            failures.append(f"L3 {call}: {e} above {ACCURACY:g}")
    report["defect_worst"] = max((e if e is not None else 1.0) for e in defect)
    for name, after in elsewhere.items():
        prior = before["elsewhere"][name]
        now, then = inferior(after), inferior(prior)
        report[name] = {"inferior_before": len(then), "inferior_after": len(now)}
        for k in sorted(now - then):
            p = after[k]
            e = p.get("library_rel", 1.0)
            report["new"].append((name, p["kstats_call"], prior[k]["library_rel"], e,
                                  p.get("excel_rel")))
            if e > LOSS:
                failures.append(f"L4 {p['kstats_call']}: newly inferior in the {name} at "
                                f"{e:.3g}, above {LOSS:g}")
        if len(now) > len(then):
            failures.append(f"L4 {len(now)} points of the {name} inferior, {len(then)} before")
    return failures, report


def check():
    with open(BEFORE, encoding="utf-8") as f:
        before = json.load(f)
    failures, report = judge(S1.verdicts(GRID),
                             {name: S1.verdicts(path) for name, path in ELSEWHERE}, before)
    for cls, c in sorted(report["classes"].items()):
        print(f"  {cls}  {c['points']:>3} points: worst {c['worst']:.3g} (before "
              f"{c['worst_before']:.3g}); inferior to Excel {c['inferior']} (before "
              f"{c['inferior_before']})")
    print(f"  the ten points of {ISSUE}, worst     {report['defect_worst']:.3g}")
    for name, _ in ELSEWHERE:
        r = report[name]
        print(f"  {name}: inferior to Excel {r['inferior_after']} (before {r['inferior_before']})")
    for name, call, was, now, excel in report["new"]:
        print(f"  newly inferior ({name}): {call}  {was:.3g} -> {now:.3g}  Excel {excel:.3g}")
    for text in failures:
        print(f"FAIL: {text}")
    if failures:
        print(f"REJECTED: {len(failures)} finding(s)")
        return 1
    print("ACCEPTED: L1 to L4 hold")
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
