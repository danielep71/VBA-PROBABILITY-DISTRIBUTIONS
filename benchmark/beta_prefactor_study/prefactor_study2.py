"""
#175, second registration (PREREGISTRATION_2.md): a fresh holdout and rules
that separate a real loss from a point crossing the floor by rounding.

The first candidate was rejected by its rule H2 at two points
(RESULT_1.md). The rule itself is unchanged; what changes is how it is
judged, and that is fixed here before any of these points is evaluated.

The grid is benchmark/excel_function_map/prefactor_holdout2_grid.csv, in the
format of the Excel floor grid, exported by Export_PrefactorHoldout2 as the
study export xfp_prefactor2. No point of it is in the first holdout.

Rules, on the exported grid (--check), against the export of the source of
before (prefactor2_before.json, written by --freeze-before):

  K1  forward     every ib point is within FORWARD of the reference;
  K2  floor       no holdout point is newly inferior to Excel with a library
                  error above LOSS; the newly inferior points are at most
                  NEW_SHARE of the holdout; and at most HALF as many points
                  are inferior as before;
  K3  numbers     every row holds an observation and no library value is ERROR;
  K4  the defect  each of the twelve points the Excel floor records against
                  #175 is within DEFECT of the reference; the floor grid has
                  no newly inferior point with a library error above LOSS, and
                  fewer inferior points than before.

Reported, not gating: every newly inferior point, the dv points before and
after.

    python prefactor_study2.py --build           write the holdout grid (blank observations)
    python prefactor_study2.py --build --check   fail if the committed grid's keys differ
    python prefactor_study2.py --freeze-before   keep the verdicts of the source of before
    python prefactor_study2.py --check           judge the exported grids; exit 0 when K1-K4 hold
"""
import argparse
import json
import math
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import prefactor_study as S1  # noqa: E402

GRID = os.path.join(S1.MAP, "prefactor_holdout2_grid.csv")
BEFORE = os.path.join(HERE, "prefactor2_before.json")
FORWARD = 2e-13          # K1: the balanced Beta guarantee
LOSS = 3e-14             # K2, K4: a newly inferior point above it is a real loss
DEFECT = 3e-14           # K4
NEW_SHARE = 0.02         # K2: share of the holdout that may cross the floor by rounding
HALF = 0.5               # K2: inferior points after, as a share of before


def design():
    """(class, function, arguments) for every point; none is in the first holdout."""
    out = []

    def add(cls, fn, *args):
        args = tuple(float(a) for a in args)
        with S1.mp.workdps(S1.DPS):
            ref = S1.REFERENCE[fn](*[S1.mp.mpf(a) for a in args])
        if S1.mp.isfinite(ref) and ref > S1.SMALLEST:
            out.append((cls, fn, args))

    for a, b in ((3.5, 11.0), (18.0, 42.0), (70.0, 70.0), (95.0, 210.0), (220.0, 130.0),
                 (410.0, 55.0), (350.0, 620.0), (880.0, 5.0), (11.0, 520.0), (1.25, 400.0),
                 (0.6, 45.0)):
        for z in (-6, -2, -0.5, 1, 3, 7):
            x = S1._beta_x(a, b, z)
            if x is not None:
                add("ib", "Beta_Cumulative", x, a, b)
                add("ib", "Beta_Survival", x, a, b)
    for s, p in ((60, 0.25), (220, 0.7), (380, 0.9), (15, 0.08)):
        mean = s * (1 - p) / p
        sd = math.sqrt(s * (1 - p)) / p
        for z in (-2.5, -0.5, 1, 3, 6):
            f = max(0, round(mean + z * sd))
            if s + f + 1 < 1000:
                add("ib", "NegativeBinomial_Cumulative", f, s, p)
                add("ib", "NegativeBinomial_Survival", f, s, p)
    for n, p in ((90, 0.35), (520, 0.5), (880, 0.06), (640, 0.8)):
        sd = math.sqrt(n * p * (1 - p))
        for z in (-5, -2, 0.3, 1.5, 4):
            k = min(n - 1, max(0, round(n * p + z * sd)))
            add("ib", "Binomial_Cumulative", k, n, p)
            add("ib", "Binomial_Survival", k, n, p)
    for d1, d2 in ((9, 60), (80, 80), (500, 700), (900, 12)):
        for x in (0.4, 0.8, 1.2, 1.8, 3.0, 6.0):
            add("ib", "F_Cumulative", x, d1, d2)
            add("ib", "F_Survival", x, d1, d2)
    for df in (9, 45, 240, 900, 1700):
        for t in (0.6, 1.3, 2.4, 4.0, 7.0, 11.0, 14.0):
            add("ib", "StudentT_Survival", t, df)

    for n, p in ((8000, 0.4), (300000, 0.02)):
        sd = math.sqrt(n * p * (1 - p))
        for z in (-20, -9, -4, 4, 9, 20):
            k = round(n * p + z * sd)
            if 0 <= k <= n:
                add("dv", "Binomial_PMF", k, n, p)
    for mean in (5.0, 120.0, 4000.0, 60000.0):
        for factor in (0.4, 0.7, 1.4, 2.5):
            add("dv", "Poisson_PMF", round(mean * factor), mean)
    for shape in (6.0, 150.0, 5000.0):
        for factor in (0.3, 0.7, 1.5, 2.5):
            add("dv", "Gamma_Density", S1._short(shape * factor), shape, 1.0)
    for a, b in ((18.0, 42.0), (220.0, 130.0), (350.0, 620.0), (6000.0, 2500.0)):
        for z in (-7, 5):
            if S1._beta_x(a, b, z) is not None:
                add("dv", "Beta_Density", S1._beta_x(a, b, z), a, b)
    for d1, d2 in ((9, 60), (500, 700)):
        for x in (0.3, 2.0, 5.0):
            add("dv", "F_Density", x, d1, d2)
    for df in (7, 60, 700):
        for t in (3.0, 8.0, 16.0):
            add("dv", "StudentT_Density", t, df)
    return out


def _with_design(function, *args):
    """Run a function of the first study on this round's design and grid."""
    saved = S1.design, S1.GRID
    S1.design, S1.GRID = design, GRID
    try:
        return function(*args)
    finally:
        S1.design, S1.GRID = saved


def build(check):
    code = _with_design(S1.build, check)
    return code


def overlap():
    """Points of this design that are also in the first holdout: must be none."""
    first = {(fn, args) for _, fn, args in S1.design()}
    return [(fn, args) for _, fn, args in design() if (fn, args) in first]


def freeze_before():
    data = {"issue": S1.ISSUE, "holdout": S1._slim(S1.verdicts(GRID)),
            "floor": S1._slim(S1.verdicts(S1.FB.GRID)), "defect_points": S1.defect_points()}
    with open(BEFORE, "w", encoding="utf-8", newline="\n") as f:
        json.dump(data, f, indent=1)
        f.write("\n")
    print(f"kept {len(data['holdout'])} holdout and {len(data['floor'])} floor verdicts, "
          f"{len(data['defect_points'])} defect points, in {BEFORE}")
    return 0


def judge(holdout, floor, before):
    """(failures, report) from the verdicts after and the frozen verdicts of before."""
    failures, report = [], {"new": []}

    def inferior(points):
        return {k for k, p in points.items() if p["verdict"] == "inferior"}

    for p in holdout.values():
        if p["verdict"] == "library error":
            failures.append(f"K3 {p['kstats_call']}: the library returns an error")
    ib = {k: p for k, p in holdout.items() if k.startswith("ib:")}
    worst = max((p.get("library_rel") or 0.0) for p in ib.values()) if ib else 0.0
    for p in ib.values():
        if (p.get("library_rel") or 0.0) > FORWARD:
            failures.append(f"K1 {p['kstats_call']}: {p['library_rel']:.3g} above {FORWARD:g}")
    report["ib_worst"] = worst
    report["ib_worst_before"] = max((v["library_rel"] or 0.0)
                                    for k, v in before["holdout"].items() if k.startswith("ib:"))
    for name, rule, after, prior in (("holdout", "K2", holdout, before["holdout"]),
                                     ("floor grid", "K4", floor, before["floor"])):
        now, then = inferior(after), inferior(prior)
        new = sorted(now - then)
        report[name] = {"inferior_before": len(then), "inferior_after": len(now),
                        "newly_inferior": len(new)}
        for k in new:
            p = after[k]
            e = p.get("library_rel", 1.0)
            report["new"].append((name, p["kstats_call"], prior[k]["library_rel"], e,
                                  p.get("excel_rel")))
            if e > LOSS:
                failures.append(f"{rule} {p['kstats_call']}: newly inferior in the {name} at "
                                f"{e:.3g}, above {LOSS:g}")
        if name == "holdout":
            if len(new) > NEW_SHARE * len(after):
                failures.append(f"K2 {len(new)} newly inferior holdout points, above "
                                f"{NEW_SHARE:.0%} of {len(after)}")
            if len(now) > HALF * len(then):
                failures.append(f"K2 {len(now)} holdout points inferior, above half of "
                                f"{len(then)} before")
        elif len(now) >= len(then):
            failures.append(f"K4 {len(now)} floor points inferior, {len(then)} before")
    defect = []
    for k in before["defect_points"]:
        e = floor[k].get("library_rel")
        defect.append(e)
        if e is None or e > DEFECT:
            failures.append(f"K4 {floor[k]['kstats_call']}: {e} above {DEFECT:g}")
    report["defect_worst"] = max((e or 1.0) for e in defect) if defect else 0.0
    dv = {k: p for k, p in holdout.items() if k.startswith("dv:")}
    prior = before["holdout"]
    report["dv"] = {
        "points": len(dv),
        "worst_before": max((prior[k]["library_rel"] or 0.0) for k in dv) if dv else 0.0,
        "worst_after": max((p.get("library_rel") or 0.0) for p in dv.values()) if dv else 0.0,
        "improved": sum(1 for k, p in dv.items()
                        if (p.get("library_rel") or 0.0) < (prior[k]["library_rel"] or 0.0)),
        "worsened": sum(1 for k, p in dv.items()
                        if (p.get("library_rel") or 0.0) > (prior[k]["library_rel"] or 0.0)),
        "worst_rise": max(((p.get("library_rel") or 0.0) - (prior[k]["library_rel"] or 0.0))
                          for k, p in dv.items()) if dv else 0.0,
    }
    return failures, report


def check():
    with open(BEFORE, encoding="utf-8") as f:
        before = json.load(f)
    failures, report = judge(S1.verdicts(GRID), S1.verdicts(S1.FB.GRID), before)
    print(f"  ib points, worst                  {report['ib_worst']:.3g} "
          f"(before {report['ib_worst_before']:.3g})")
    for name in ("holdout", "floor grid"):
        r = report[name]
        print(f"  {name}, inferior to Excel    {r['inferior_after']} (before "
              f"{r['inferior_before']}), {r['newly_inferior']} newly")
    print(f"  the points of {S1.ISSUE}, worst         {report['defect_worst']:.3g}")
    d = report["dv"]
    print(f"  dv points (reported)              worst {d['worst_after']:.3g} (before "
          f"{d['worst_before']:.3g}); {d['improved']} improved, {d['worsened']} worsened of "
          f"{d['points']}; largest rise {d['worst_rise']:.3g}")
    for name, call, was, now, excel in report["new"]:
        print(f"  newly inferior ({name}): {call}  {was:.3g} -> {now:.3g}  Excel {excel:.3g}")
    for text in failures:
        print(f"FAIL: {text}")
    if failures:
        print(f"REJECTED: {len(failures)} finding(s)")
        return 1
    print("ACCEPTED: K1 to K4 hold")
    return 0


def main(argv=None):
    ap = argparse.ArgumentParser(description=(__doc__ or "").split("\n\n")[0])
    ap.add_argument("--build", action="store_true")
    ap.add_argument("--freeze-before", action="store_true")
    ap.add_argument("--check", action="store_true")
    ap.add_argument("--model", action="store_true")
    a = ap.parse_args(argv)
    if a.model:
        return _with_design(S1.model)
    if a.build:
        return build(a.check)
    if a.freeze_before:
        return freeze_before()
    if a.check:
        return check()
    ap.error("choose --build, --freeze-before, --check or --model")
    return 2


if __name__ == "__main__":
    sys.exit(main())
