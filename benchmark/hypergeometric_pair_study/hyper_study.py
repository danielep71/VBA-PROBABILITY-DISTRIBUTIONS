"""
#496 (PREREGISTRATION.md): the holdout of the Hypergeometric mass function
carried as pairs of Doubles, and the judge of its rules.

The grid is benchmark/excel_function_map/hypergeometric_pair_holdout_grid.csv,
in the format of the Excel floor grid, exported by
Export_HypergeometricPairHoldout as the study export xhg_pair. No point of it
is in an earlier holdout.

Rules, on the exported grids (--check), against the export of the source of
before (hyper_before.json, written by --freeze-before):

  M1  accuracy    every hg point is within ACCURACY of the reference;
  M2  floor       no hg point is inferior to Excel, every row holds an
                  observation and no library value is ERROR;
  M3  the defect  the point of the issue, a row of the floor grid, is within
                  ACCURACY;
  M4  elsewhere   in the floor grid of #493, in the two holdouts of #175 and
                  in the holdouts of #498 and #501 no point is newly inferior
                  to Excel with a library error above LOSS, and none has more
                  inferior points than before.

Reported, not gating: the nb points (NegativeBinomial_PMF, which the rule
does not change), before and after.

    python hyper_study.py --build           write the holdout grid (blank observations)
    python hyper_study.py --build --check   fail if the committed grid's keys differ
    python hyper_study.py --model           the rule against today on the hg points, in the model
    python hyper_study.py --freeze-before   keep the verdicts of the source of before
    python hyper_study.py --check           judge the exported grids; exit 0 when M1-M4 hold
"""
import argparse
import json
import math
import os
import sys

import mpmath as mp

HERE = os.path.dirname(os.path.abspath(__file__))
BENCH = os.path.dirname(HERE)
sys.path.insert(0, os.path.join(BENCH, "density_pair_study"))
sys.path.insert(0, HERE)
import density_study as S5  # noqa: E402
import hyper_model as H  # noqa: E402

S4 = S5.S4

S1, S2, P = S4.S1, S4.S2, S4.P
GRID = os.path.join(S1.MAP, "hypergeometric_pair_holdout_grid.csv")
BEFORE = os.path.join(HERE, "hyper_before.json")
ISSUE = "#496"
ACCURACY = 1e-14         # M1, M3: the floor of CONTRIBUTING.md
LOSS = 3e-14             # M4: a newly inferior point above it is a real loss
ELSEWHERE = (("floor grid", S1.FB.GRID), ("first holdout of #175", S1.GRID),
             ("second holdout of #175", S2.GRID), ("holdout of #498", S4.GRID),
             ("holdout of #501", S5.GRID))
DEFECT_CALLS = ("K_STATS_Hypergeometric_PMF(100;200;2000;50000)",)
DROPS = S4.DROPS

S1.REFERENCE.setdefault(
    "Hypergeometric_PMF",
    lambda k, n, s, pop: mp.binomial(s, k) * mp.binomial(pop - s, n - k) / mp.binomial(pop, n))
S1.REFERENCE.setdefault(
    "NegativeBinomial_PMF",
    lambda f, s, p: mp.binomial(f + s - 1, f) * p ** s * (1 - p) ** f)


def design():
    """(class, function, arguments); on each side of the mode, the points where the
    logarithm of the mass has fallen by each of DROPS."""
    out = []

    def add(cls, fn, *args):
        args = tuple(float(a) for a in args)
        with mp.workdps(S1.DPS):
            ref = S1.REFERENCE[fn](*[mp.mpf(a) for a in args])
        if mp.isfinite(ref) and ref > S1.SMALLEST and (cls, fn, args) not in out:
            out.append((cls, fn, args))

    def sides(log_f, centre, low, high):
        for drop in DROPS:
            for lo, hi in ((low, None), (None, high)):
                x = S4._at_drop(log_f, centre, lo, hi, drop, True)
                if x is not None:
                    yield x

    def lchoose(n, k):
        return math.lgamma(n + 1) - math.lgamma(k + 1) - math.lgamma(n - k + 1)

    for n, s, pop in ((100, 50, 1000), (300, 1500, 40000), (5000, 30000, 100000),
                      (300000, 400000, 1000000), (150, 120, 400), (20000, 700, 90000),
                      (3333, 12345, 77777), (70001, 250000, 999983), (613, 4871, 9973)):
        def log_f(k, n=n, s=s, pop=pop):
            return lchoose(s, k) + lchoose(pop - s, n - k)
        low, high = max(0, n - (pop - s)) + 1.0, min(n, s) - 1.0
        for k in sides(log_f, float(round(n * s / pop)), low, high):
            if H.interior(k, n, s, pop):
                add("hg", "Hypergeometric_PMF", k, n, s, pop)
    for s, p in ((40, 0.3), (2500, 0.6), (150000, 0.9)):
        def log_f(f, s=s, p=p):
            return lchoose(f + s - 1, f) + f * math.log1p(-p)
        centre = float(round((s - 1) * (1 - p) / p))
        for f in sides(log_f, centre, 0.0, centre * 40 + 4000.0):
            add("nb", "NegativeBinomial_PMF", f, s, p)
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
    for d in (S1.design, S2.design, S4.design, S5.design):
        earlier |= {(fn, args) for _, fn, args in d()}
    return [(fn, args) for _, fn, args in design() if (fn, args) in earlier]


def model():
    worst_today = worst_rule = 0.0
    above = points = 0
    where = None
    for cls, fn, args in design():
        if cls != "hg":
            continue
        with mp.workdps(60):
            ref = S1.REFERENCE[fn](*[mp.mpf(a) for a in args])
        today = P.rel(H.hypergeometric_pmf(*args, rule=False), ref)
        rule = P.rel(H.hypergeometric_pmf(*args), ref)
        points += 1
        worst_today = max(worst_today, today)
        above += today > 1e-13
        if rule > worst_rule:
            worst_rule, where = rule, args
    print(f"  hg  {points:>3} points: today worst {worst_today:.3g}, {above} above 1E-13; "
          f"rule worst {worst_rule:.3g} at Hypergeometric_PMF{where}")
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
        if p["verdict"] == "library error":
            failures.append(f"M2 {p['kstats_call']}: the library returns an error")
            continue
        e = p.get("library_rel")
        if e is None:
            continue                                   # not compared: Excel returned no number
        c["worst"] = max(c["worst"], e)
        if p["verdict"] == "inferior":
            c["inferior"] += 1
        if not k.startswith("hg:"):
            continue                                   # nb: reported, not gating
        if e > ACCURACY:
            failures.append(f"M1 {p['kstats_call']}: {e:.3g} above {ACCURACY:g}")
        if p["verdict"] == "inferior":
            failures.append(f"M2 {p['kstats_call']}: inferior to Excel ({e:.3g}, "
                            f"Excel {p.get('excel_rel', 0.0):.3g})")
    by_call = {p["kstats_call"]: p for p in elsewhere["floor grid"].values()}
    defect = []
    for call in DEFECT_CALLS:
        e = by_call[call].get("library_rel") if call in by_call else None
        defect.append(e)
        if e is None or e > ACCURACY:
            failures.append(f"M3 {call}: {e} above {ACCURACY:g}")
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
                failures.append(f"M4 {p['kstats_call']}: newly inferior in the {name} at "
                                f"{e:.3g}, above {LOSS:g}")
        if len(now) > len(then):
            failures.append(f"M4 {len(now)} points of the {name} inferior, {len(then)} before")
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
    print(f"  the point of {ISSUE}           {report['defect_worst']:.3g}")
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
    print("ACCEPTED: M1 to M4 hold")
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
