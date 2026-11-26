"""
#498 (PREREGISTRATION.md): the holdout of the deviance carried as a pair of
Doubles, and the judge of its rules.

The grid is benchmark/excel_function_map/deviance_pair_holdout_grid.csv, in
the format of the Excel floor grid, exported by Export_DeviancePairHoldout as
the study export xdp_pair. No point of it is in the two holdouts of #175.

Rules, on the exported grids (--check), against the export of the source of
before (pair_before.json, written by --freeze-before):

  J1  accuracy    every holdout point is within ACCURACY of the reference;
  J2  floor       no holdout point is inferior to Excel, every row holds an
                  observation and no library value is ERROR;
  J3  the defect  each of the five points of the issue, which are rows of the
                  second holdout of #175, is within ACCURACY;
  J4  elsewhere   in the floor grid of #493 and in the two holdouts of #175 no
                  point is newly inferior to Excel with a library error above
                  LOSS, and none has more inferior points than before.

Reported, not gating: the holdout before and after, by class.

    python pair_study.py --build           write the holdout grid (blank observations)
    python pair_study.py --build --check   fail if the committed grid's keys differ
    python pair_study.py --model           the rule against today on the holdout, in the model
    python pair_study.py --freeze-before   keep the verdicts of the source of before
    python pair_study.py --check           judge the exported grids; exit 0 when J1-J4 hold
"""
import argparse
import json
import math
import os
import sys

import mpmath as mp

HERE = os.path.dirname(os.path.abspath(__file__))
BENCH = os.path.dirname(HERE)
sys.path.insert(0, os.path.join(BENCH, "beta_prefactor_study"))
sys.path.insert(0, HERE)
import pair_model as P  # noqa: E402
import prefactor_study as S1  # noqa: E402
import prefactor_study2 as S2  # noqa: E402

GRID = os.path.join(S1.MAP, "deviance_pair_holdout_grid.csv")
BEFORE = os.path.join(HERE, "pair_before.json")
ISSUE = "#498"
ACCURACY = 1e-14         # J1, J3: the floor of CONTRIBUTING.md
LOSS = 3e-14             # J4: a newly inferior point above it is a real loss
ELSEWHERE = (("floor grid", S1.FB.GRID), ("first holdout of #175", S1.GRID),
             ("second holdout of #175", S2.GRID))
DEFECT_CALLS = ("K_STATS_Poisson_PMF(5600;4000)", "K_STATS_Gamma_Density(7500;5000;1)",
                "K_STATS_Gamma_Density(3500;5000;1)", "K_STATS_Binomial_PMF(4076;8000;0.4)",
                "K_STATS_Binomial_PMF(4466;300000;0.02)")

S1.REFERENCE.setdefault(
    "ChiSquare_Density",
    lambda x, df: mp.exp((df / 2 - 1) * mp.log(x / 2) - x / 2 - mp.loggamma(df / 2)) / 2)


DROPS = (0.5, 20.0, 150.0, 400.0, 600.0)     # Log(f(centre) / f(x)) of the points, each side


def _at_drop(log_f, centre, low, high, drop, integer):
    """x between centre and one end where Log f has fallen by `drop`; None if it never does."""
    top = log_f(centre)
    end = low if low is not None else high
    if top - log_f(end) < drop:
        return None
    a, b = centre, end
    for _ in range(200):
        mid = (a + b) / 2
        if top - log_f(mid) < drop:
            a = mid
        else:
            b = mid
    x = round(b) if integer else S1._short(b)
    return x if x != centre else None


def design():
    """(class, function, arguments) for every point; none is in a holdout of #175.

    On each side of the centre, the points where the logarithm of the function has
    fallen by each of DROPS: the deviance the rule carries is about that large.
    """
    out = []

    def add(cls, fn, *args):
        args = tuple(float(a) for a in args)
        with mp.workdps(S1.DPS):
            ref = S1.REFERENCE[fn](*[mp.mpf(a) for a in args])
        if mp.isfinite(ref) and ref > S1.SMALLEST and (cls, fn, args) not in out:
            out.append((cls, fn, args))

    def sides(log_f, centre, low, high, integer):
        for drop in DROPS:
            for lo, hi in ((low, None), (None, high)):
                x = _at_drop(log_f, centre, lo, hi, drop, integer)
                if x is not None:
                    yield x

    for mean in (7.0, 41.5, 300.0, 1234.5, 9000.0, 250000.0):
        def log_f(k, mean=mean):
            return k * math.log(mean) - mean - math.lgamma(k + 1)
        for k in sides(log_f, float(round(mean)), 1.0, 40.0 * mean + 400.0, True):
            add("pm", "Poisson_PMF", k, mean)
    for n, p in ((600, 0.3), (20000, 0.45), (150000, 0.05), (900000, 0.6), (5000, 0.93)):
        def log_f(k, n=n, p=p):
            return (math.lgamma(n + 1) - math.lgamma(k + 1) - math.lgamma(n - k + 1)
                    + k * math.log(p) + (n - k) * math.log1p(-p))
        for k in sides(log_f, float(round(n * p)), 1.0, n - 1.0, True):
            add("bn", "Binomial_PMF", k, n, p)
    for shape, scale in ((9.0, 1.0), (400.0, 1.0), (12000.0, 1.0), (400.0, 3.0), (2500.0, 0.7),
                         (62.5, 0.01)):
        def log_f(x, shape=shape, scale=scale):
            return (shape - 1) * math.log(x / scale) - x / scale
        centre = (shape - 1) * scale
        for x in sides(log_f, centre, centre * 1e-4, centre * 40 + 400 * scale, False):
            add("gd", "Gamma_Density", x, shape, scale)
    for df in (14, 900, 30000):
        def log_f(x, df=df):
            return (df / 2 - 1) * math.log(x / 2) - x / 2
        centre = float(df - 2)
        for x in sides(log_f, centre, centre * 1e-4, centre * 40 + 800, False):
            add("cs", "ChiSquare_Density", x, df)
    return out


def _with_design(function, *args):
    saved = S1.design, S1.GRID
    S1.design, S1.GRID = design, GRID
    try:
        return function(*args)
    finally:
        S1.design, S1.GRID = saved


def overlap():
    """Points of this design that are in a holdout of #175: must be none."""
    earlier = {(fn, args) for _, fn, args in S1.design()} | {(fn, args) for _, fn, args in S2.design()}
    return [(fn, args) for _, fn, args in design() if (fn, args) in earlier]


MODEL = {"Poisson_PMF": ("poisson", P.poisson_pmf),
         "Binomial_PMF": ("binomial", P.binomial_pmf),
         "Gamma_Density": ("gamma", P.gamma_density),
         "ChiSquare_Density": ("chisquare", P.chisquare_density)}


def model():
    """The rule and today on the holdout points, in the model."""
    by_class = {}
    for cls, fn, args in design():
        with mp.workdps(60):
            ref = S1.REFERENCE[fn](*[mp.mpf(a) for a in args])
        f = MODEL[fn][1]
        s = by_class.setdefault(cls, [0, 0.0, 0.0, 0, None])
        today, rule = P.rel(f(*args, rule=False), ref), P.rel(f(*args), ref)
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
        cls = k.split(":")[0]
        c = report["classes"].setdefault(cls, {"points": 0, "worst": 0.0, "worst_before": 0.0,
                                               "inferior": 0, "inferior_before": 0})
        c["points"] += 1
        prior = before["holdout"][k]
        c["worst_before"] = max(c["worst_before"], prior["library_rel"] or 0.0)
        c["inferior_before"] += prior["verdict"] == "inferior"
        if p["verdict"] == "library error" or p.get("library_missing"):
            failures.append(f"J2 {p['kstats_call']}: the library returns an error")
            continue
        e = p.get("library_rel")
        if e is None:
            continue                                   # a count, not a relative error
        c["worst"] = max(c["worst"], e)
        if e > ACCURACY:
            failures.append(f"J1 {p['kstats_call']}: {e:.3g} above {ACCURACY:g}")
        if p["verdict"] == "inferior":
            c["inferior"] += 1
            failures.append(f"J2 {p['kstats_call']}: inferior to Excel ({e:.3g}, "
                            f"Excel {p.get('excel_rel', 0.0):.3g})")
    second = elsewhere["second holdout of #175"]
    by_call = {p["kstats_call"]: p for p in second.values()}
    defect = []
    for call in DEFECT_CALLS:
        e = by_call[call].get("library_rel") if call in by_call else None
        defect.append(e)
        if e is None or e > ACCURACY:
            failures.append(f"J3 {call}: {e} above {ACCURACY:g}")
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
                failures.append(f"J4 {p['kstats_call']}: newly inferior in the {name} at "
                                f"{e:.3g}, above {LOSS:g}")
        if len(now) > len(then):
            failures.append(f"J4 {len(now)} points of the {name} inferior, {len(then)} before")
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
    print(f"  the five points of {ISSUE}, worst    {report['defect_worst']:.3g}")
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
    print("ACCEPTED: J1 to J4 hold")
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
