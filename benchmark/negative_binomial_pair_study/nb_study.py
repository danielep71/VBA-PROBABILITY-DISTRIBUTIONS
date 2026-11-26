"""
#504 (PREREGISTRATION.md): the holdout of the Negative Binomial and Geometric
mass functions carried as pairs of Doubles, and the judge of its rules.

The grid is benchmark/excel_function_map/negative_binomial_pair_holdout_grid.csv,
in the format of the Excel floor grid, exported by
Export_NegativeBinomialPairHoldout as the study export xnb_pair. No point of
it is in an earlier holdout.

Rules, on the exported grids (--check), against the export of the source of
before (nb_before.json, written by --freeze-before):

  N1  accuracy    every judged point is within ACCURACY of the reference;
  N2  floor       no judged point is inferior to Excel, every row holds an
                  observation and no library value is ERROR;
  N3  the defect  each of the eight points of the issue, rows of the holdout
                  of #496, is within ACCURACY;
  N4  elsewhere   in the floor grid of #493 and in the five earlier holdouts
                  no point is newly inferior to Excel with a library error
                  above LOSS, and none has more inferior points than before;
  N5  Geometric   where the Geometric rule does not apply (see below), no ge
                  observation changes.

The nb points are always judged. The ge points are judged only if the
Geometric rule applies, and it applies if and only if at least one ge point
is inferior to Excel in the export of before: geometric_rule(before).

    python nb_study.py --build           write the holdout grid (blank observations)
    python nb_study.py --build --check   fail if the committed grid's keys differ
    python nb_study.py --model           the rule against today on the holdout, in the model
    python nb_study.py --freeze-before   keep the verdicts of the source of before
    python nb_study.py --check           judge the exported grids; exit 0 when N1-N5 hold
"""
import argparse
import json
import math
import os
import sys

import mpmath as mp

HERE = os.path.dirname(os.path.abspath(__file__))
BENCH = os.path.dirname(HERE)
sys.path.insert(0, os.path.join(BENCH, "hypergeometric_pair_study"))
sys.path.insert(0, HERE)
import hyper_study as S6  # noqa: E402
import nb_model as NB  # noqa: E402

S5, S4 = S6.S5, S6.S4

S1, S2, P = S4.S1, S4.S2, S4.P
GRID = os.path.join(S1.MAP, "negative_binomial_pair_holdout_grid.csv")
BEFORE = os.path.join(HERE, "nb_before.json")
ISSUE = "#504"
ACCURACY = 1e-14         # N1, N3: the floor of CONTRIBUTING.md
LOSS = 3e-14             # N4: a newly inferior point above it is a real loss
ELSEWHERE = (("floor grid", S1.FB.GRID), ("first holdout of #175", S1.GRID),
             ("second holdout of #175", S2.GRID), ("holdout of #498", S4.GRID),
             ("holdout of #501", S5.GRID), ("holdout of #496", S6.GRID))
DEFECT_CALLS = (
    "K_STATS_NegativeBinomial_PMF(21622;150000;0.9)",
    "K_STATS_NegativeBinomial_PMF(12200;150000;0.9)",
    "K_STATS_NegativeBinomial_PMF(481;2500;0.6)",
    "K_STATS_NegativeBinomial_PMF(2697;2500;0.6)",
    "K_STATS_NegativeBinomial_PMF(15814;150000;0.9)",
    "K_STATS_NegativeBinomial_PMF(2096;40;0.3)",
    "K_STATS_NegativeBinomial_PMF(719;40;0.3)",
    "K_STATS_NegativeBinomial_PMF(1348;2500;0.6)")
DROPS = S4.DROPS

S1.REFERENCE.setdefault("Geometric_PMF", lambda k, p: p * (1 - p) ** k)


def geometric_rule(before):
    """True when the Geometric rule applies: a ge point inferior to Excel before."""
    return any(v["verdict"] == "inferior" for k, v in before["holdout"].items()
               if k.startswith("ge:"))


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
                if x is not None and x >= 1:
                    yield x

    def lchoose(n, k):
        return math.lgamma(n + 1) - math.lgamma(k + 1) - math.lgamma(n - k + 1)

    for s, p in ((3, 0.02), (25, 0.15), (700, 0.45), (9000, 0.2), (60000, 0.8), (900000, 0.97)):
        def log_f(f, s=s, p=p):
            return lchoose(f + s - 1, f) + f * math.log1p(-p)
        centre = float(max(1, round((s - 1) * (1 - p) / p)))
        for f in sides(log_f, centre, 1.0, centre * 60 + 60000.0):
            add("nb", "NegativeBinomial_PMF", f, s, p)
    for p in (0.00001, 0.002, 0.08, 0.4, 0.93):
        for drop in DROPS:
            k = round(drop / -math.log1p(-p))
            if k >= 1:
                add("ge", "Geometric_PMF", k, p)
                add("ge", "Geometric_PMF", k + 1, p)
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
    for d in (S1.design, S2.design, S4.design, S5.design, S6.design):
        earlier |= {(fn, args) for _, fn, args in d()}
    return [(fn, args) for _, fn, args in design() if (fn, args) in earlier]


MODEL = {"NegativeBinomial_PMF": NB.negative_binomial_pmf, "Geometric_PMF": NB.geometric_pmf}


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
    geometric = report["geometric_rule"] = geometric_rule(before)

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
            failures.append(f"N2 {p['kstats_call']}: the library returns an error")
            continue
        e = p.get("library_rel")
        if e is None:
            continue                                   # a count, not a relative error
        c["worst"] = max(c["worst"], e)
        if p["verdict"] == "inferior":
            c["inferior"] += 1
        if k.startswith("ge:") and not geometric:
            if p.get("library_rel") != prior["library_rel"]:
                failures.append(f"N5 {p['kstats_call']}: the observation changed, and the "
                                "Geometric rule does not apply")
            continue
        if e > ACCURACY:
            failures.append(f"N1 {p['kstats_call']}: {e:.3g} above {ACCURACY:g}")
        if p["verdict"] == "inferior":
            failures.append(f"N2 {p['kstats_call']}: inferior to Excel ({e:.3g}, "
                            f"Excel {p.get('excel_rel', 0.0):.3g})")
    by_call = {p["kstats_call"]: p for p in elsewhere["holdout of #496"].values()}
    defect = []
    for call in DEFECT_CALLS:
        e = by_call[call].get("library_rel") if call in by_call else None
        defect.append(e)
        if e is None or e > ACCURACY:
            failures.append(f"N3 {call}: {e} above {ACCURACY:g}")
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
                failures.append(f"N4 {p['kstats_call']}: newly inferior in the {name} at "
                                f"{e:.3g}, above {LOSS:g}")
        if len(now) > len(then):
            failures.append(f"N4 {len(now)} points of the {name} inferior, {len(then)} before")
    return failures, report


def check():
    with open(BEFORE, encoding="utf-8") as f:
        before = json.load(f)
    failures, report = judge(S1.verdicts(GRID),
                             {name: S1.verdicts(path) for name, path in ELSEWHERE}, before)
    print("  the Geometric rule " + ("applies: a ge point was inferior to Excel before"
                                  if report["geometric_rule"] else
                                  "does not apply: no ge point was inferior to Excel before"))
    for cls, c in sorted(report["classes"].items()):
        print(f"  {cls}  {c['points']:>3} points: worst {c['worst']:.3g} (before "
              f"{c['worst_before']:.3g}); inferior to Excel {c['inferior']} (before "
              f"{c['inferior_before']})")
    print(f"  the eight points of {ISSUE}, worst  {report['defect_worst']:.3g}")
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
    print("ACCEPTED: N1 to N5 hold")
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
