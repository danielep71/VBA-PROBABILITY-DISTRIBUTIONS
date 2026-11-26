"""
#175: the incomplete-beta prefactor below A + B = 1000, and the far deviance.

The rule and its acceptance rules are in PREREGISTRATION.md; the model of the
rule is prefactor_model.py. This module owns the holdout grid, which Excel
evaluates on both sides (the library call and every Excel formulation of the
same quantity), and judges it.

The grid is benchmark/excel_function_map/prefactor_holdout_grid.csv, in the
format of the Excel floor grid (#493), exported by Export_PrefactorHoldout as
the study export xfp_prefactor. Its points are in two classes:

  ib   quantities that reach PROB_TryBetaRegularized with A + B below 1000 and
       both shapes at least 1/2: Beta, Negative binomial, Binomial and F
       cumulative and survival, Student t survival;
  dv   quantities that reach the far branch of PROB_TryDeviancePart: Binomial
       and Poisson PMF, Gamma, Beta, F and Student t density, far from the mean.

Rules, on the exported grid (--check), against the export of the source of
before (prefactor_before.json, written by --freeze-before):

  H1  forward     every ib point is within FORWARD of the reference;
  H2  floor       no holdout point is inferior to Excel that was not inferior
                  before, and no more points are inferior than before;
  H3  numbers     every row holds an observation and no library value is ERROR;
  H4  the defect  each of the twelve points the Excel floor records against
                  #175 is within DEFECT of the reference, and the floor grid
                  has no inferior point that was not inferior before.

Reported, not gating: the dv points before and after.

    python prefactor_study.py --build           write the holdout grid (blank observations)
    python prefactor_study.py --build --check   fail if the committed grid's keys differ
    python prefactor_study.py --freeze-before   keep the verdicts of the source of before
    python prefactor_study.py --check           judge the exported grids; exit 0 when H1-H4 hold
    python prefactor_study.py --model           the rule against today on the ib points, in the model
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
MAP = os.path.join(BENCH, "excel_function_map")
sys.path.insert(0, MAP)
sys.path.insert(0, HERE)
import floor_build as FB  # noqa: E402
import floor_check as FC  # noqa: E402

GRID = os.path.join(MAP, "prefactor_holdout_grid.csv")
BEFORE = os.path.join(HERE, "prefactor_before.json")
FLOOR_DISPOSITIONS = FC.DISPOSITIONS
ISSUE = "#175"
FORWARD = 2e-13          # H1: the balanced Beta guarantee
DEFECT = 3e-14           # H4
DPS = 50
SMALLEST = mp.mpf(10) ** -290


def _lower(a, b, x):
    """I_x(a, b); an upper tail is the lower tail of the swapped shapes at 1 - x."""
    return mp.betainc(a, b, 0, x, regularized=True)


def _f_arg(x, d1, d2):
    return d1 * x / (d1 * x + d2)


def _log_beta_pdf(x, a, b):
    return (a - 1) * mp.log(x) + (b - 1) * mp.log(1 - x) - mp.log(mp.beta(a, b))


# function -> the value from its definition, at mpf arguments
REFERENCE = {
    "Beta_Cumulative": lambda x, a, b: _lower(a, b, x),
    "Beta_Survival": lambda x, a, b: _lower(b, a, 1 - x),
    "Beta_Density": lambda x, a, b: mp.exp(_log_beta_pdf(x, a, b)),
    "NegativeBinomial_Cumulative": lambda f, s, p: _lower(s, f + 1, p),
    "NegativeBinomial_Survival": lambda f, s, p: _lower(f + 1, s, 1 - p),
    "Binomial_Cumulative": lambda k, n, p: _lower(n - k, k + 1, 1 - p),
    "Binomial_Survival": lambda k, n, p: _lower(k + 1, n - k, p),
    "Binomial_PMF": lambda k, n, p: mp.binomial(n, k) * p ** k * (1 - p) ** (n - k),
    "F_Cumulative": lambda x, d1, d2: _lower(d1 / 2, d2 / 2, _f_arg(x, d1, d2)),
    "F_Survival": lambda x, d1, d2: _lower(d2 / 2, d1 / 2, d2 / (d1 * x + d2)),
    "F_Density": lambda x, d1, d2: mp.exp(
        _log_beta_pdf(_f_arg(x, d1, d2), d1 / 2, d2 / 2)
        + mp.log(d1 * d2) - 2 * mp.log(d1 * x + d2)),
    "StudentT_Survival": lambda t, df: _lower(df / 2, mp.mpf(1) / 2, df / (df + t * t)) / 2,
    "StudentT_Density": lambda t, df: mp.exp(
        mp.loggamma((df + 1) / 2) - mp.loggamma(df / 2) - mp.log(df * mp.pi) / 2
        - (df + 1) / 2 * mp.log1p(t * t / df)),
    "Poisson_PMF": lambda k, m: mp.exp(-m + k * mp.log(m) - mp.loggamma(k + 1)),
    "Gamma_Density": lambda x, shape, scale: mp.exp(
        (shape - 1) * mp.log(x / scale) - x / scale - mp.loggamma(shape)) / scale,
}


def _short(x):
    """x at six significant digits: short in the grid and exact for Excel's parser."""
    return float(f"{x:.6g}")


def _beta_x(a, b, z):
    mean = a / (a + b)
    sd = math.sqrt(a * b / ((a + b) ** 2 * (a + b + 1)))
    x = mean + z * sd
    return _short(x) if 1e-6 < x < 1 - 1e-6 else None      # None: outside the support


def design():
    """(class, function, arguments) for every holdout point, in the order of the grid."""
    out = []

    def add(cls, fn, *args):
        args = tuple(float(a) for a in args)
        with mp.workdps(DPS):
            ref = REFERENCE[fn](*[mp.mpf(a) for a in args])
        if mp.isfinite(ref) and ref > SMALLEST:       # a value a Double holds comfortably
            out.append((cls, fn, args))

    for a, b in ((2.5, 7.0), (12.0, 30.0), (45.0, 45.0), (60.0, 140.0), (150.0, 90.0),
                 (300.0, 40.0), (480.0, 500.0), (700.0, 3.0), (8.0, 640.0), (0.75, 300.0)):
        for z in (-8, -3, -1, 0.5, 2, 5):
            x = _beta_x(a, b, z)
            if x is not None:
                add("ib", "Beta_Cumulative", x, a, b)
                add("ib", "Beta_Survival", x, a, b)
    for s, p in ((40, 0.3), (150, 0.6), (450, 0.85), (25, 0.05)):
        mean = s * (1 - p) / p
        sd = math.sqrt(s * (1 - p)) / p
        for z in (-3, -1, 0.5, 2, 5):
            f = max(0, round(mean + z * sd))
            if s + f + 1 < 1000:
                add("ib", "NegativeBinomial_Cumulative", f, s, p)
                add("ib", "NegativeBinomial_Survival", f, s, p)
    for n, p in ((60, 0.2), (400, 0.5), (950, 0.03), (700, 0.9)):
        sd = math.sqrt(n * p * (1 - p))
        for z in (-4, -1.5, 0.5, 2, 5):
            k = min(n - 1, max(0, round(n * p + z * sd)))
            add("ib", "Binomial_Cumulative", k, n, p)
            add("ib", "Binomial_Survival", k, n, p)
    for d1, d2 in ((6, 40), (50, 50), (300, 900), (1200, 7)):
        for x in (0.3, 0.7, 1.0, 1.4, 2.5, 5.0):
            add("ib", "F_Cumulative", x, d1, d2)
            add("ib", "F_Survival", x, d1, d2)
    for df in (7, 61, 350, 1200, 1900):
        for t in (0.4, 1.1, 1.9, 3.2, 5.5, 9.0, 15.0):
            add("ib", "StudentT_Survival", t, df)

    for n, p in ((5000, 0.3), (200000, 0.01)):
        sd = math.sqrt(n * p * (1 - p))
        for z in (-25, -12, -6, 6, 12, 25):
            k = round(n * p + z * sd)
            if 0 <= k <= n:
                add("dv", "Binomial_PMF", k, n, p)
    for mean in (3.0, 80.0, 2500.0, 100000.0):
        for factor in (0.3, 0.6, 1.6, 3.0):
            add("dv", "Poisson_PMF", round(mean * factor), mean)
    for shape in (4.0, 90.0, 3000.0):
        for factor in (0.25, 0.6, 1.7, 3.0):
            add("dv", "Gamma_Density", _short(shape * factor), shape, 1.0)
    for a, b in ((12.0, 30.0), (150.0, 90.0), (480.0, 500.0), (4000.0, 9000.0)):
        for z in (-8, 6):
            if _beta_x(a, b, z) is not None:
                add("dv", "Beta_Density", _beta_x(a, b, z), a, b)
    for d1, d2 in ((6, 40), (300, 900)):
        for x in (0.2, 3.0, 6.0):
            add("dv", "F_Density", x, d1, d2)
    for df in (5, 80, 900):
        for t in (4.0, 9.0, 20.0):
            add("dv", "StudentT_Density", t, df)
    return out


def rows():
    """The grid rows: one per Excel formulation of each point."""
    out = []
    with mp.workdps(DPS):
        for index, (cls, fn, args) in enumerate(design(), 1):
            arity, truncated, forms = FB.EXCEL[fn]
            if len(args) != arity or any(args[i] != math.floor(args[i]) for i in truncated):
                raise SystemExit(f"{fn}{args}: not a fair comparison with Excel")
            tokens = [FB.token(repr(a)) for a in args]
            if any(t is None for t in tokens):
                raise SystemExit(f"{fn}{args}: an argument needs more than 15 digits")
            ref = REFERENCE[fn](*[mp.mpf(a) for a in args])
            kind = f"{cls}:P{index:03d}:{fn}"
            call = f"K_STATS_{fn}(" + ";".join(tokens) + ")"
            named = {f"a{i}": t for i, t in enumerate(tokens)}
            for j, form in enumerate(forms):
                out.append([f"p{index:03d}{chr(97 + j)}", kind, form.format(**named), call,
                            mp.nstr(ref, 34), "", ""])
    return out


def render():
    buf = io.StringIO()
    w = csv.writer(buf, lineterminator="\n")
    w.writerow(FB.FIELDS)
    for r in rows():
        if any("," in field for field in r[:4]):
            raise SystemExit(f"a field holds a comma: {r}")
        w.writerow(r)
    return buf.getvalue()


def build(check):
    text = render()
    if check:
        with open(GRID, encoding="utf-8") as f:
            committed = [line.split(",")[:5] for line in f.read().splitlines()]
        if committed != [line.split(",")[:5] for line in text.splitlines()]:
            print(f"FAIL: {os.path.basename(GRID)} differs from its generator; rebuild it")
            return 1
        print(f"PASS: {os.path.basename(GRID)} matches its generator ({len(design())} points)")
        return 0
    with open(GRID, "w", encoding="utf-8", newline="") as f:
        f.write(text)
    print(f"wrote {GRID}: {len(design())} points, {len(text.splitlines()) - 1} rows")
    return 0


def read(path):
    with open(path, encoding="utf-8-sig") as f:
        return list(csv.DictReader(f))


def verdicts(path):
    """{kind: point} from an exported grid in the floor format; exits on a blank row."""
    points, blank = FC.judge(read(path))
    if blank:
        raise SystemExit(f"FAIL: {blank} row(s) of {os.path.basename(path)} hold no observation: "
                         "export the grid first")
    return {p["kind"]: p for p in points}


def defect_points():
    """The floor-grid points recorded against this issue."""
    _, recorded = FC.load_dispositions(FLOOR_DISPOSITIONS)
    return sorted(k for k, e in recorded.items() if e["issue"] == ISSUE)


def _slim(points):
    return {k: {"verdict": p["verdict"], "library_rel": p.get("library_rel"),
                "excel_rel": p.get("excel_rel")} for k, p in sorted(points.items())}


def freeze_before():
    """Keep the verdicts of the source of before: the holdout and the whole floor grid."""
    data = {"issue": ISSUE, "holdout": _slim(verdicts(GRID)), "floor": _slim(verdicts(FB.GRID)),
            "defect_points": defect_points()}
    with open(BEFORE, "w", encoding="utf-8", newline="\n") as f:
        json.dump(data, f, indent=1)
        f.write("\n")
    print(f"kept {len(data['holdout'])} holdout and {len(data['floor'])} floor verdicts, "
          f"{len(data['defect_points'])} defect points, in {BEFORE}")
    return 0


def judge(holdout, floor, before):
    """(failures, report) from the verdicts after and the frozen verdicts of before."""
    failures, report = [], {}

    def inferior(points):
        return {k for k, p in points.items() if p["verdict"] == "inferior"}

    ib = {k: p for k, p in holdout.items() if k.startswith("ib:")}
    dv = {k: p for k, p in holdout.items() if k.startswith("dv:")}
    for k, p in holdout.items():
        if p["verdict"] == "library error":
            failures.append(f"H3 {p['kstats_call']}: the library returns an error")
    worst = 0.0
    for k, p in ib.items():
        e = p.get("library_rel")
        if e is None:
            continue
        worst = max(worst, e)
        if e > FORWARD:
            failures.append(f"H1 {p['kstats_call']}: {e:.3g} above {FORWARD:g}")
    report["ib_worst"] = worst
    report["ib_worst_before"] = max((v["library_rel"] or 0.0) for k, v in before["holdout"].items()
                                    if k.startswith("ib:"))
    for name, after, prior in (("holdout", holdout, before["holdout"]),
                               ("floor grid", floor, before["floor"])):
        now, then = inferior(after), inferior(prior)
        report[name] = {"inferior_before": len(then), "inferior_after": len(now)}
        rule = "H2" if name == "holdout" else "H4"
        for k in sorted(now - then):
            p = after[k]
            failures.append(f"{rule} {p['kstats_call']}: newly inferior to Excel in the {name} "
                            f"(library {p.get('library_rel', 1.0):.3g}, "
                            f"Excel {p.get('excel_rel', 0.0):.3g})")
        if name == "holdout" and len(now) > len(then):
            failures.append(f"H2 {len(now)} holdout points inferior, {len(then)} before")
    defect = []
    for k in before["defect_points"]:
        e = floor[k].get("library_rel")
        defect.append((k, e))
        if e is None or e > DEFECT:
            failures.append(f"H4 {floor[k]['kstats_call']}: {e} above {DEFECT:g}")
    report["defect_worst"] = max((e or 1.0) for _, e in defect) if defect else 0.0
    report["dv"] = {
        "points": len(dv),
        "worst_before": max((v["library_rel"] or 0.0) for k, v in before["holdout"].items()
                            if k.startswith("dv:")),
        "worst_after": max((p.get("library_rel") or 0.0) for p in dv.values()),
        "improved": sum(1 for k, p in dv.items()
                        if (p.get("library_rel") or 0.0)
                        < (before["holdout"][k]["library_rel"] or 0.0)),
        "worsened": sum(1 for k, p in dv.items()
                        if (p.get("library_rel") or 0.0)
                        > (before["holdout"][k]["library_rel"] or 0.0)),
    }
    return failures, report


def check():
    with open(BEFORE, encoding="utf-8") as f:
        before = json.load(f)
    failures, report = judge(verdicts(GRID), verdicts(FB.GRID), before)
    print(f"  ib points, worst                  {report['ib_worst']:.3g} "
          f"(before {report['ib_worst_before']:.3g})")
    for name in ("holdout", "floor grid"):
        r = report[name]
        print(f"  {name}, inferior to Excel    {r['inferior_after']} (before {r['inferior_before']})")
    print(f"  the points of {ISSUE}, worst         {report['defect_worst']:.3g}")
    d = report["dv"]
    print(f"  dv points (reported)              worst {d['worst_after']:.3g} "
          f"(before {d['worst_before']:.3g}); {d['improved']} improved, {d['worsened']} worsened "
          f"of {d['points']}")
    for text in failures:
        print(f"FAIL: {text}")
    if failures:
        print(f"REJECTED: {len(failures)} finding(s)")
        return 1
    print("ACCEPTED: H1 to H4 hold")
    return 0


def beta_call(fn, args):
    """(x, y, a, b) of the incomplete beta an ib point reduces to, in Doubles."""
    if fn == "Beta_Cumulative":
        x, a, b = args
        return x, 1.0 - x, a, b
    if fn == "Beta_Survival":
        x, a, b = args
        return 1.0 - x, x, b, a
    if fn == "NegativeBinomial_Cumulative":
        f, s, p = args
        return p, 1.0 - p, s, f + 1.0
    if fn == "NegativeBinomial_Survival":
        f, s, p = args
        return 1.0 - p, p, f + 1.0, s
    if fn == "Binomial_Cumulative":
        k, n, p = args
        return 1.0 - p, p, n - k, k + 1.0
    if fn == "Binomial_Survival":
        k, n, p = args
        return p, 1.0 - p, k + 1.0, n - k
    if fn in ("F_Cumulative", "F_Survival"):
        x, d1, d2 = args
        lower = d1 * x / (d1 * x + d2)
        upper = d2 / (d1 * x + d2)
        return (lower, upper, d1 / 2, d2 / 2) if fn == "F_Cumulative" else \
            (upper, lower, d2 / 2, d1 / 2)
    if fn == "StudentT_Survival":
        t, df = args                        # the logistic pair of PROB_TF_TrySquareRatioPair
        ratio = math.exp(2.0 * math.log(t) - math.log(df))
        return 1.0 / (1.0 + ratio), ratio / (1.0 + ratio), df / 2, 0.5
    raise KeyError(fn)


def model():
    """The rule against the source of today on the ib points, in both ports of the model."""
    import prefactor_model as M
    for semantics in ("x87", "double"):
        current, candidate, refused = [], [], 0
        with mp.workdps(DPS):
            for cls, fn, args in design():
                if cls != "ib":
                    continue
                x, y, a, b = beta_call(fn, args)
                ref = REFERENCE[fn](*[mp.mpf(v) for v in args])
                scale = 2 if fn == "StudentT_Survival" else 1
                try:
                    c = M.current(x, y, a, b, semantics)
                    k = M.candidate(x, y, a, b, semantics)
                except M.NotModelled:
                    refused += 1
                    continue
                current.append(float(abs((mp.mpf(c) / scale - ref) / ref)))
                candidate.append((float(abs((mp.mpf(k) / scale - ref) / ref)), fn, args))
        worst = max(candidate)
        print(f"  {semantics:6s} {len(current)} points, {refused} not modelled: today worst "
              f"{max(current):.3g}, {sum(e > 1e-13 for e in current)} above 1E-13; rule worst "
              f"{worst[0]:.3g} at {worst[1]}{worst[2]}, "
              f"{sum(e > 3e-14 for e, _, _ in candidate)} above 3E-14")
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
        return build(a.check)
    if a.freeze_before:
        return freeze_before()
    if a.check:
        return check()
    ap.error("choose --build, --freeze-before or --check")
    return 2


if __name__ == "__main__":
    sys.exit(main())
