"""
#493: the library against Excel on every comparable point of the certified grids.

The release rule (CONTRIBUTING.md, "Excel as the floor") is that the library is
nowhere inferior to Excel: at no point is Excel's relative error smaller than
the library's while the library's is above 1E-14. This module writes the grid
Excel evaluates for that comparison; floor_check.py judges the export.

Every row of the main grid, the consolidated holdout and the older-families
holdout is a candidate. A row becomes a point when

  - its function has an Excel counterpart in EXCEL below (functions without
    one are out of scope, as #493 says);
  - every argument Excel truncates (degrees of freedom, counts) is an integer,
    so both sides see the same arguments;
  - every argument is written exactly by 15 significant digits, because
    Excel's formula parser keeps 15: a longer literal would change the
    argument for both sides and leave the reference behind;
  - the reference is a finite number, not zero.

At each point the grid carries the library call and every Excel formulation of
the same quantity. Where Excel has several (a right-tail function and 1 - CDF,
or a symmetry), all are evaluated and the floor takes the most accurate: the
comparison never understates Excel.

    python floor_build.py           write excel_floor_grid.csv
    python floor_build.py --check   fail if the committed grid differs
"""
import argparse
import csv
import io
import math
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
BENCH = os.path.dirname(HERE)
GRID = os.path.join(HERE, "excel_floor_grid.csv")
SOURCES = (
    ("main", "m", os.path.join(BENCH, "probability_accuracy_grid.csv")),
    ("holdout", "h", os.path.join(BENCH, "holdout", "holdout_grid.csv")),
    ("holdout_older", "o", os.path.join(BENCH, "holdout_older", "holdout_older_grid.csv")),
)
FIELDS = ["point_id", "kind", "excel_formula", "kstats_call", "reference",
          "observed_excel", "observed_kstats"]

# function -> (number of arguments, indices Excel truncates, Excel formulations).
# A formulation is a format string over the arguments a0..a3.
EXCEL = {
    "NormalStandard_Density": (1, (), ["NORM.S.DIST({a0};FALSE)"]),
    "NormalStandard_Cumulative": (1, (), ["NORM.S.DIST({a0};TRUE)"]),
    "NormalStandard_Survival": (1, (), ["1-NORM.S.DIST({a0};TRUE)", "NORM.S.DIST(-{a0};TRUE)"]),
    "NormalStandard_InverseCumulative": (1, (), ["NORM.S.INV({a0})"]),
    "NormalStandard_InverseCumulativeFast": (1, (), ["NORM.S.INV({a0})"]),
    "NormalStandard_InverseSurvival": (1, (), ["NORM.S.INV(1-{a0})", "-NORM.S.INV({a0})"]),
    "NormalStandard_IntervalProbability": (2, (), [
        "NORM.S.DIST({a1};TRUE)-NORM.S.DIST({a0};TRUE)",
        "NORM.S.DIST(-{a0};TRUE)-NORM.S.DIST(-{a1};TRUE)"]),
    "Normal_Density": (3, (), ["NORM.DIST({a0};{a1};{a2};FALSE)"]),
    "Normal_Cumulative": (3, (), ["NORM.DIST({a0};{a1};{a2};TRUE)"]),
    "Normal_Survival": (3, (), ["1-NORM.DIST({a0};{a1};{a2};TRUE)",
                                "NORM.S.DIST(-STANDARDIZE({a0};{a1};{a2});TRUE)"]),
    "Normal_ZScore": (3, (), ["STANDARDIZE({a0};{a1};{a2})"]),
    "Normal_InverseCumulative": (3, (), ["NORM.INV({a0};{a1};{a2})"]),
    "Normal_InverseSurvival": (3, (), ["NORM.INV(1-{a0};{a1};{a2})",
                                       "{a1}-{a2}*NORM.S.INV({a0})"]),
    "Lognormal_Density": (3, (), ["LOGNORM.DIST({a0};{a1};{a2};FALSE)"]),
    "Lognormal_Cumulative": (3, (), ["LOGNORM.DIST({a0};{a1};{a2};TRUE)"]),
    "Lognormal_Survival": (3, (), ["1-LOGNORM.DIST({a0};{a1};{a2};TRUE)",
                                   "NORM.S.DIST(-(LN({a0})-{a1})/{a2};TRUE)"]),
    "Lognormal_InverseCumulative": (3, (), ["LOGNORM.INV({a0};{a1};{a2})"]),
    "Lognormal_InverseSurvival": (3, (), ["LOGNORM.INV(1-{a0};{a1};{a2})",
                                          "EXP({a1}-{a2}*NORM.S.INV({a0}))"]),
    "StudentT_Density": (2, (1,), ["T.DIST({a0};{a1};FALSE)"]),
    "StudentT_Cumulative": (2, (1,), ["T.DIST({a0};{a1};TRUE)", "T.DIST.RT(-{a0};{a1})"]),
    "StudentT_Survival": (2, (1,), ["T.DIST.RT({a0};{a1})", "T.DIST(-{a0};{a1};TRUE)"]),
    "StudentT_InverseCumulative": (2, (1,), ["T.INV({a0};{a1})"]),
    "ChiSquare_Density": (2, (1,), ["CHISQ.DIST({a0};{a1};FALSE)"]),
    "ChiSquare_Cumulative": (2, (1,), ["CHISQ.DIST({a0};{a1};TRUE)"]),
    "ChiSquare_Survival": (2, (1,), ["CHISQ.DIST.RT({a0};{a1})"]),
    "ChiSquare_InverseCumulative": (2, (1,), ["CHISQ.INV({a0};{a1})"]),
    "F_Density": (3, (1, 2), ["F.DIST({a0};{a1};{a2};FALSE)"]),
    "F_Cumulative": (3, (1, 2), ["F.DIST({a0};{a1};{a2};TRUE)"]),
    "F_Survival": (3, (1, 2), ["F.DIST.RT({a0};{a1};{a2})"]),
    "F_InverseCumulative": (3, (1, 2), ["F.INV({a0};{a1};{a2})"]),
    "Gamma_Density": (3, (), ["GAMMA.DIST({a0};{a1};{a2};FALSE)"]),
    "Gamma_Cumulative": (3, (), ["GAMMA.DIST({a0};{a1};{a2};TRUE)"]),
    "Gamma_Survival": (3, (), ["1-GAMMA.DIST({a0};{a1};{a2};TRUE)"]),
    "Gamma_InverseCumulative": (3, (), ["GAMMA.INV({a0};{a1};{a2})"]),
    "Beta_Density": (3, (), ["BETA.DIST({a0};{a1};{a2};FALSE)"]),
    "Beta_Cumulative": (3, (), ["BETA.DIST({a0};{a1};{a2};TRUE)"]),
    "Beta_Survival": (3, (), ["1-BETA.DIST({a0};{a1};{a2};TRUE)",
                              "BETA.DIST(1-{a0};{a2};{a1};TRUE)"]),
    "Beta_InverseCumulative": (3, (), ["BETA.INV({a0};{a1};{a2})"]),
    "Exponential_Density": (2, (), ["EXPON.DIST({a0};{a1};FALSE)"]),
    "Exponential_Cumulative": (2, (), ["EXPON.DIST({a0};{a1};TRUE)"]),
    "Exponential_Survival": (2, (), ["1-EXPON.DIST({a0};{a1};TRUE)", "EXP(-{a1}*{a0})"]),
    "Weibull_Density": (3, (), ["WEIBULL.DIST({a0};{a1};{a2};FALSE)"]),
    "Weibull_Cumulative": (3, (), ["WEIBULL.DIST({a0};{a1};{a2};TRUE)"]),
    "Weibull_Survival": (3, (), ["1-WEIBULL.DIST({a0};{a1};{a2};TRUE)",
                                 "EXP(-({a0}/{a2})^{a1})"]),
    "Binomial_PMF": (3, (0, 1), ["BINOM.DIST({a0};{a1};{a2};FALSE)"]),
    "Binomial_Cumulative": (3, (0, 1), ["BINOM.DIST({a0};{a1};{a2};TRUE)"]),
    "Binomial_Survival": (3, (0, 1), ["1-BINOM.DIST({a0};{a1};{a2};TRUE)",
                                      "BINOM.DIST.RANGE({a1};{a2};{a0}+1;{a1})"]),
    "Binomial_InverseCumulative": (3, (1,), ["BINOM.INV({a1};{a2};{a0})"]),
    "Poisson_PMF": (2, (0,), ["POISSON.DIST({a0};{a1};FALSE)"]),
    "Poisson_Cumulative": (2, (0,), ["POISSON.DIST({a0};{a1};TRUE)"]),
    "Poisson_Survival": (2, (0,), ["1-POISSON.DIST({a0};{a1};TRUE)"]),
    "NegativeBinomial_PMF": (3, (0, 1), ["NEGBINOM.DIST({a0};{a1};{a2};FALSE)"]),
    "NegativeBinomial_Cumulative": (3, (0, 1), ["NEGBINOM.DIST({a0};{a1};{a2};TRUE)"]),
    "NegativeBinomial_Survival": (3, (0, 1), ["1-NEGBINOM.DIST({a0};{a1};{a2};TRUE)"]),
    "Geometric_PMF": (2, (0,), ["NEGBINOM.DIST({a0};1;{a1};FALSE)"]),
    "Geometric_Cumulative": (2, (0,), ["NEGBINOM.DIST({a0};1;{a1};TRUE)"]),
    "Geometric_Survival": (2, (0,), ["1-NEGBINOM.DIST({a0};1;{a1};TRUE)",
                                     "(1-{a1})^({a0}+1)"]),
    "Hypergeometric_PMF": (4, (0, 1, 2, 3), ["HYPGEOM.DIST({a0};{a1};{a2};{a3};FALSE)"]),
    "Hypergeometric_Cumulative": (4, (0, 1, 2, 3), ["HYPGEOM.DIST({a0};{a1};{a2};{a3};TRUE)"]),
    "Hypergeometric_Survival": (4, (0, 1, 2, 3), ["1-HYPGEOM.DIST({a0};{a1};{a2};{a3};TRUE)"]),
}
DISCRETE_INVERSE = {"Binomial_InverseCumulative"}


def token(text):
    """The argument as Excel will read it, or None when 15 digits do not write it exactly."""
    x = float(text)
    if not math.isfinite(x):
        return None
    if x == 0.0:
        return "0"
    short = f"{x:.15g}"
    if float(short) != x:
        return None
    return short.replace("e+", "E+").replace("e-", "E-")


def reference_ok(text):
    try:
        v = float(text)
    except ValueError:
        return False
    return math.isfinite(v) and v != 0.0


def points():
    """(point_id, kind, excel_formula, kstats_call, reference), and the skip counts."""
    out, skipped = [], {"no_counterpart": 0, "fractional_truncated": 0, "over_15_digits": 0,
                        "reference": 0}
    for source, prefix, path in SOURCES:
        with open(path, encoding="utf-8-sig") as f:
            rows = list(csv.DictReader(f))
        for number, r in enumerate(rows, 2):
            fn = r["function"]
            if fn not in EXCEL:
                skipped["no_counterpart"] += 1
                continue
            arity, truncated, forms = EXCEL[fn]
            raw = [r.get(f"arg{i + 1}", "") or "" for i in range(arity)]
            if any(not t.strip() for t in raw) or not reference_ok(r["reference"]):
                skipped["reference"] += 1
                continue
            values = [float(t) for t in raw]
            if any(values[i] != math.floor(values[i]) for i in truncated):
                skipped["fractional_truncated"] += 1
                continue
            toks = [token(t) for t in raw]
            if any(t is None for t in toks):
                skipped["over_15_digits"] += 1
                continue
            args = {f"a{i}": t for i, t in enumerate(toks)}
            kind = f"{source}:{number}:{fn}"
            call = f"K_STATS_{fn}(" + ";".join(toks) + ")"
            for j, form in enumerate(forms):
                pid = f"{prefix}{number:05d}{chr(97 + j)}"
                out.append((pid, kind, form.format(**args), call, r["reference"]))
    return out, skipped


def render():
    rows, skipped = points()
    buf = io.StringIO()
    w = csv.writer(buf, lineterminator="\n")
    w.writerow(FIELDS)
    for pid, kind, xl, ks, ref in rows:
        for field in (xl, ks, kind):
            if "," in field:
                raise SystemExit(f"a field holds a comma: {field}")
        w.writerow([pid, kind, xl, ks, ref, "", ""])
    return buf.getvalue(), rows, skipped


def main(argv=None):
    ap = argparse.ArgumentParser(description=(__doc__ or "").split("\n\n")[0])
    ap.add_argument("--check", action="store_true")
    a = ap.parse_args(argv)
    text, rows, skipped = render()
    points_n = len({r[1] for r in rows})
    if a.check:
        with open(GRID, encoding="utf-8") as f:
            committed = f.read()
        keys = [line.split(",")[:5] for line in committed.splitlines()]
        expected = [line.split(",")[:5] for line in text.splitlines()]
        if keys != expected:
            print("FAIL: excel_floor_grid.csv differs from floor_build.py; rerun it")
            return 1
        print(f"PASS: excel_floor_grid.csv matches its generator ({points_n} points)")
        return 0
    with open(GRID, "w", encoding="utf-8", newline="") as f:
        f.write(text)
    print(f"wrote {GRID}: {points_n} points, {len(rows)} rows; skipped {skipped}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
