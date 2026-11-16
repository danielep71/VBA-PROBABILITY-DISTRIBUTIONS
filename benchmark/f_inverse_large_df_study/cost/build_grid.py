"""
#359: the holdout of the cheaper beta inverse (PREREGISTRATION.md).

One design, two files with the same rows: beta_inverse_cost_baseline.csv, which
Excel fills from the source of v1.0.0, and beta_inverse_cost_grid.csv, which
Excel fills from the candidate. No point of the fitting sweep (fit_sweep.py),
of the production holdout of #35 or of the cost benchmark is repeated here.

  direct_beta   Beta at seven pairs of shapes of at least 0.5 and six
                probabilities from 1E-10 to 1 - 1E-10;
  small_shape   Beta with a shape below 0.5, which the Python model of the
                candidate does not cover, at four probabilities;
  balanced_f    F at six pairs of moderate degrees of freedom and six
                probabilities;
  large_f       F with one degree of freedom far above the other, where the
                solved member sits on the reflected branch, at six
                probabilities;
  branch_seam   Beta at the probability whose quantile is the boundary
                between the direct and the reflected branch,
                (a + 1) / (a + b + 2), moved by 0, 1 and 4 units in the last
                place either way;
  half_seam     F at the orientation boundary: the lower tail at 1/2, moved
                by 0, 1 and 4 units in the last place either way;
  student_t     Student-t through its small-df branch, away from the median.

Columns (8), as the production holdout: point_id, kind, func, p_bits, a_bits,
b_bits, ref, value. M_STATS_PROBDIST_BIC.bas writes value. ref is the quantile
to 40 digits, from 120-digit arithmetic.

Run:
    python3 build_grid.py --write   (only when the files are absent)
    python3 build_grid.py --check   (the committed files are the design)
"""
import argparse
import importlib.util
import itertools
import os
import sys

import mpmath as mp

HERE = os.path.dirname(os.path.abspath(__file__))


def _load_production():
    """The grid helpers of the production holdout, under a name of their own:
    that file is also called build_grid.py."""
    path = os.path.join(os.path.dirname(HERE), "production", "build_grid.py")
    spec = importlib.util.spec_from_file_location("f35_production_build_grid", path)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


production = _load_production()

FILES = {"baseline": os.path.join(HERE, "beta_inverse_cost_baseline.csv"),
         "candidate": os.path.join(HERE, "beta_inverse_cost_grid.csv")}
TAILS = (1e-10, 0.05, 0.35, 0.65, 0.95, 1 - 1e-10)


def seam_probability(a, b, x):
    """The Double nearest I_x(a, b), from 120-digit arithmetic."""
    with mp.workdps(production.DIGITS):
        return float(mp.betainc(mp.mpf(a), mp.mpf(b), 0, mp.mpf(x), regularized=True))


def points():
    rows = [("direct_beta", "Beta", p, a, b) for (a, b), p in itertools.product(
        ((0.6, 9.0), (9.0, 0.6), (1.3, 45.0), (45.0, 1.3), (8.0, 11.0), (140.0, 220.0),
         (3.5, 1800.0)), TAILS)]
    rows += [("small_shape", "Beta", p, a, b) for (a, b), p in itertools.product(
        ((0.2, 3.0), (3.0, 0.2), (0.15, 700.0), (0.4, 0.45), (0.25, 60.0)),
        (1e-7, 0.25, 0.75, 1 - 1e-7))]
    rows += [("balanced_f", "F", p, d1, d2) for (d1, d2), p in itertools.product(
        ((4.0, 9.0), (9.0, 4.0), (60.0, 60.0), (250.0, 130.0), (2.0, 2.0), (1.0, 7.0)), TAILS)]
    rows += [("large_f", "F", p, d1, d2) for (d1, d2), p in itertools.product(
        ((11.0, 6e5), (6e5, 11.0), (35.0, 7e7), (7e7, 35.0), (220.0, 5e9), (5e9, 220.0)),
        (1e-9, 0.15, 0.45, 0.55, 0.85, 1 - 1e-9))]
    for a, b in ((3.0, 40.0), (0.8, 12.0)):
        h = seam_probability(a, b, (a + 1.0) / (a + b + 2.0))
        rows += [("branch_seam", "Beta", production.step(h, n), a, b) for n in (-4, -1, 0, 1, 4)]
    for d1, d2 in ((6.0, 10.0), (10.0, 6.0)):
        h = seam_probability(d1 / 2.0, d2 / 2.0, 0.5)
        rows += [("half_seam", "F", production.step(h, n), d1, d2) for n in (-4, -1, 0, 1, 4)]
    rows += [("student_t", "T", p, df, None) for df in (1.0, 2.0, 4.0, 7.0, 15.0, 40.0)
             for p in (0.55, 0.7, 0.9, 0.999)]
    return rows


def build():
    out, counts = [], {}
    for kind, func, p, a, b in points():
        n = counts.get(kind, 0)
        counts[kind] = n + 1
        with mp.workdps(production.DIGITS):
            ref = mp.nstr(production.reference(func, p, a, b), 40)
        out.append({"point_id": f"{kind[:2]}{n:03d}", "kind": kind, "func": func,
                    "p_bits": production.bits(p), "a_bits": production.bits(a),
                    "b_bits": "NA" if b is None else production.bits(b), "ref": ref})
    return out


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    mode = ap.add_mutually_exclusive_group(required=True)
    mode.add_argument("--write", action="store_true")
    mode.add_argument("--check", action="store_true")
    a = ap.parse_args(argv)
    text = production.render(build())
    for name, path in FILES.items():
        if a.check:
            with open(path, encoding="utf-8") as handle:
                if production.identity(handle.read()) != production.identity(text):
                    print(f"FAIL: the identity columns of the {name} file differ from a fresh build")
                    return 1
        elif os.path.exists(path):
            print(f"refusing: {path} exists; it may hold observations. Use --check.")
            return 1
        else:
            with open(path, "w", encoding="utf-8", newline="") as handle:
                handle.write(text)
    print(f"{'PASS' if a.check else 'wrote'}: {len(production.identity(text)) - 1} rows in each of two files")
    return 0


if __name__ == "__main__":
    sys.exit(main())
