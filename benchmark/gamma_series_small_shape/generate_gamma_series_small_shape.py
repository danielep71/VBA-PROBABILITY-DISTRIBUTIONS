"""
Build the #23 study grid from the frozen HOLDOUT_DESIGN.md, with its references.

The grid holds exactly the inputs the design froze before the production
source changed (bde92dd):

  holdout         kernel shapes 2E-6, 7E-6, ..., 2E-1 (11) at standardized
                  arguments 0.125 and 0.75, on two public routes each:
                    Gamma_Cumulative(X = z, Shape = s, ScaleParam = 1)
                    ChiSquare_Cumulative(X = 2z, DegreesFreedom = 2s)
                  Both transformations are exact in binary64, so the two
                  routes reach the same kernel P(s, z).
  scope_probe     the three real-Excel scope points of the design, Gamma_
                  Cumulative at (0.5, 0.6, 1), (1.0, 5.3, 1), (10.8, 10.3, 1).
  counterexample  the issue's minimal calls, Gamma_Cumulative(0.5, 1E-4, 1)
                  and ChiSquare_Cumulative(1, 2E-4).

No seam rows: the selected fix (MOVEMENT_MANIFEST.md, post-selection
explanation) uses no shape boundary, which is the condition of the design's
seam section.

Each reference is the regularized lower incomplete gamma P(a, x) at the binary64
inputs, at 60 significant digits (mpmath). The exporter echoes the binary64
inputs it parsed; analyze_gamma_series_small_shape.py checks the echoes equal
these inputs before it scores a row.

The design's table gives 25-digit scope references. They were evaluated at the
decimal inputs (5.3, 10.8, ...), not at the binary64 values the library
receives: they agree with a decimal-input evaluation to about 1E-25, and differ
from the binary64 one by up to 3.2E-16 relative, at (1.0, 5.3). That gap is a
tenth of the 3E-15 threshold, so the grid uses binary64 references, as the
design requires for the holdout and as #159 established for every survival
reference. --check recomputes both readings and reports the gap.

Columns (13): function, vba_kernel, claim, metric, arg1, arg2, arg3,
reference, regime, evidence_set, echo_arg1, echo_arg2, observed_vba. The first
ten identify the row; Export_GammaSeriesSmallShape fills the last three.

    python generate_gamma_series_small_shape.py           report only
    python generate_gamma_series_small_shape.py --write   create the grid; refused
                                                         once it exists
    python generate_gamma_series_small_shape.py --check   the committed grid's
                                                         identity columns equal a
                                                         fresh build
"""
import argparse
import csv
import io
import os
import sys

import mpmath as mp

HERE = os.path.dirname(os.path.abspath(__file__))
GRID = os.path.join(HERE, "gamma_series_small_shape_grid.csv")
FIELDS = ["function", "vba_kernel", "claim", "metric", "arg1", "arg2", "arg3",
          "reference", "regime", "evidence_set", "echo_arg1", "echo_arg2", "observed_vba"]
KEY_COLUMNS = 10
CLAIM = "rel<=3E-15"
EVIDENCE_SET = "gamma_series_small_shape"

# HOLDOUT_DESIGN.md, "Independent small-shape holdout" and "Real-Excel scope probe".
HOLDOUT_SHAPES = (2e-6, 7e-6, 2e-5, 7e-5, 2e-4, 7e-4, 2e-3, 7e-3, 2e-2, 7e-2, 2e-1)
HOLDOUT_ARGS = (0.125, 0.75)
SCOPE_PROBES = ((0.5, 0.6), (1.0, 5.3), (10.8, 10.3))      # (X, Shape), ScaleParam 1
DESIGN_REFERENCES = {(0.5, 0.6): "0.6189010171407936557816432",
                     (1.0, 5.3): "0.002157194707472261441497618",
                     (10.8, 10.3): "0.6014210971218611799875902"}
# The issue's minimal counterexample and its Chi-square equivalent.
COUNTEREXAMPLE = (0.5, 1e-4)


def lower_p(shape, x):
    """P(shape, x) at the binary64 values, 60 significant digits."""
    with mp.workdps(80):
        return mp.gammainc(mp.mpf(shape), 0, mp.mpf(x), regularized=True)


def row(function, a1, a2, a3, reference, regime):
    return {"function": function, "vba_kernel": "K_STATS_" + function, "claim": CLAIM,
            "metric": "rel", "arg1": repr(a1), "arg2": repr(a2),
            "arg3": "" if a3 is None else repr(a3), "reference": mp.nstr(reference, 60),
            "regime": regime, "evidence_set": EVIDENCE_SET,
            "echo_arg1": "", "echo_arg2": "", "observed_vba": ""}


def gamma_and_chisquare(shape, z, regime):
    reference = lower_p(shape, z)
    return [row("Gamma_Cumulative", z, shape, 1.0, reference, regime),
            row("ChiSquare_Cumulative", 2.0 * z, 2.0 * shape, None, reference, regime)]


def build():
    rows = []
    for z in HOLDOUT_ARGS:
        for shape in HOLDOUT_SHAPES:
            rows += gamma_and_chisquare(shape, z, "holdout")
    for x, shape in SCOPE_PROBES:
        rows.append(row("Gamma_Cumulative", x, shape, 1.0, lower_p(shape, x), "scope_probe"))
    rows += gamma_and_chisquare(COUNTEREXAMPLE[1], COUNTEREXAMPLE[0], "counterexample")
    return rows


def render(rows):
    buf = io.StringIO()
    writer = csv.DictWriter(buf, fieldnames=FIELDS, lineterminator="\n")
    writer.writeheader()
    writer.writerows(rows)
    return buf.getvalue()


def design_agreement():
    """The design's 25-digit scope references, read as values at the decimal
    inputs: (worst gap to a decimal-input evaluation, worst gap to the binary64
    references the grid uses)."""
    decimal_gap = binary_gap = 0
    for (x, shape), text in DESIGN_REFERENCES.items():
        with mp.workdps(80):
            theirs = mp.mpf(text)
            decimal = mp.gammainc(mp.mpf(repr(shape)), 0, mp.mpf(repr(x)), regularized=True)
            decimal_gap = max(decimal_gap, abs(decimal - theirs) / theirs)
            binary_gap = max(binary_gap, abs(lower_p(shape, x) - theirs) / theirs)
    return float(decimal_gap), float(binary_gap)


def identity(text):
    return [r[:KEY_COLUMNS] for r in csv.reader(io.StringIO(text))]


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    mode = ap.add_mutually_exclusive_group()
    mode.add_argument("--write", action="store_true")
    mode.add_argument("--check", action="store_true")
    a = ap.parse_args(argv)
    rows = build()
    text = render(rows)
    decimal_gap, binary_gap = design_agreement()
    if decimal_gap > 1e-24:
        print(f"FAIL: HOLDOUT_DESIGN.md's scope references are not P at the decimal inputs "
              f"(gap {decimal_gap:.2e})")
        return 1
    note = (f"the design's scope references are P at the decimal inputs (to {decimal_gap:.1e}); "
            f"the binary64 references differ from them by up to {binary_gap:.2e}")
    if a.check:
        with open(GRID, encoding="utf-8") as f:
            committed = f.read()
        if identity(committed) != identity(text):
            print("FAIL: the grid's identity columns differ from a fresh build of the design")
            return 1
        print(f"PASS: {len(rows)} rows match the frozen design; {note}")
        return 0
    if a.write:
        if os.path.exists(GRID):
            print(f"refusing: {GRID} exists; it may hold observations. Use --check.")
            return 1
        with open(GRID, "w", encoding="utf-8", newline="") as f:
            f.write(text)
        print(f"wrote {len(rows)} rows to {GRID}")
        return 0
    counts = {}
    for r in rows:
        counts[r["regime"]] = counts.get(r["regime"], 0) + 1
    print(f"{len(rows)} rows {counts}; {note}. Report only: pass --write to create the grid.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
