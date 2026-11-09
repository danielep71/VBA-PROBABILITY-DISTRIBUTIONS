"""
Build the #24 holdout grid from the frozen HOLDOUT_DESIGN.md, with its references.

The rows are exactly design.py's inputs: for each shape A and each argument X,
Gamma_Cumulative(X, A, 1), Gamma_Survival(X, A, 1), ChiSquare_Cumulative(2X, 2A)
and ChiSquare_Survival(2X, 2A). References are P and Q computed separately by
mpmath at the binary64 inputs, at 80 and 120 digits; the two readings must agree
to 1E-60 relative. Q is never formed as 1 - P.

Columns (13), as the #23 study grid: function, vba_kernel, claim, metric,
arg1, arg2, arg3, reference, regime, evidence_set, echo_arg1, echo_arg2,
observed_vba. The first ten are the design; the exporter
(M_STATS_PROBDIST_GLQ.bas) writes the last three.

Run:
    python3 generate_gamma_lower_q.py --write   (only when the grid is absent)
    python3 generate_gamma_lower_q.py --check   (the committed grid is the design)
"""
import argparse
import csv
import io
import os
import sys

import mpmath as mp

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
from design import SHAPES, points  # noqa: E402

GRID = os.path.join(HERE, "gamma_lower_q_grid.csv")
FIELDS = ["function", "vba_kernel", "claim", "metric", "arg1", "arg2", "arg3",
          "reference", "regime", "evidence_set", "echo_arg1", "echo_arg2", "observed_vba"]
KEY_COLUMNS = 10
CLAIMS = {"Cumulative": "rel<=3E-15", "Survival": "rel<=5E-15"}
EVIDENCE_SET = "gamma_lower_q"


def tail(shape, x, upper, dps):
    with mp.workdps(dps):
        a, z = mp.mpf(shape), mp.mpf(x)
        if upper:
            return mp.gammainc(a, z, mp.inf, regularized=True)
        return mp.gammainc(a, 0, z, regularized=True)


def reference(shape, x, upper):
    low, high = tail(shape, x, upper, 80), tail(shape, x, upper, 120)
    with mp.workdps(120):
        if not high > 0 or abs(low - high) > mp.mpf("1E-60") * high:
            raise ValueError(f"unstable reference at A={shape!r}, X={x!r}, upper={upper}")
        return mp.nstr(high, 60)


def row(function, a1, a2, a3, ref, kind):
    return {"function": function, "vba_kernel": "K_STATS_" + function, "claim": CLAIMS[kind],
            "metric": "rel", "arg1": repr(a1), "arg2": repr(a2),
            "arg3": "" if a3 is None else repr(a3), "reference": ref,
            "regime": "holdout", "evidence_set": EVIDENCE_SET,
            "echo_arg1": "", "echo_arg2": "", "observed_vba": ""}


def build():
    rows = []
    for shape in SHAPES:
        for x in points(shape):
            for kind, upper in (("Cumulative", False), ("Survival", True)):
                ref = reference(shape, x, upper)
                rows.append(row("Gamma_" + kind, x, shape, 1.0, ref, kind))
                rows.append(row("ChiSquare_" + kind, 2.0 * x, 2.0 * shape, None, ref, kind))
    return rows


def render(rows):
    buf = io.StringIO()
    writer = csv.DictWriter(buf, fieldnames=FIELDS, lineterminator="\n")
    writer.writeheader()
    writer.writerows(rows)
    return buf.getvalue()


def identity(text):
    return [r[:KEY_COLUMNS] for r in csv.reader(io.StringIO(text))]


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    mode = ap.add_mutually_exclusive_group(required=True)
    mode.add_argument("--write", action="store_true")
    mode.add_argument("--check", action="store_true")
    a = ap.parse_args(argv)
    rows = build()
    text = render(rows)
    if a.check:
        with open(GRID, encoding="utf-8") as f:
            committed = f.read()
        if identity(committed) != identity(text):
            print("FAIL: the grid's identity columns differ from a fresh build of the design")
            return 1
        print(f"PASS: {len(rows)} rows match the frozen #24 design")
        return 0
    if os.path.exists(GRID):
        print(f"refusing: {GRID} exists; it may hold observations. Use --check.")
        return 1
    with open(GRID, "w", encoding="utf-8", newline="") as f:
        f.write(text)
    print(f"wrote {len(rows)} rows to {GRID}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
