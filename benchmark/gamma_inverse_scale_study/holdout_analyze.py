"""
#14: verdict of the holdout against PREREGISTRATION.md.

Reads gamma_inverse_scale_grid.csv as the exporter filled it. Every rule is stated in the
preregistration; the bounds are fixed there and nothing is fitted here.

  G1  accuracy    where the reference is a normal Double, the Gamma inverse
                  returns a number within BOUND relative of it.
  G2  limits      a reference below every Double returns the exact zero; one
                  above every Double returns #NUM!; a subnormal one is within
                  one unit of the subnormal spacing plus BOUND relative.
  G3  monotone    within a seam cluster, over increasing probability, the
                  quantile does not fall by more than BOUND times the sum of
                  the two references: the invariant of #327.
  G4  families    at scale 2, ChiSquare_InverseCumulative(p, 2A) returns the
                  same Double, or the same error, as the Gamma inverse.
  G5  round trip  where the quantile is a normal Double, Gamma_Cumulative at
                  it is within (FORWARD + A * BOUND) relative of p.
  G6  numbers     no cell holds a runtime fault or a failed kernel.
  G7  seam        every seam cluster holds both branches, by the exported
                  L0, so the crossover was probed in Excel and not only in
                  the model.

Reported, not gating: the worst error by kind and by branch, and whether the
rows answered in log space stay below PREDICTED.

The rules serve both holdouts: the first (PREREGISTRATION.md), which was
rejected on G7, and the second (PREREGISTRATION_2.md, --holdout 2).

Usage: python3 holdout_analyze.py [grid.csv] [--holdout 2]
Exit 0 only when every rule holds on every row.
"""
import argparse
import csv
import os
import struct
import sys
from collections import defaultdict

from mpmath import mp, mpf

HERE = os.path.dirname(os.path.abspath(__file__))
GRID = os.path.join(HERE, "gamma_inverse_scale_grid.csv")
GRID_2 = os.path.join(HERE, "gamma_inverse_scale_grid2.csv")
BOUND = mpf("1e-12")          # inverse guarantee of #327
FORWARD = mpf("5e-13")        # loosest forward contract in the region (#13)
PREDICTED = mpf("1e-14")
PREDICTED_2 = mpf("5e-15")
CROSSOVER = -40.0
MIN_NORMAL = mpf(2) ** -1022
MIN_SUBNORMAL = mpf(2) ** -1074
ZERO_BITS = "0" * 16
OBSERVED = ("gamma_inv", "chi_inv", "gamma_cdf", "log_unit")


def from_bits(text):
    return struct.unpack("<d", struct.pack("<Q", int(text, 16)))[0]


def is_bits(text):
    return len(text) == 16 and all(c in "0123456789ABCDEFabcdef" for c in text)


def number(text):
    """The Double in a cell, or None for CVERR, NA, a fault or a blank."""
    return from_bits(text) if is_bits(text) else None


def analyze(rows):
    mp.dps = 60
    failures, counts = [], defaultdict(int)
    worst = defaultdict(lambda: mpf(0))
    clusters = defaultdict(list)

    def fail(rule, row, text):
        failures.append(f"{rule} {row['point_id']} ({row['kind']}): {text}")

    for r in rows:
        if not all(r.get(c, "").strip() for c in OBSERVED):
            fail("G6", r, "an observed cell is blank: the grid was not exported")
            continue
        cells = {c: r[c].strip() for c in OBSERVED}
        for c, v in cells.items():
            if v in ("ERROR", "FAIL"):
                fail("G6", r, f"{c} is {v}")
        p, a, s = (from_bits(r[k]) for k in ("p_bits", "a_bits", "s_bits"))
        value = number(cells["gamma_inv"])
        ref_text = r["ref_x"].strip()
        log_unit = number(cells["log_unit"])
        branch = None if log_unit is None else ("log" if log_unit <= CROSSOVER else "ordinary")
        counts[f"branch: {branch}"] += 1

        if ref_text == "NUM":
            counts["overflow"] += 1
            if cells["gamma_inv"] != "CVERR":
                fail("G2", r, f"expected #NUM!, got {cells['gamma_inv']}")
        elif ref_text == "ZERO":
            counts["exact zero"] += 1
            if cells["gamma_inv"] != ZERO_BITS:
                fail("G2", r, f"expected the exact zero, got {cells['gamma_inv']}")
        else:
            ref = mpf(ref_text)
            if value is None:
                fail("G1", r, f"expected a number, got {cells['gamma_inv']}")
            elif ref < MIN_NORMAL:
                counts["subnormal"] += 1
                if abs(mpf(value) - ref) > MIN_SUBNORMAL + BOUND * ref:
                    fail("G2", r, f"subnormal quantile {value!r} against {ref_text}")
            else:
                counts["scored"] += 1
                error = abs(mpf(value) - ref) / ref
                worst[f"kind {r['kind']}"] = max(worst[f"kind {r['kind']}"], error)
                worst[f"branch {branch}"] = max(worst[f"branch {branch}"], error)
                if error > BOUND:
                    fail("G1", r, f"relative error {mp.nstr(error, 4)}")
                if r["kind"] == "seam":
                    clusters[r["a_bits"]].append((p, value, ref, branch, r))

        if s == 2.0:
            counts["families"] += 1
            if cells["chi_inv"] != cells["gamma_inv"]:
                fail("G4", r, f"Chi-square {cells['chi_inv']} against Gamma {cells['gamma_inv']}")
        elif cells["chi_inv"] != "NA":
            fail("G4", r, "a Chi-square value where the scale is not 2")

        if value is not None and mpf(value) >= MIN_NORMAL:
            forward = number(cells["gamma_cdf"])
            if forward is None:
                fail("G5", r, f"Gamma_Cumulative at the quantile is {cells['gamma_cdf']}")
            else:
                counts["round trips"] += 1
                residual = abs(mpf(forward) - mpf(p)) / mpf(p)
                worst["round trip"] = max(worst["round trip"], residual)
                if residual > FORWARD + mpf(a) * BOUND:
                    fail("G5", r, f"round-trip residual {mp.nstr(residual, 4)}")

    for a_bits, cluster in sorted(clusters.items()):
        cluster.sort(key=lambda item: item[0])
        branches = {item[3] for item in cluster}
        if not {"log", "ordinary"} <= branches:
            fail("G7", cluster[0][4], f"shape {from_bits(a_bits)!r}: only {sorted(map(str, branches))}")
        for left, right in zip(cluster, cluster[1:]):
            counts["monotone pairs"] += 1
            if mpf(right[1]) < mpf(left[1]) - BOUND * (left[2] + right[2]):
                fail("G3", right[4], f"quantile falls from {left[1]!r} to {right[1]!r}")
    return failures, dict(counts), {k: v for k, v in worst.items()}


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    ap.add_argument("grid", nargs="?", default=None)
    ap.add_argument("--holdout", type=int, choices=(1, 2), default=1)
    a = ap.parse_args(argv)
    path = a.grid or (GRID if a.holdout == 1 else GRID_2)
    with open(path, newline="", encoding="utf-8") as handle:
        rows = list(csv.DictReader(handle))
    failures, counts, worst = analyze(rows)
    print(f"{len(rows)} rows; counts {dict(sorted(counts.items()))}")
    for key in sorted(worst):
        print(f"   worst relative error, {key}: {mp.nstr(worst[key], 4)}")
    log_worst = worst.get("branch log")
    predicted = PREDICTED if a.holdout == 1 else PREDICTED_2
    if log_worst is not None:
        print(f"   prediction (log space below {mp.nstr(predicted, 2)}): "
              f"{'met' if log_worst < predicted else 'NOT met'}")
    for line in failures[:40]:
        print("  " + line)
    if len(failures) > 40:
        print(f"  ... {len(failures) - 40} more")
    if failures:
        print(f"REJECTED: {len(failures)} failure(s)")
        return 1
    print("ACCEPTED: G1 to G7 hold on every row")
    return 0


if __name__ == "__main__":
    sys.exit(main())
