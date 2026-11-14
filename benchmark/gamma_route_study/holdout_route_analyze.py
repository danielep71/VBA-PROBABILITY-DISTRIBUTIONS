"""
#314: verdict of the routing holdout against FAST_ROUTE_PREREGISTRATION.md.

Reads holdout_route_grid.csv as the exporter filled it. Every rule is stated
in the preregistration; the bounds are fixed there and nothing is fitted here.

  R1  accuracy    every public tail whose reference is at least FLOOR is within
                  BOUND relative of it. Where the reference is an exact 0 or 1
                  (X = 0, or an upper tail below every Double), the value is
                  that exact Double.
  R2  families    the Gamma and the Chi-square route return the same Double.
  R3  monotone    per shape, over increasing X, the CDF does not fall and the
                  survival does not rise by more than the two rows' allowances
                  (BOUND times each reference): the invariant of #327.
  R4  complement  abs(P + Q - 1) <= BOUND * (P + Q).
  R5  delegation  a tail whose reference is below FLOOR is returned by the
                  double-double entry point: the public value equals dd_p or
                  dd_q bit for bit.
  R6  numbers     every public route returns a number at every row.

Reported, not gating: the worst error by row kind, how many public values
equal the model's (fast_route.routed) bit for bit, and how many equal the
double-double entry points'.

Usage: python3 holdout_route_analyze.py [grid.csv] [--no-model]
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
sys.path.insert(0, HERE)

GRID = os.path.join(HERE, "holdout_route_grid.csv")
BOUND = mpf("1e-13")
FLOOR = mpf("1e-290")
DELEGATED_BELOW = mpf("0.99e-290")
PUBLIC = ("gamma_cdf", "gamma_sf", "chi_cdf", "chi_sf")
STATUSES = ("CVERR", "ERROR", "FAIL", "")


def from_bits(text):
    return struct.unpack("<d", struct.pack("<Q", int(text, 16)))[0]


def value(cell):
    """The Double in a result cell, or None for a status or a blank."""
    cell = (cell or "").strip()
    if cell in STATUSES or len(cell) != 16:
        return None
    try:
        return from_bits(cell)
    except ValueError:
        return None


def load(path):
    with open(path, newline="") as handle:
        rows = list(csv.DictReader(handle))
    for r in rows:
        r["a"], r["x"] = from_bits(r["a_bits"]), from_bits(r["x_bits"])
        r["refs"] = {"P": mpf(r["ref_p"]) if r["ref_p"] else None,
                     "Q": mpf(r["ref_q"]) if r["ref_q"] else None}
        r["exact"] = {"P": r["ref_p"] in ("0", "1"), "Q": r["ref_q"] in ("0", "1")}
        r["v"] = {k: value(r[k]) for k in PUBLIC + ("dd_p", "dd_q")}
    return rows


def relative(v, ref):
    return float(abs(mpf(v) - ref) / abs(ref))


def analyze(rows, model=None):
    """(failures, report). model(a, x) -> (P, Q), or None to skip the fidelity count."""
    mp.dps = 60
    fails = []
    worst = defaultdict(float)
    counts = defaultdict(int)

    def fail(rule, r, text):
        fails.append(f"{rule} {r['point_id']} ({r['kind']}): {text}")

    for r in rows:
        v = r["v"]
        # R6
        missing = [k for k in PUBLIC if v[k] is None]
        if missing:
            fail("R6", r, "no number from " + ", ".join(f"{k}={r[k] or 'blank'}" for k in missing))
            continue
        # R2
        if v["gamma_cdf"] != v["chi_cdf"] or v["gamma_sf"] != v["chi_sf"]:
            fail("R2", r, "the Gamma and Chi-square routes differ")
        for tail, key, dd in (("P", "gamma_cdf", "dd_p"), ("Q", "gamma_sf", "dd_q")):
            ref, got = r["refs"][tail], v[key]
            if ref is None:
                counts["unevaluated"] += 1
                continue
            if r["exact"][tail]:
                counts["exact"] += 1
                if got != float(ref):
                    fail("R1", r, f"{tail} = {got!r}, the exact value is {float(ref)!r}")
            elif ref < DELEGATED_BELOW:
                counts["below floor"] += 1
                if v[dd] is None or got != v[dd]:
                    fail("R5", r, f"{tail} = {got!r} is not the double-double value {r[dd]}")
            elif ref < FLOOR:
                counts["at the floor"] += 1     # within 1% of FLOOR: either path
            else:
                e = relative(got, ref)
                counts["scored"] += 1
                worst[r["kind"]] = max(worst[r["kind"]], e)
                if e > BOUND:
                    fail("R1", r, f"{tail} relative error {e:.3e}")
            if v[dd] is not None and got == v[dd]:
                counts["equal to double-double"] += 1
        # R4
        p, q = mpf(v["gamma_cdf"]), mpf(v["gamma_sf"])
        if abs(p + q - 1) > BOUND * (p + q):
            fail("R4", r, f"P + Q - 1 = {float(p + q - 1):.3e}")
        if model is not None and r["x"] > 0:
            mp_, mq_ = model(r["a"], r["x"])
            counts["equal to model"] += (v["gamma_cdf"] == mp_) + (v["gamma_sf"] == mq_)
            counts["compared with model"] += 2

    # R3
    by_shape = defaultdict(list)
    for r in rows:
        if all(r["v"][k] is not None for k in PUBLIC):
            by_shape[r["a_bits"]].append(r)
    for shape_rows in by_shape.values():
        shape_rows.sort(key=lambda r: r["x"])
        for left, right in zip(shape_rows, shape_rows[1:]):
            for tail, key, sign in (("P", "gamma_cdf", 1), ("Q", "gamma_sf", -1)):
                refs = (left["refs"][tail], right["refs"][tail])
                if None in refs:
                    continue
                back = sign * (mpf(left["v"][key]) - mpf(right["v"][key]))
                if back > BOUND * (refs[0] + refs[1]):
                    fail("R3", right, f"{tail} runs backwards by {float(back):.3e} from "
                                      f"{left['point_id']}")
                counts["monotone pairs"] += 1
    return fails, {"worst": dict(worst), "counts": dict(counts), "rows": len(rows)}


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    ap.add_argument("grid", nargs="?", default=GRID)
    ap.add_argument("--no-model", action="store_true",
                    help="skip the bit-for-bit comparison with the Python model")
    a = ap.parse_args(argv)
    rows = load(a.grid)
    model = None
    if not a.no_model:
        import fast_route
        model = fast_route.routed
    fails, report = analyze(rows, model)
    print(f"{report['rows']} rows; counts {report['counts']}")
    for kind, e in sorted(report["worst"].items()):
        print(f"   worst relative error, {kind}: {e:.3e}")
    if fails:
        print(f"REJECTED: {len(fails)} failure(s)")
        for f in fails[:40]:
            print("   " + f)
        return 1
    print("ACCEPTED: R1 to R6 hold on every row")
    return 0


if __name__ == "__main__":
    sys.exit(main())
