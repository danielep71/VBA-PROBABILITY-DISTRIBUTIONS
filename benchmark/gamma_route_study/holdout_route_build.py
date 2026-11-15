"""
#314: materialize the routing holdout and its mechanical boundary probes.

Rows of holdout_route_grid.csv, in this order per shape:

  holdout      the 117 frozen points of holdout_design.py;
  zero         X = 0, where the tails are exactly 0 and 1;
  transition   for every change of the rule's (P source, Q source) along X:
               the last Double of the old state, the first of the new one,
               and its successor;
  shape_probe  at the internal boundary A = 1: the predecessor of 1, 1 and
               its successor, at every small-shape argument of the design
               and at their own X = fl(A + 1) triples.

Transitions are located from the rule's predicate alone (fast_route.paths),
never from an error or a reference: X runs over a ladder of powers of two,
2^-1074 to 2^1020 in steps of 2^6, with fl(A + 1) and the shape's holdout
points added, and each interval whose ends differ is bisected on the integer
bit pattern. A change that begins and ends between two ladder points is not
found; the ladder is part of the frozen design.

Columns (12): point_id, kind, a_bits, x_bits, ref_p, ref_q, gamma_cdf,
gamma_sf, chi_cdf, chi_sf, dd_p, dd_q. The first six are the design and its
references; M_STATS_PROBDIST_GRH.bas writes the rest. a_bits and x_bits are
the 16 hexadecimal digits of the Doubles, so no decimal is parsed in Excel.
References are those of the map (generate_binary64_map.references), at 40
and 60 digits; where X is so large that the logarithm of the upper tail is
below -2000, the references are the exact 1 and 0. A row whose reference
cannot be computed keeps both fields blank and is reported as unevaluated.

Run:
    python3 holdout_route_build.py --write   (only when the grid is absent)
    python3 holdout_route_build.py --check   (the committed grid is the design)
"""
import argparse
import csv
import io
import math
import os
import struct
import sys

from mpmath import loggamma, mp, mpf

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import fast_route as F  # noqa: E402
import holdout_design as H  # noqa: E402
from generate_binary64_map import references  # noqa: E402
sys.path.insert(0, os.path.dirname(HERE))
from _igamma import IGammaNonConvergence  # noqa: E402

GRID = os.path.join(HERE, "holdout_route_grid.csv")
FIELDS = ["point_id", "kind", "a_bits", "x_bits", "ref_p", "ref_q", "gamma_cdf", "gamma_sf",
          "chi_cdf", "chi_sf", "dd_p", "dd_q"]
KEY_COLUMNS = 6
LADDER = tuple(math.ldexp(1.0, k) for k in range(-1074, 1021, 6))
BOUNDARY_SHAPES = (math.nextafter(1.0, 0.0), 1.0, math.nextafter(1.0, 2.0))


def bits(x):
    return struct.unpack("<Q", struct.pack("<d", x))[0]


def from_bits(n):
    return struct.unpack("<d", struct.pack("<Q", n))[0]


def hexbits(x):
    return f"{bits(x):016X}"


def seam_triple(a):
    seam = a + 1.0
    return (math.nextafter(seam, 0.0), seam, math.nextafter(seam, math.inf))


def transitions(a, anchors):
    """[(last of old state, first of new state, its successor, old, new)]."""
    xs = sorted(set(LADDER) | set(anchors) | {a + 1.0})
    states = [F.paths(a, x) for x in xs]
    found = []
    for (lo, hi), (s_lo, s_hi) in zip(zip(xs, xs[1:]), zip(states, states[1:])):
        if s_lo == s_hi:
            continue
        low, high = bits(lo), bits(hi)
        while high - low > 1:
            mid = (low + high) // 2
            if F.paths(a, from_bits(mid)) == s_lo:
                low = mid
            else:
                high = mid
        first = from_bits(high)
        found.append((from_bits(low), first, math.nextafter(first, math.inf), s_lo,
                      F.paths(a, first)))
    return found


def reference_pair(a, x):
    if x == 0.0:
        return "0", "1"
    # Far above the body the upper tail is below every Double: log Q is about
    # (A - 1) log X - X - log Gamma(A). The references are then the exact 1
    # and 0, and the high-precision routine is not asked to sum that far.
    with mp.workdps(40):
        log_q = (mpf(a) - 1) * mp.log(mpf(x)) - mpf(x) - loggamma(mpf(a))
    if x > a + 1.0 and log_q < -2000:
        return "1", "0"
    try:
        return references(a, x)
    except (ValueError, ArithmeticError, IGammaNonConvergence):
        return "", ""


def design():
    """[(point_id, kind, a, x)] with exact duplicates of (a, x) removed."""
    out, seen = [], set()

    def add(pid, kind, a, x):
        if (a, x) in seen:
            return
        seen.add((a, x))
        out.append((pid, kind, a, x))

    for si, a in enumerate(H.SHAPES):
        pts = H.points(a)
        for xi, x in enumerate(pts):
            add(f"h{si:02d}_x{xi:02d}", "holdout", a, x)
        add(f"h{si:02d}_zero", "zero", a, 0.0)
        for ti, (last, first, after, _old, _new) in enumerate(transitions(a, pts)):
            for tag, x in (("a", last), ("b", first), ("c", after)):
                add(f"h{si:02d}_t{ti:02d}{tag}", "transition", a, x)
    for bi, a in enumerate(BOUNDARY_SHAPES):
        for xi, x in enumerate(sorted(set(H.SMALL_X) | set(seam_triple(a)))):
            add(f"b{bi}_x{xi:02d}", "shape_probe", a, x)
    return out


def build():
    rows = []
    for pid, kind, a, x in design():
        ref_p, ref_q = reference_pair(a, x)
        rows.append({"point_id": pid, "kind": kind, "a_bits": hexbits(a), "x_bits": hexbits(x),
                     "ref_p": ref_p, "ref_q": ref_q, **{k: "" for k in FIELDS[KEY_COLUMNS:]}})
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
    text = render(build())
    if a.check:
        with open(GRID, encoding="utf-8") as handle:
            committed = handle.read()
        if identity(committed) != identity(text):
            print("FAIL: the routing holdout's identity columns differ from a fresh build")
            return 1
        print(f"PASS: {len(identity(text)) - 1} rows match the frozen design and its probes")
        return 0
    if os.path.exists(GRID):
        print(f"refusing: {GRID} exists; it may hold observations. Use --check.")
        return 1
    with open(GRID, "w", encoding="utf-8", newline="") as handle:
        handle.write(text)
    print(f"wrote {len(identity(text)) - 1} rows to {GRID}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
