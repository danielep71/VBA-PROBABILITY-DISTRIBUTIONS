"""
#14: materialize the holdout of the Gamma inverse at the final scale.

Rows of gamma_inverse_scale_grid.csv, by kind:

  underflowed  the unit-scale quantile is below the smallest normal Double,
               down to about exp(-1415); the scale brings the final quantile
               back into the normal range. This is the defect of #14.
  log_region   the unit-scale quantile is between the smallest normal Double
               and exp(-40.5): the rest of the region the rule answers in log
               space. Half at scale 1, half at a random scale.
  seam         for ten shapes, thirteen adjacent probabilities around the
               crossover L0 = -40: the last Double the model answers in log
               space, and its neighbours at 1, 2, 4, 8, 16 and 32 units in
               the last place on either side. Scale 1.
  chisquare    scale exactly 2, where ChiSquare_InverseCumulative(p, 2A) is
               the same call: unit quantiles in [MIN_NORMAL / 2, MIN_NORMAL),
               in the log region, and just above the crossover.
  limits       a final quantile below every Double (the exact zero is
               expected), a subnormal one, and, at shape 1 where the
               quantile is -ln(1 - p) times the scale, one just inside the
               Double range and four above it (#NUM! is expected).

The seeded points come from random.Random(SEED) and are never evaluated by
the model or by Excel before the verdict. The seam is located from the
model's predicate alone (inv_scale_model.uses_log), by bisection on the bit
pattern of the probability; the tests exercise the same construction at
other shapes and with another seed.

Columns (10): point_id, kind, p_bits, a_bits, s_bits, ref_x, gamma_inv,
chi_inv, gamma_cdf, log_unit. The first six are the design and its
reference; M_STATS_PROBDIST_GIS.bas writes the rest. The inputs are the 16
hexadecimal digits of the Doubles, so no decimal is parsed in Excel. ref_x
is the quantile at the final scale to 40 digits, or ZERO, or NUM.

Run:
    python3 holdout_build.py --write   (only when the grid is absent)
    python3 holdout_build.py --check   (the committed grid is the design)
"""
import argparse
import csv
import io
import math
import os
import random
import struct
import sys

from mpmath import mp, mpf

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import inv_scale_model as M  # noqa: E402

GRID = os.path.join(HERE, "gamma_inverse_scale_grid.csv")
FIELDS = ["point_id", "kind", "p_bits", "a_bits", "s_bits", "ref_x", "gamma_inv", "chi_inv",
          "gamma_cdf", "log_unit"]
KEY_COLUMNS = 6
SEED = 14
LOG_MIN_NORMAL = math.log(M.MIN_NORMAL)
SEAM_SHAPES = (1e-4, 1e-3, 0.01, 0.02, 0.1, 0.5, 1.0, 2.5, 7.0, 15.0)
SEAM_OFFSETS = (-32, -16, -8, -4, -2, -1, 0, 1, 2, 4, 8, 16, 32)
DBL_MAX = mpf(2) ** 1024 * (1 - mpf(2) ** -53)
HALF_MIN_SUBNORMAL = mpf(2) ** -1075


def bits(x):
    return f"{struct.unpack('<Q', struct.pack('<d', x))[0]:016X}"


def from_bits(text):
    return struct.unpack("<d", struct.pack("<Q", int(text, 16)))[0]


def step(x, n):
    return struct.unpack("<d", struct.pack("<Q", struct.unpack("<Q", struct.pack("<d", x))[0] + n))[0]


def reference(p, a, s):
    """ref_x. Shape 1 is the exponential distribution, in closed form; any
    other shape needs a unit quantile small enough for inv_scale_model.exact."""
    if a == 1.0:
        with mp.workdps(60):
            x = -mp.log1p(-mpf(p)) * mpf(s)
    else:
        x = M.exact(p, a, s)
    if x > DBL_MAX:
        return "NUM"
    if x < HALF_MIN_SUBNORMAL:
        return "ZERO"
    with mp.workdps(60):
        return mp.nstr(x, 40)


def probability(rng, a, lo, hi):
    """A normal probability below one whose true unit log-quantile is in (lo, hi)."""
    for _ in range(200):
        target = rng.uniform(lo, hi)
        log_p = target * a - M.log_gamma1p(a)
        if not math.log(M.MIN_NORMAL) < log_p < -1e-12:
            continue
        p = math.exp(log_p)
        if M.MIN_NORMAL <= p < 1.0 and lo < M.log_unit(p, a) < hi:
            return p
    return None


def scale_for(rng, p, a, lo, hi):
    """A scale that puts the log of the final quantile in (lo, hi)."""
    unit = M.log_unit(p, a)
    low, high = max(lo - unit, -700.0), min(hi - unit, 709.0)
    if low >= high:
        return None
    return math.exp(rng.uniform(low, high))


def sample(rng, count, shape_range, unit_range, scale=None, final_range=(-700.0, 700.0)):
    out = []
    while len(out) < count:
        a = 10 ** rng.uniform(*shape_range)
        p = probability(rng, a, *unit_range)
        if p is None:
            continue
        s = scale if scale is not None else scale_for(rng, p, a, *final_range)
        if s is None:
            continue
        out.append((p, a, s))
    return out


def seam(a):
    """Thirteen probabilities around the model's crossover at shape a."""
    low = math.exp(-41.0 * a - M.log_gamma1p(a))      # log side
    high = math.exp(-39.0 * a - M.log_gamma1p(a))     # ordinary side
    assert M.uses_log(low, a) and not M.uses_log(high, a)
    lo = struct.unpack("<Q", struct.pack("<d", low))[0]
    hi = struct.unpack("<Q", struct.pack("<d", high))[0]
    while hi - lo > 1:
        mid = (lo + hi) // 2
        if M.uses_log(struct.unpack("<d", struct.pack("<Q", mid))[0], a):
            lo = mid
        else:
            hi = mid
    last_log = struct.unpack("<d", struct.pack("<Q", lo))[0]
    return [step(last_log, n) for n in SEAM_OFFSETS]


def build(seed=SEED, seam_shapes=SEAM_SHAPES):
    rng = random.Random(seed)
    rows = []

    def add(kind, points):
        for i, (p, a, s) in enumerate(points):
            rows.append({"point_id": f"{kind[:2]}{i:03d}", "kind": kind, "p_bits": bits(p),
                         "a_bits": bits(a), "s_bits": bits(s), "ref_x": reference(p, a, s)})

    # The counterexample of the parent issue, then the seeded points
    add("underflowed", [(1e-4, 0.01, 1e100)]
        + sample(rng, 119, (-8.0, 0.02), (-1415.0, LOG_MIN_NORMAL - 0.01)))
    add("log_region", sample(rng, 60, (-8.0, 1.2), (LOG_MIN_NORMAL + 0.01, -40.5), scale=1.0)
        + sample(rng, 60, (-8.0, 1.2), (LOG_MIN_NORMAL + 0.01, -40.5)))
    add("seam", [(p, a, 1.0) for a in seam_shapes for p in seam(a)])
    add("chisquare",
        sample(rng, 20, (-6.0, 0.0), (LOG_MIN_NORMAL - math.log(2.0) + 0.001, LOG_MIN_NORMAL - 0.001),
               scale=2.0)
        + sample(rng, 20, (-6.0, 1.0), (LOG_MIN_NORMAL + 0.01, -40.5), scale=2.0)
        + sample(rng, 20, (-6.0, 1.0), (-39.5, -20.0), scale=2.0))
    add("limits",
        sample(rng, 10, (-8.0, 0.0), (-1415.0, -800.0), final_range=(-1400.0, -760.0))
        + sample(rng, 10, (-8.0, 0.0), (-1415.0, -800.0), final_range=(-743.0, -709.5))
        + [(0.5, 1.0, 1e308)] + [(p, 1.0, 1.5e308) for p in (0.7, 0.9, 0.99, 0.9999)])
    return rows


def render(rows):
    buf = io.StringIO()
    writer = csv.DictWriter(buf, fieldnames=FIELDS, lineterminator="\n", restval="")
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
            print("FAIL: the holdout's identity columns differ from a fresh build")
            return 1
        print(f"PASS: {len(identity(text)) - 1} rows match the frozen design")
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
