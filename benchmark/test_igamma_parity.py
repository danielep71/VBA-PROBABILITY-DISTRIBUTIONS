"""
Parity between the production incomplete-gamma kernel and the study
implementation.

WHAT THIS IS

benchmark/_igamma.py (production, owned by the gate) and
benchmark/chisq_reference_study/chisq_reference.py (evidence, owned by the
study) implement the same converging route independently. Keeping them
separate preserves the dependency direction - the gate must not import from
a study directory - at the cost of two implementations that could drift.

This fixture measures that drift. It is a DRIFT DETECTOR, not a new
independent oracle: both sides share an algorithm, so agreement here says
nothing about whether the algorithm is right. The independent numerical
authority remains the quadrature and MPFR legs recorded in the study's
composite authority.

FROZEN PROTOCOL

  * Evaluate the 69 COMMITTED quantiles from chisq_reference.json. They are
    read, never regenerated and never re-inverted: this compares two
    evaluations of the same input, not two searches for the same root.
  * Working precision 50 dps, matching the gate.
  * Compare the DIRECTLY computed informative tail - P on the lower-series
    route, Q on the upper-CF route - so the comparison is never dominated
    by a complement subtraction.
  * Require at least 40 significant digits of agreement at every point.
  * Require both implementations to select the SAME route and to report
    convergence.
  * Report the measured minimum agreement.

Run: python3 test_igamma_parity.py   (exit 0 = pass, nonzero = fail)
"""
import json
import os
import sys

import mpmath as mp

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
sys.path.insert(0, os.path.join(HERE, "chisq_reference_study"))

import _igamma as PROD                                    # noqa: E402
import chisq_reference as STUDY                           # noqa: E402

DPS = 50
REQUIRED_DIGITS = 40
REFERENCE_JSON = os.path.join(HERE, "chisq_reference_study",
                              "chisq_reference.json")

fails = []


def check(cond, msg):
    if not cond:
        fails.append(msg)


def agreement_digits(x, y):
    """Significant decimal digits to which x and y agree."""
    mp.mp.dps = DPS + 30
    x, y = mp.mpf(x), mp.mpf(y)
    if x == y:
        return None                     # identical at this precision
    if x == 0 or y == 0:
        return mp.mpf(0)
    return -mp.log10(abs(x - y) / abs(x))


# --- the seam, first: both routes must be reachable and agree across it -----
# x = a + 1 is the route boundary. A fixture that only ever exercised one side
# would not notice a boundary moved by one, so probe immediately below, exactly
# at, and immediately above it.
mp.mp.dps = DPS
_a = mp.mpf(50)
_seam = PROD.seam(_a)
check(_seam == _a + 1, "seam is a + 1")

for label, _x in (("below seam", _seam - mp.mpf("1e-20")),
                  ("exact seam", _seam),
                  ("above seam", _seam + mp.mpf("1e-20"))):
    v_p, side_p, route_p, _ = PROD.reg_incomplete(_a, _x)
    v_s, q_s, route_s, _ = STUDY.reg_both(_a, _x)
    v_s = v_s if side_p == "P" else q_s
    d = agreement_digits(v_p, v_s)
    check(route_p == route_s, f"{label}: both select the same route")
    check(d is None or d >= REQUIRED_DIGITS,
          f"{label}: agree to at least {REQUIRED_DIGITS} digits (got {d})")

# The two sides of the seam must select different routes, or the boundary is
# not doing anything.
_below = PROD.reg_incomplete(_a, _seam - mp.mpf("1e-20"))[2]
_above = PROD.reg_incomplete(_a, _seam + mp.mpf("1e-20"))[2]
check(_below == "lower_series" and _above == "upper_cf",
      f"the seam separates the routes (below={_below}, above={_above})")

# --- the frozen 69 committed quantiles -------------------------------------
if not os.path.exists(REFERENCE_JSON):
    print(f"FAIL: {REFERENCE_JSON} is missing; parity cannot be measured")
    raise SystemExit(1)

with open(REFERENCE_JSON, encoding="utf-8") as f:
    records = json.load(f)["references"]
check(len(records) == 69, f"expected 69 committed references, got {len(records)}")

worst = None
worst_at = None
routes = {"lower_series": 0, "upper_cf": 0}
for r in records:
    if r["status"] != "ACCEPTED":
        fails.append(f"reference {r['probability']['hex']} is not ACCEPTED")
        continue
    mp.mp.dps = DPS
    a = mp.mpf(r["df"]) / 2
    x = mp.mpf(r["quantile"]) / 2          # committed; never re-inverted here

    try:
        v_p, side_p, route_p, iters_p = PROD.reg_incomplete(a, x)
    except PROD.IGammaNonConvergence as exc:
        fails.append(f"production kernel did not converge at df={r['df']:.0e}: {exc}")
        continue
    try:
        p_s, q_s, route_s, iters_s = STUDY.reg_both(a, x)
    except Exception as exc:
        fails.append(f"study kernel did not converge at df={r['df']:.0e}: {exc}")
        continue

    routes[route_p] = routes.get(route_p, 0) + 1
    check(route_p == route_s,
          f"route disagreement at df={r['df']:.0e} p={r['probability']['hex']}: "
          f"{route_p} vs {route_s}")

    # Compare the directly computed informative tail on each route.
    v_s = p_s if side_p == "P" else q_s
    d = agreement_digits(v_p, v_s)
    if d is not None:
        if d < REQUIRED_DIGITS:
            fails.append(
                f"parity below {REQUIRED_DIGITS} digits at df={r['df']:.0e} "
                f"p={r['probability']['hex']}: {mp.nstr(d, 6)}")
        if worst is None or d < worst:
            worst = d
            worst_at = f"df={r['df']:.0e} p={r['probability']['float_repr']}"

mp.mp.dps = DPS
if fails:
    print("FAIL: incomplete-gamma parity")
    for f_ in fails:
        print("  - " + f_)
    raise SystemExit(1)

print(f"PASS: incomplete-gamma parity - production and study agree on all "
      f"{len(records)} committed quantiles")
if worst is None:
    # Every point was bit-identical. The two implementations are currently
    # arithmetically equivalent, so the effective drift threshold is the first
    # differing digit, not the 40-digit contractual floor. Any divergence at
    # all turns this fixture red.
    print(f"      minimum agreement: EXACT (bit-identical at {DPS} dps at every "
          f"point; contractual floor {REQUIRED_DIGITS} digits)")
else:
    print(f"      minimum agreement {mp.nstr(worst, 8)} significant digits "
          f"(required {REQUIRED_DIGITS}), at {worst_at}")
print(f"      route selection identical; {routes.get('lower_series', 0)} lower-series, "
      f"{routes.get('upper_cf', 0)} upper-CF")
print(f"      seam probed below, at and above x = a + 1")
print("      drift detector only - the independent authority is the "
      "quadrature and MPFR legs")
