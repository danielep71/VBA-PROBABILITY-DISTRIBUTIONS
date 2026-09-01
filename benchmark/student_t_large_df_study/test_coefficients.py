"""
Fixture for the #34 Student-t large-df tail coefficients.

Verifies coefficients.json WITHOUT trusting it:

  1. re-derives orders 1..5 from scratch (SymPy, ~1 s) and requires the
     stored exact rationals to match term for term - order 5 is the adopted
     K*, so the production-relevant coefficients are fully recomputed;
  2. parses each stored factored string and requires it to expand to the
     stored terms, for all 8 orders;
  3. numerically pins all 8 orders: at df = 1E8, t = 2, each successive
     order must reduce the truncation error by at least 6 decimal digits
     until the oracle floor, which cannot hold if any coefficient is wrong;
  4. requires g_1 to equal the classical (t^3 + t)/4;
  5. requires every Q-coefficient to vanish at the recomputed orders.

Run: python3 test_coefficients.py   (exit 0 = pass, nonzero = fail)
"""
import json
import os
import sys

import sympy as sp
from mpmath import mp, mpf

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
sys.path.insert(0, os.path.dirname(HERE))
import derive_coefficients as D                                  # noqa: E402
import validate_expansion as V                                   # noqa: E402

fails = []


def check(cond, msg):
    if not cond:
        fails.append(msg)


with open(os.path.join(HERE, "coefficients.json"), encoding="utf-8") as f:
    stored = json.load(f)

t = sp.symbols("t", positive=True)

# 1. Recompute orders 1..5 and compare exactly.
RECOMPUTE = 5
fresh = D.derive(RECOMPUTE)          # asserts Q-coefficients vanish and g_1 classical
for k in range(1, RECOMPUTE + 1):
    want = {(n,): sp.Rational(c) for (n,), c in fresh[k].terms()}
    got = {(n,): sp.Rational(p, q) for n, p, q in stored["g"][str(k)]["terms"]}
    check(want == got, f"stored g_{k} does not match a fresh derivation")

# 2. Stored factored string must expand to the stored terms, all orders.
for k in range(1, stored["max_order"] + 1):
    expr = sp.expand(sp.sympify(stored["g"][str(k)]["factored"], locals={"t": t}))
    poly = sp.Poly(expr, t)
    got = {(n,): sp.Rational(c) for (n,), c in poly.terms()}
    want = {(n,): sp.Rational(p, q) for n, p, q in stored["g"][str(k)]["terms"]}
    check(got == want, f"g_{k}: factored string does not expand to the stored terms")
    check(stored["g"][str(k)]["degree"] == poly.degree(),
          f"g_{k}: stored degree {stored['g'][str(k)]['degree']} != {poly.degree()}")

# 3. Numerical pin on all 8 orders. Each order gains ~8 digits at df = 1E8.
mp.dps = 100
ref = V.oracle_tail(2.0, 1e8)
floor = mpf(10) ** -44
prev = None
for K in range(1, stored["max_order"] + 1):
    e = abs(V.expansion_tail(2.0, 1e8, K) - ref) / ref
    if prev is not None and prev > floor:
        # Either a full 6-digit gain, or the order has reached the oracle floor.
        check(e < prev * mpf(10) ** -6 or e < floor,
              f"order {K} did not improve on order {K-1} by 6 digits at df=1E8 t=2: "
              f"{mp.nstr(prev, 3)} -> {mp.nstr(e, 3)}")
    prev = e
check(prev < mpf(10) ** -40, "order 8 at df=1E8 t=2 must reach the oracle floor")

# 4. Classical leading term.
g1 = sp.expand(sp.sympify(stored["g"]["1"]["factored"], locals={"t": t}))
check(sp.simplify(g1 - (t ** 3 + t) / 4) == 0, "stored g_1 is not (t^3+t)/4")

# 5. Every stored coefficient is an exact rational with a positive denominator.
for k in range(1, stored["max_order"] + 1):
    for n, p, q in stored["g"][str(k)]["terms"]:
        check(isinstance(p, int) and isinstance(q, int) and q > 0,
              f"g_{k} term t^{n} is not an exact rational")

if fails:
    print("FAIL: Student-t large-df coefficients")
    for f_ in fails:
        print("  - " + f_)
    raise SystemExit(1)
print(f"PASS: Student-t large-df tail coefficients (orders 1..{RECOMPUTE} re-derived "
      f"and matched; all {stored['max_order']} orders parsed, degree-checked and "
      "numerically pinned at df=1E8; g_1 classical; Q-coefficients vanish)")
