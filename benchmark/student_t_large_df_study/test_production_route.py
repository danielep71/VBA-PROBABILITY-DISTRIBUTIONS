"""
CI fixture for the production preregistration of the Student-t large-df route
(#34, PRODUCTION_PREREGISTRATION.md).

  1. Every coefficient of the route is the Double nearest its exact rational in
     coefficients.json, and its recorded hi + lo literal pair rebuilds exactly
     that Double.
  2. production_predictions.json reproduces from the model and the grids, so a
     change to the model, the coefficients, the normal kernel model or an
     affected grid row fails until the predictions are deliberately re-recorded.
  3. The recorded measurement meets the preregistered conditions: A1 (worst
     error on V within 2^-50) and A2 (no study fitting point where the route is
     worse than the current continued fraction).

Run: python3 test_production_route.py   (exit 0 = pass)
"""
import json
import os
import sys
from fractions import Fraction

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import predict_production as PP  # noqa: E402
import production_route_model as P  # noqa: E402

fails = []


def check(cond, msg):
    if not cond:
        fails.append(msg)


with open(os.path.join(HERE, "coefficients.json"), encoding="utf-8") as f:
    g = json.load(f)["g"]
with open(PP.OUT, encoding="utf-8") as f:
    recorded = json.load(f)

# 1. Coefficients and literals
for k, row in enumerate(P.COEFFICIENTS, start=1):
    exact_terms = {p: Fraction(n, d) for p, n, d in g[str(k)]["terms"]}
    check(len(row) == (max(exact_terms) - 1) // 2 + 1, f"g_{k}: wrong number of U powers")
    for j, (exact, double, hi, lo) in enumerate(row):
        check(exact == exact_terms.get(2 * j + 1, Fraction(0)), f"g_{k} U^{j}: not the exact rational")
        check(double == float(exact), f"g_{k} U^{j}: not the nearest Double")
        check(float(hi) + float(lo) == double, f"g_{k} U^{j}: the pair does not rebuild the Double")
        check(len(hi.lstrip("-").replace(".", "").lstrip("0").split("e")[0]) <= 15,
              f"g_{k} U^{j}: hi literal has more than 15 significant digits")
check(recorded["route"]["literals"] == [[[hi, lo] for _, _, hi, lo in row] for row in P.COEFFICIENTS],
      "the recorded literal pairs differ from the model's")
check((recorded["route"]["order"], recorded["route"]["df_min"], recorded["route"]["t_max"])
      == (5, 1e5, 8.0), "the route is not the study's adopted (5, 1E5, 8)")

# 2. Reproduction
if PP.main(["--check"]) != 0:
    fails.append("production_predictions.json does not reproduce")

# 3. Preregistered conditions
m, i = recorded["measurement"], recorded["improvement"]
check(m["within_bound"] and m["worst"] <= P.BOUND, f"A1: worst {m['worst']:.3e} exceeds 2^-50")
check(m["worst_at_100_digits"] <= P.BOUND, "A1: the worst point exceeds 2^-50 at 100 digits")
check(m["points"] >= 1600, f"A1: only {m['points']} validation points")
check(not i["route_worse_than_current"], f"A2: route worse at {i['route_worse_than_current']}")
check(i["points"] == 36, f"A2: expected the 36 fitting points in the region, got {i['points']}")
p = recorded["predictions"]
check((p["PT1"], p["PT2"]) == (58, 18), f"expected 58 PT1 and 18 PT2 rows, got {p['PT1']}, {p['PT2']}")
check(recorded["holdout_read"] is False, "the independent holdout must not be read")

if fails:
    print("FAIL: #34 production route preregistration")
    for msg in fails:
        print("  - " + msg)
    sys.exit(1)
print(f"PASS: #34 production route: {sum(len(r) for r in P.COEFFICIENTS)} coefficient literals exact; "
      f"V worst {m['worst']:.2e} within 2^-50 on {m['points']} points; no study fitting point worse "
      f"than the current route; {p['PT1']} PT1 and {p['PT2']} PT2 predictions reproduce")
