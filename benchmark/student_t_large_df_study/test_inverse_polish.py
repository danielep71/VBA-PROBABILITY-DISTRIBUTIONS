"""
CI fixture for the inverse preregistration of the Student-t large-df route
(#34, INVERSE_PREREGISTRATION.md).

  1. The polish predicate: it reaches df >= 1E5 (up to the 1E8 cap) with a
     target tail of at most 0.4, so p = 0.6 is in and the Double after it is
     out; p = 0.5 and df just below 1E5 are out.
  2. The polish arithmetic: at a correctly rounded quantile the polish stays
     within a few Doubles, a start 16 Doubles away is pulled back inside the
     outcome set, and a start above t = 8 is returned unchanged.
  3. inverse_predictions.json reproduces from the model and the grids, so a
     change to the model, the route model or an affected grid row fails until
     the predictions are deliberately re-recorded.
  4. The recorded conditions hold: every validation point and every PI row has
     its whole outcome set within its bound.

Run: python3 test_inverse_polish.py   (exit 0 = pass)
"""
import json
import math
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import predict_inverse as PI  # noqa: E402

fails = []


def check(cond, msg):
    if not cond:
        fails.append(msg)


# 1. Predicate
check(PI.tail_of(0.6) == 0.4 and PI.reaches_polish(0.6, 1e6), "p = 0.6 must reach the polish")
check(not PI.reaches_polish(math.nextafter(0.6, 0.0), 1e6), "the Double below p = 0.6 must not")
check(PI.reaches_polish(0.4, 1e6) and not PI.reaches_polish(math.nextafter(0.4, 1.0), 1e6),
      "the lower tail mirrors the upper")
check(not PI.reaches_polish(0.5, 1e6), "p = 0.5 must not reach the polish")
check(PI.reaches_polish(0.9, 1e5) and not PI.reaches_polish(0.9, math.nextafter(1e5, 0.0)),
      "df = 1E5 is in, the Double below it is out")
check(PI.reaches_polish(0.9, 1e8) and not PI.reaches_polish(0.9, math.nextafter(1e8, 2e8)),
      "df = 1E8 is in, the Double above it is out")

# 2. Arithmetic
q, found, worst, bound = PI.outcomes(0.999, 1e6)
check(worst <= bound, "p = 0.999, df = 1E6: an outcome exceeds the bound")
check(PI.polish(q, PI.tail_of(0.999), 1e6) in found, "a correctly rounded start leaves the outcome set")
check(PI.polish(q + 16 * math.ulp(q), PI.tail_of(0.999), 1e6) in found,
      "a start 16 Doubles away is not pulled into the outcome set")
check(PI.polish(8.5, 1e-16, 1e6) == 8.5, "a start above t = 8 must be returned unchanged")

# 3. Reproduction
if PI.main(["--check"]) != 0:
    fails.append("inverse_predictions.json does not reproduce")

# 4. Recorded conditions
with open(PI.OUT, encoding="utf-8") as f:
    recorded = json.load(f)
check(recorded["validation"]["within_bound"], "a validation point exceeds its bound")
check(recorded["predictions"]["PI_within_bound"], "a PI row exceeds its bound")
check(not recorded["new_holdout_read"], "the new holdout must not be read")

if fails:
    print("FAIL: #34 inverse polish fixture")
    for f in fails:
        print("  - " + f)
    sys.exit(1)
print(f"PASS: #34 inverse polish: predicate, arithmetic, {recorded['predictions']['PI']} PI rows "
      f"and {recorded['validation']['points']} validation points within bound; predictions reproduce")
