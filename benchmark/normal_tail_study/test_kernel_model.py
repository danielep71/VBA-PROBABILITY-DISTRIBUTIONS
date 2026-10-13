"""
Fixture for the #72 normal-tail kernel study (PREREGISTRATION.md §8).

  1. The coefficient file matches the preregistered design, and every hi + lo
     literal pair is bit-identical to its target Double (A3).
  2. results.json is the recorded result for this coefficient file. The study
     FAILED acceptance condition A2 and nothing else: exactly one coefficient,
     the numerator's z^7 term, is negative (about -2.5E-10). The holdout H was
     therefore not evaluated, and the independent holdout was not read.
  3. A subsample of A5 is recomputed: the candidate kernel's Q, 1 - Q and
     density meet 2^-50 on part of the validation set and at the split seam.
  4. A7 is recomputed: the #34 order-5 route on the candidate kernel meets
     2^-50 at the 36 fitting points of its region.
  5. Model fidelity: on every committed Excel observation of the direct
     surfaces in the main and survival_boundary grids, the model of today's
     code reproduces the observation within 4 ulps.

If the coefficients, the kernel model or the recorded result change, this
fixture fails until the study is re-run.

Run: python3 test_kernel_model.py   (exit 0 = pass, nonzero = fail)
"""
import hashlib
import json
import os
import sys

import mpmath as mp

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import kernel_model as km                                        # noqa: E402
import validate_kernel as vk                                     # noqa: E402

fails = []


def check(cond, msg):
    if not cond:
        fails.append(msg)


B = 2.0 ** -50

# 1. Design and coefficient properties --------------------------------------
raw = km.COEFFICIENTS["raw"]
fit = raw["fit"]
check(fit["interval"] == [0, 3] and fit["degree"] == [7, 7], "fit interval or degree is not the preregistered [0, 3], (7, 7)")
check((fit["nodes"], fit["iterations"], fit["dps"]) == (500, 120, 60), "fit settings differ from the preregistration")
check((raw["split"], raw["continued_fraction_levels"], raw["exponent_split_scale"]) == (3.0, 60, 256),
      "split, continued-fraction depth or exponent scale differs from the preregistration")
check(len(raw["numerator"]) == 8 and len(raw["denominator"]) == 8, "expected 8 + 8 coefficients")
check(mp.mpf(raw["denominator"][0]["value"]) == 1, "D(0) must be 1")
non_positive = [(name, i) for name in ("numerator", "denominator")
                for i, r in enumerate(raw[name]) if mp.mpf(r["value"]) <= 0]
check(non_positive == [("numerator", 7)], f"non-positive coefficients changed: {non_positive}")
check(-1e-9 < mp.mpf(raw["numerator"][7]["value"]) < 0, "the z^7 numerator coefficient left (-1E-9, 0)")
with mp.workdps(40):
    p = [mp.mpf(r["value"]) for r in raw["numerator"]]
    condition = max(sum(abs(c) * mp.mpf(k) ** j / 1000 ** j for j, c in enumerate(p))
                    / sum(c * mp.mpf(k) ** j / 1000 ** j for j, c in enumerate(p)) for k in range(0, 3001, 10))
check(condition - 1 < 2.0 ** -20, f"numerator cancellation on [0, 3] is {condition - 1}")
for r in raw["numerator"] + raw["denominator"] + [raw["inv_sqrt_two_pi"]]:
    x = float(r["double"])
    check(float(r["hi"]) + float(r["lo"]) == x, f"literal pair {r['hi']} + {r['lo']} is not {r['double']} (A3)")
    check(float(mp.mpf(r["value"])) == x, f"{r['double']} is not the nearest Double to {r['value']}")
    for lit in (r["hi"], r["lo"]):
        mantissa = lit.lstrip("-").split("e")[0].replace(".", "").lstrip("0")
        check(len(mantissa) <= 15, f"literal {lit} has more than 15 significant digits")

# 2. Recorded result --------------------------------------------------------
with open(os.path.join(HERE, "results.json"), encoding="utf-8") as f:
    res = json.load(f)
with open(km.COEFFICIENTS_PATH, "rb") as f:
    check(res["coefficients_sha256"] == hashlib.sha256(f.read()).hexdigest(),
          "results.json was not produced from the committed kernel_coefficients.json")
check(res["status"] == "FAIL" and res["problems"] == ["A2 failed"],
      f"recorded status {res['status']}: {res['problems']}")
check(res["acceptance"]["A2"]["pass"] is False, "A2 not recorded as failed")
for name in ("A1", "A3", "A4", "A5", "A7", "A8"):
    check(res["acceptance"].get(name, {}).get("pass") is True, f"{name} not recorded as passed")
check("A6" not in res["acceptance"], "the holdout must not be evaluated after a failed condition")
check(res["holdout_559_inspected"] is False, "the independent holdout must not be read")
check(res["oracle"]["stability_min_digits"] >= 50, "oracle stability below 50 digits")
check(res["oracle"]["quadrature_min_digits"] >= 40, "quadrature agreement below 40 digits")


# 3. A5 subsample -----------------------------------------------------------
def q_ref(z):
    with mp.workdps(60):
        return mp.erfc(mp.mpf(z) / mp.sqrt(2)) / 2


def pdf_ref(z):
    with mp.workdps(60):
        return mp.exp(-mp.mpf(z) ** 2 / 2) / mp.sqrt(2 * mp.pi)


def rel(v, r):
    with mp.workdps(60):
        return float(abs((mp.mpf(v) - r) / r))


grid, uniform = vk.validation_set()
sample = grid[::40] + uniform[::50] + vk.around(3.0)
worst, measured = 0.0, 0
for z in sample:
    q = q_ref(z)
    if q < 2.0 ** -1022:
        continue
    measured += 1
    worst = max(worst, rel(km.candidate_upper_tail(z), q), rel(km.candidate_pdf(z), pdf_ref(z)),
                rel(km.cdf(z), 1 - q))
check(worst <= B, f"candidate kernel exceeds 2^-50 on the A5 subsample: {worst:.3e}")
check(measured >= 450, f"A5 subsample measured only {measured} points")

# 4. A7 on the fitting points ----------------------------------------------
route_rows = res["acceptance"]["A7"]["fitting"]["rows"]
check(len(route_rows) == 36, f"expected 36 route fitting points, recorded {len(route_rows)}")
for r in route_rows:
    e = rel(vk.route_candidate(r["t"], r["df"]), vk.rm.oracle_tail(r["t"], r["df"]))
    check(e <= B, f"#34 route on the candidate kernel exceeds 2^-50 at t={r['t']}, df={r['df']}: {e:.3e}")

# 5. Fidelity of the model of today's code ----------------------------------
direct = [r for r in res["predictions"]["rows"]
          if r["prediction"] == "P1" and r["grid"] != "holdout_older/holdout_older_grid.csv"]
check(len(direct) >= 30, f"too few direct rows recorded: {len(direct)}")
for r in direct:
    u = r["model_vs_committed_ulps"]
    check(u is not None and u <= 4, f"model of today's code is {u} ulps from Excel at {r['function']} {r['args']}")

if fails:
    print("FAIL: #72 normal-tail kernel study fixture")
    for msg in fails:
        print("  -", msg)
    sys.exit(1)
print(f"PASS: #72 normal-tail kernel: design and literals pinned; recorded result FAIL on A2 only "
      f"(z^7 numerator coefficient {float(raw['numerator'][7]['value']):.2e}, cancellation {float(condition - 1):.1e}); "
      f"A5 subsample worst {worst:.2e}; #34 route meets 2^-50 at 36 points; "
      f"today's code reproduced within 4 ulps at {len(direct)} Excel observations")
