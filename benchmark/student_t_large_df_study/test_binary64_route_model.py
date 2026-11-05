"""
Fixture for finding F1 (PRODUCTION_FEASIBILITY.md): the adopted #34 route is
not implementable with the existing VBA normal-tail kernel.

  0. Source binding: the model is bound to the VBA procedures and constants
     it mirrors (model_fidelity.py, source_binding.json). A change to any of
     them, or to the model, fails this fixture until the finding is
     re-measured and the binding re-recorded.
  1. Fidelity: the binary64 port of PROB_NormalUpperTailPositive reproduces
     every committed Excel observation of NormalStandard_Survival and
     NormalStandard_Cumulative to within 2 true binary64 ULPs (ordered-integer
     distance), 9 of the 12 exactly and the others 1, 1 and 2 ULPs apart.
  2. The order-5 series itself is sound in binary64: with a correctly rounded
     Q(t) the route meets the preregistered bound 2^-50 at every forward
     fitting point inside the adopted region (df >= 1E5, |t| <= 8).
  3. With the existing kernel the route is worse than the current continued
     fraction - violating the preregistered improvement requirement (section
     8) - at exactly the points recorded in PRODUCTION_FEASIBILITY.md, and is
     at least 100 times worse at t = 6, df = 1E6.
  4. The kernel's own relative error grows as numerical_limitations.csv
     records for SurvivalTailRel: above 1E-14 by z = 3, above 1E-11 by z = 5.

If the normal-tail kernel or the adopted predicate changes, this fixture fails
until the finding is re-measured and the document updated. The kernel is
checked by its bound source text, not only by the committed observations,
whose largest |z| is 3.

Run: python3 test_binary64_route_model.py   (exit 0 = pass, nonzero = fail)
"""
import csv
import os
import sys
from decimal import Decimal

import mpmath as mp

HERE = os.path.dirname(os.path.abspath(__file__))
BENCHMARK = os.path.dirname(HERE)
sys.path.insert(0, HERE)
import binary64_route_model as M                                 # noqa: E402
import model_fidelity as F                                       # noqa: E402

fails = []


def check(cond, msg):
    if not cond:
        fails.append(msg)


# 0. The model is the one bound to today's VBA source.
for problem in F.stale(F.load_record()):
    fails.append("source binding: " + problem)

# 1. Port fidelity against Excel.
kernels = {"NormalStandard_Survival": M.vba_survival, "NormalStandard_Cumulative": M.vba_cdf}
seen = exact = 0
apart = []
for grid in ("probability_accuracy_grid.csv", os.path.join("holdout", "holdout_grid.csv")):
    with open(os.path.join(BENCHMARK, grid), newline="", encoding="utf-8") as f:
        for row in csv.DictReader(f):
            kernel = kernels.get(row["function"])
            raw = (row.get("observed_vba") or "").strip()
            if kernel is None or not raw or "ERR" in raw.upper():
                continue
            observed = float(sum(Decimal(part) for part in raw.split(";")))
            modelled = kernel(float(row["arg1"]))
            seen += 1
            exact += observed == modelled
            ulps = F.ulp_distance(observed, modelled)
            apart += [ulps] if ulps else []
            check(F.within_ulps(observed, modelled),
                  f"{row['function']}({row['arg1']}): port differs from Excel by {ulps} ULPs "
                  f"(limit {F.FIDELITY_LIMIT_ULPS})")
check(seen >= 12, f"expected at least 12 committed normal observations, found {seen}")
check(exact * 4 >= seen * 3, f"only {exact}/{seen} observations reproduced exactly")
check(sorted(apart) == [1, 1, 2], f"ULP distances of the inexact observations are {sorted(apart)}, "
                                  "not the [1, 1, 2] that PRODUCTION_FEASIBILITY.md records")

# 2-3. Route accuracy inside the adopted region.
rows = M.feasibility()
check(len(rows) == 36, f"expected 36 fitting points inside the region, got {len(rows)}")
check(all(r["route_exact_q"] <= M.BOUND for r in rows),
      "with an exact Q the order-5 route meets 2^-50 at every fitting point")
worse = sorted((r["t"], r["df"]) for r in rows
               if r["route_vba_kernel"] > max(r["current_cf"], M.BOUND))
expected = [(5.0, 1e5), (5.0, 1e6), (5.0, 1e7), (6.0, 1e5), (6.0, 1e6), (6.0, 1e7), (6.0, 1e8)]
check(worse == expected, f"points where the existing-kernel route is worse than today: {worse}")
at = {(r["t"], r["df"]): r for r in rows}
check(at[(6.0, 1e6)]["route_vba_kernel"] >= 100 * at[(6.0, 1e6)]["current_cf"],
      "t = 6, df = 1E6: the existing-kernel route is at least 100x worse than today")
check(at[(2.0, 1e8)]["route_vba_kernel"] < at[(2.0, 1e8)]["current_cf"] / 1e5,
      "t = 2, df = 1E8 (the #34 counterexample): the route improves it by over 1E5")

# 4. The kernel's documented degradation, measured directly.
def kernel_error(z):
    with mp.workdps(60):
        q = mp.erfc(mp.mpf(z) / mp.sqrt(2)) / 2
        return float(abs((mp.mpf(M.vba_upper_tail(z)) - q) / q))


check(kernel_error(2.0) < 2e-15, "kernel is accurate near the body (z = 2)")
check(kernel_error(3.0) > 1e-14, "kernel error exceeds 1E-14 by z = 3")
check(kernel_error(5.0) > 1e-11, "kernel error exceeds 1E-11 by z = 5")
check(kernel_error(8.0) < 1e-15, "the far-tail continued fraction (z >= 7.07) is accurate")

if fails:
    print("FAIL: #34 production feasibility (finding F1)")
    for f in fails:
        print("  - " + f)
    raise SystemExit(1)
print(f"PASS: #34 production feasibility (finding F1): port matches Excel on {seen} normal "
      f"observations ({exact} exactly, the others {sorted(apart)} true ULPs apart, limit "
      f"{F.FIDELITY_LIMIT_ULPS}); source binding current; exact-Q route meets 2^-50 at all "
      f"36 fitting points; existing-kernel route worse than today at {len(worse)} points")
