"""
Proofs for the #24 main-grid promotion of the X = 0.5 survival ladder.

  1. The promotion is exactly the eleven frozen x05_* shapes of
     gamma_q_complement.csv below 0.5, each as Gamma_Survival(0.5, A, 1) and
     the exact doubling ChiSquare_Survival(1, 2A), sharing one reference.
  2. Each reference is the upper tail Q(A, 0.5) computed directly.
  3. All 22 rows are in the main grid once, with the generated reference, claim
     and metric, and no study observation was copied into them.
  4. The 5E-15 bound passes an exact value and refuses a 1E-14 error and any
     unmeasured value. In the grid the rows are either all pending export or all
     observed, and every observed row meets 5E-15, including the Chi-square rows
     whose registry contract is broader.

Run: python3 test_promote_x05_ladder.py   (exit 0 = pass)
"""
from decimal import Decimal, localcontext
from pathlib import Path
import csv
import io
import subprocess
import sys

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(HERE.parent))
import promote_x05_ladder as P  # noqa: E402
from _grid_safety import canonical_key  # noqa: E402

import mpmath as mp  # noqa: E402

FROZEN_SHAPES = [1e-06, 3e-06, 1e-05, 3e-05, 0.0001, 0.0003, 0.001, 0.003,
                 0.01, 0.03, 0.1]
GRID = HERE.parent / "probability_accuracy_grid.csv"


def main():
    fails = []

    def check(condition, message):
        if not condition:
            fails.append(message)

    # 1. The frozen ladder, both routes, exact doubling.
    check(P.ladder_shapes() == FROZEN_SHAPES, f"ladder shapes changed: {P.ladder_shapes()}")
    rows = P.build_rows()
    check(len(rows) == 22, f"{len(rows)} rows, not 22")
    for g, c in zip(rows[0::2], rows[1::2]):
        check(g["function"] == "Gamma_Survival" and c["function"] == "ChiSquare_Survival",
              "routes are not paired Gamma, Chi-square")
        check(float(g["arg1"]) == 0.5 and g["arg3"] == "1.0", f"Gamma row is not X = 0.5, scale 1: {g}")
        check(float(c["arg1"]) == 2 * float(g["arg1"]) and float(c["arg2"]) == 2 * float(g["arg2"])
              and c["arg3"] == "", f"Chi-square row is not the exact doubling of {g['arg2']}")
        check(c["reference"] == g["reference"], "paired routes have different references")
    check(all(r["claim"] == "rel<=5E-15" and r["metric"] == "rel" and r["regime"] == "all"
              and r["evidence_set"] == "main grid" and not r["observed_vba"] for r in rows),
          "promoted rows must claim rel<=5E-15 with blank observations")

    # 2. Each reference is Q(A, 0.5) computed directly, never 1 - P.
    with mp.workdps(60):
        for g, shape in zip(rows[0::2], FROZEN_SHAPES):
            direct = mp.gammainc(shape, 0.5, mp.inf, regularized=True)
            check(abs(mp.mpf(g["reference"]) / direct - 1) < mp.mpf("1E-50"),
                  f"shape {shape}: reference is not the direct upper tail")

    # 3. Permanent main-grid membership, without copied observations.
    with open(GRID, newline="", encoding="utf-8") as f:
        grid = list(csv.DictReader(f))
    by = {}
    for g in grid:
        by.setdefault(canonical_key(g), []).append(g)
    promoted = []
    for r in rows:
        matches = by.get(canonical_key(r), [])
        check(len(matches) == 1, f"{r['function']} {r['arg1']}, {r['arg2']}: "
                                 f"{len(matches)} main-grid rows, not 1")
        if len(matches) == 1:
            m = matches[0]
            check(all(m[k] == r[k] for k in ("reference", "claim", "metric")),
                  f"{r['function']} {r['arg2']}: main-grid reference, claim or metric changed")
            promoted.append(m)
    # The pre-#24 ladder export (#23 wave, e278318). A later re-export of the
    # ladder legitimately returns the same Doubles as the promoted rows.
    pre = subprocess.check_output(["git", "show", "e278318:benchmark/positive_ratio_subnormal/"
                                   "gamma_q_complement.csv"], cwd=HERE, text=True)
    study_obs = {r["survival"] for r in csv.DictReader(io.StringIO(pre))}
    check(not any(m["observed_vba"] in study_obs for m in promoted),
          "a study observation was copied into the main grid")

    # 4. The 5E-15 bound.
    sample = rows[8]
    with localcontext() as context:
        context.prec = 80
        degraded = str(Decimal(sample["reference"]) * (1 + Decimal("1E-14")))
    check(P.meets_frozen_bound(dict(sample, observed_vba=sample["reference"])),
          "an exact value must pass")
    check(not P.meets_frozen_bound(dict(sample, observed_vba=degraded)), "1E-14 must fail")
    for invalid in ("", "ERROR", "NaN", "Infinity", "not numeric"):
        check(not P.meets_frozen_bound(dict(sample, observed_vba=invalid)),
              f"unmeasured value must fail: {invalid!r}")
    observed = [m for m in promoted if m["observed_vba"].strip()]
    check(len(observed) in (0, len(promoted)), "the promoted rows are partly observed")
    for m in observed:
        check(P.meets_frozen_bound(m), f"{m['function']} {m['arg1']}, {m['arg2']}: "
                                       f"observation breaches 5E-15")

    if fails:
        print("FAIL: #24 X = 0.5 main-grid promotion")
        for f in fails:
            print("  - " + f)
        return 1
    state = "observed, all within 5E-15" if observed else "pending export"
    print(f"PASS: #24 X = 0.5 ladder promoted as 22 main-grid rows ({state})")
    return 0


if __name__ == "__main__":
    sys.exit(main())
