"""
Proofs for the #327 main-grid promotion of the modelled worst balanced Beta
points.

  1. The promotion is exactly the six points of promote_worst_points.POINTS,
     each as Beta_Cumulative and Beta_Survival in the balanced regime, with
     certified references at the binary64 arguments.
  2. All twelve rows are in the main grid once, with the generated reference,
     claim and regime.
  3. In the grid the rows are either all pending export or all observed. Once
     observed, every row is within the 2E-13 balanced Beta guarantee and
     equals the value beta_route_model.py predicts, bit for bit: the Excel
     confirmation #327 asks for.
  4. NEGATIVE: an observation one ulp away from the model's prediction, or one
     above the guarantee, is refused.

Run: python3 test_promote_worst_points.py   (exit 0 = pass)
"""
from decimal import Decimal, localcontext
from pathlib import Path
import csv
import math
import sys

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(HERE.parent))
import promote_worst_points as P  # noqa: E402
from _contract_eval import parse_observed  # noqa: E402
from _grid_safety import canonical_key  # noqa: E402

GRID = HERE.parent / "probability_accuracy_grid.csv"
GUARANTEE = Decimal("2E-13")


def confirms(row, predicted):
    """True when an observed row equals the model and meets the guarantee."""
    with localcontext() as context:
        context.prec = 80
        observed = parse_observed(row.get("observed_vba", ""))
        if observed is None:
            return False
        ref = Decimal(row["reference"])
        within = abs(observed - ref) / abs(ref) <= GUARANTEE
        return within and float(observed) == predicted


def main():
    fails = []

    def check(condition, message):
        if not condition:
            fails.append(message)

    rows = P.build_rows()
    expected = [(f, x, a, b) for x, a, b in P.POINTS
                for f in ("Beta_Cumulative", "Beta_Survival")]
    check([(r["function"], r["arg1"], r["arg2"], r["arg3"]) for r in rows] == expected,
          "1: the promotion is the six points, CDF and survival each")
    check(all(r["regime"] == "balanced" and r["metric"] == "rel" for r in rows),
          "1: every row is a balanced relative row")
    check(all(min(float(r["arg2"]), float(r["arg3"])) / max(float(r["arg2"]), float(r["arg3"]))
              >= 0.1 and min(float(r["arg2"]), float(r["arg3"])) < 100 for r in rows),
          "1: every point is in the balanced regime, not large_shape")

    with open(GRID, newline="", encoding="utf-8") as f:
        grid = list(csv.DictReader(f))
    by_key = {}
    for r in grid:
        by_key.setdefault(canonical_key(r), []).append(r)
    found = []
    for r in rows:
        hits = by_key.get(canonical_key(r), [])
        check(len(hits) == 1, f"2: {r['function']}{(r['arg1'], r['arg2'], r['arg3'])} "
                              f"appears {len(hits)} times in the main grid")
        if len(hits) == 1:
            g = hits[0]
            check((g["reference"], g["claim"], g["regime"], g["evidence_set"])
                  == (r["reference"], r["claim"], r["regime"], r["evidence_set"]),
                  f"2: {r['function']}({r['arg1']}, ...) differs from the generated row")
            found.append(g)

    observed = [g for g in found if g["observed_vba"].strip()]
    check(len(observed) in (0, len(found)),
          f"3: rows are all pending or all observed, got {len(observed)} of {len(found)}")
    for g in observed:
        predicted = P.modelled(g["function"], g["arg1"], g["arg2"], g["arg3"])
        check(confirms(g, predicted),
              f"3: {g['function']}({g['arg1']}, {g['arg2']}, {g['arg3']}) observed "
              f"{g['observed_vba']} does not equal the model {predicted!r} within 2E-13")

    # 4. Negative controls on a synthetic row.
    worst = next(r for r in rows if r["function"] == "Beta_Survival"
                 and r["arg2"] == "50.0" and r["arg1"] == "0.410058")
    predicted = P.modelled(worst["function"], worst["arg1"], worst["arg2"], worst["arg3"])
    exact = dict(worst, observed_vba=repr(predicted))
    check(confirms(exact, predicted), "4: the model's own value confirms")
    off = dict(worst, observed_vba=repr(math.nextafter(predicted, 1.0)))
    check(not confirms(off, predicted), "4: one ulp from the model is refused")
    far = dict(worst, observed_vba=format(Decimal(worst["reference"]) * Decimal("1.0000000000005"), "E"))
    check(not confirms(far, float(Decimal(far["observed_vba"]))),
          "4: an observation above 2E-13 is refused even when it matches the prediction")

    if fails:
        print(f"FAIL: {len(fails)} balanced Beta promotion check(s)")
        for m in fails:
            print("  - " + m)
        raise SystemExit(1)
    state = "observed and confirmed" if observed else "pending export"
    print(f"PASS: the six modelled worst balanced Beta points are in the main grid as "
          f"{len(found)} rows ({state}); off-model and above-guarantee observations "
          "are refused")


if __name__ == "__main__":
    main()
