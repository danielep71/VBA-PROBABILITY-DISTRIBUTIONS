"""
Check the §4 predictions of INVERSE_PREREGISTRATION.md against the Excel
export of the inverse polish (#34), and write inverse_wave_check.json.

The predictions are the ones inverse_predictions.json recorded before any VBA
was written. This script reads them, and the Excel observations now in the
grids, and checks each row as §4 states:

  PI  the 30 committed StudentT_InverseCumulative rows the polish reaches: the
      observation is a member of its recorded outcome set, and its relative
      error is at most the recorded bound.
  PU  every other Student-t row, every density row and every F row of the main
      grid, tchi_large_df and the holdout: bit-identical to the observation at
      --baseline (an error stays the same error).
  PH  the untouched holdout rows of §6 (holdout rows absent at --baseline):
      a forward row within 4 Doubles of production_route_model and at most
      2^-50 relative; an inverse row by PI's rule, with the outcome set and the
      bound computed here by predict_inverse.outcomes.

Error is relative error against the committed reference at the row's binary64
arguments. A grid none of whose PI rows changed from --baseline is reported as
not re-exported, never as passing. Any row outside its prediction falsifies
it: the script exits 1, and the result is recorded against #34 before any
contract is adopted (§8).

    python check_inverse_wave.py --baseline <commit before the export>
"""
import argparse
import json
import os
import subprocess
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import check_excel_wave as C  # noqa: E402
import predict_inverse as PI  # noqa: E402
import predict_production as PP  # noqa: E402
import production_route_model as P  # noqa: E402

OUT = os.path.join(HERE, "inverse_wave_check.json")
HOLDOUT = "holdout/holdout_grid.csv"
NEW_HOLDOUT_ROWS = 56


def member(record, row):
    value = PP.observed(C.raw(row))
    out = {**record, "observed": C.raw(row)}
    if value is None:
        return {**out, "pass": False, "why": "no numeric observation"}
    allowed = {float(x) for x in record["outcomes"]}
    error = C.rel(value, row["reference"])
    out.update(error=error, member=value in allowed)
    out["pass"] = out["member"] and error <= record["bound"]
    return out


def check(baseline):
    with open(PI.OUT, encoding="utf-8") as f:
        recorded = json.load(f)["predictions"]
    grids = recorded["grids"]
    now = {g: C.current_grid(g) for g in grids}
    base = {g: C.baseline_grid(baseline, g) for g in grids}
    result = {"PI": [], "PU": {}, "PH": [], "not_reexported": []}
    predicted = set()
    for record in recorded["rows"]:
        grid = record["grid"]
        k = C.row_key(record["function"], record["args"], record["regime"])
        predicted.add((grid, k))
        tag = {"grid": grid, "function": record["function"], "args": record["args"],
               "outcomes": record["outcomes"], "bound": record["bound"]}
        row = now[grid].get(k)
        result["PI"].append({**tag, "pass": False, "why": "row missing"} if row is None
                            else member(tag, row))
    for grid in grids:
        rows, changed = 0, []
        for k, row in now[grid].items():
            fn = k[0]
            if not (fn.startswith("StudentT_") or fn.startswith("F_")) or (grid, k) in predicted:
                continue
            old = base[grid].get(k)
            if old is None:
                continue                     # a new holdout row: PH below
            rows += 1
            if not C.same(old, row):
                changed.append({"function": fn, "args": [repr(a) for a in k[1]],
                                "baseline": C.raw(old), "observed": C.raw(row)})
        result["PU"][grid] = {"rows": rows, "changed": changed, "pass": not changed}
        moved = sum(1 for (g, k) in predicted
                    if g == grid and k in base[grid] and not C.same(base[grid][k], now[grid][k]))
        if moved == 0:
            result["not_reexported"].append(grid)
    for k, row in now[HOLDOUT].items():
        fn = k[0]
        if not fn.startswith("StudentT_") or k in base[HOLDOUT]:
            continue
        a, df = k[1]
        tag = {"grid": HOLDOUT, "function": fn, "args": [row["arg1"], row["arg2"]]}
        if fn == "StudentT_InverseCumulative":
            if not PI.reaches_polish(a, df):
                result["PH"].append({**tag, "pass": False, "why": "outside the polish"})
                continue
            qd, found, _, bound = PI.outcomes(a, df)
            sign = 1.0 if a > 0.5 else -1.0
            result["PH"].append(member({**tag, "outcomes": [repr(sign * x) for x in found],
                                        "bound": bound}, row))
        else:
            if not PP.reaches_route(abs(a), df):
                result["PH"].append({**tag, "pass": False, "why": "outside the route"})
                continue
            value = P.survival(a, df) if fn == "StudentT_Survival" else P.cumulative(a, df)
            result["PH"].append(C.forward(tag, row, value, PP.B))
    return result


def summary(result):
    out = {}
    for name in ("PI", "PH"):
        rows = result[name]
        errors = [r["error"] for r in rows if "error" in r]
        out[name] = {"rows": len(rows), "failed": sum(not r["pass"] for r in rows),
                     "worst_error": max(errors) if errors else None}
    out["PU"] = {g: {"rows": v["rows"], "changed": len(v["changed"])}
                 for g, v in result["PU"].items()}
    return out


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    ap.add_argument("--baseline", required=True, help="commit holding the baseline observations")
    a = ap.parse_args(argv)
    sha = subprocess.run(["git", "-C", C.ROOT, "rev-parse", a.baseline], capture_output=True,
                         text=True, check=True).stdout.strip()
    result = check(sha)
    ok = (not result["not_reexported"]
          and all(r["pass"] for n in ("PI", "PH") for r in result[n])
          and all(v["pass"] for v in result["PU"].values())
          and len(result["PI"]) == 30 and len(result["PH"]) == NEW_HOLDOUT_ROWS)
    record = {"issue": "#34", "design": "INVERSE_PREREGISTRATION.md §4", "baseline": sha,
              "summary": summary(result), "pass": ok, **result}
    with open(OUT, "w", encoding="utf-8", newline="\n") as f:
        json.dump(record, f, indent=1)
        f.write("\n")
    print(json.dumps(record["summary"], indent=1))
    if result["not_reexported"]:
        print("NOT RE-EXPORTED: " + ", ".join(result["not_reexported"]))
    for name in ("PI", "PH"):
        for r in result[name]:
            if not r["pass"]:
                shown = {k: v for k, v in r.items() if k != "outcomes"}
                print(f"FAIL {name}: {shown}")
    for grid, v in result["PU"].items():
        for r in v["changed"][:20]:
            print(f"FAIL PU {grid}: {r}")
    print("PASS: every §4 prediction holds" if ok else "FAIL: a §4 prediction is falsified")
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
