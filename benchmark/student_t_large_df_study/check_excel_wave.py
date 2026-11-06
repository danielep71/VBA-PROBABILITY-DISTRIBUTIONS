"""
Check the §5 predictions of PRODUCTION_PREREGISTRATION.md against the Excel
export of the large-df route (#34), and write excel_wave_check.json.

The predictions are the ones production_predictions.json recorded before any
VBA was written. This script reads them, and the Excel observations now in the
grids, and checks each row as §5 states:

  PT1  StudentT_Survival / StudentT_Cumulative rows the route reaches: within
       4 ulps of the predicted Double, and relative error <= 2^-50.
  PT2  StudentT_InverseCumulative rows whose quantile the route reaches:
       relative error <= the row's recorded bound.
  PT3  every other Student-t row and every F row of the main grid and of
       tchi_large_df: bit-identical to the observation at --baseline (an
       error stays the same error).
  PT4  the untouched holdout rows of §6 (StudentT_* rows of
       holdout/holdout_grid.csv): the PT1 and PT2 rules, with the predicted
       Double and the bound computed here by the same model functions, since
       production_predictions.json does not read the holdout.

Error is relative error against the committed reference at the row's binary64
arguments. A grid none of whose PT1 rows changed from --baseline is reported
as not re-exported, never as passing. Any row outside its prediction
falsifies it: the script exits 1, and the result is recorded against #34
before any contract is adopted (§8).

    python check_excel_wave.py --baseline <commit before the export>
"""
import argparse
import csv
import io
import json
import os
import struct
import subprocess
import sys

import mpmath as mp

HERE = os.path.dirname(os.path.abspath(__file__))
BENCH = os.path.dirname(HERE)
ROOT = os.path.dirname(BENCH)
sys.path.insert(0, HERE)
import predict_production as PP  # noqa: E402
import production_route_model as P  # noqa: E402

OUT = os.path.join(HERE, "excel_wave_check.json")
PREDICTIONS = os.path.join(HERE, "production_predictions.json")
HOLDOUT = "holdout/holdout_grid.csv"
ARGS = ("arg1", "arg2", "arg3", "arg4")


def row_key(fn, args, regime):
    return fn, tuple(float(a) for a in args if a is not None and str(a).strip()), regime


def read_rows(text):
    rows = {}
    for row in csv.DictReader(io.StringIO(text)):
        rows[row_key(row["function"], [row.get(a) for a in ARGS], row.get("regime", ""))] = row
    return rows


def baseline_grid(ref, grid):
    proc = subprocess.run(["git", "-C", ROOT, "show", f"{ref}:benchmark/{grid}"],
                          capture_output=True, text=True, encoding="utf-8")
    if proc.returncode != 0:
        raise SystemExit(f"cannot read benchmark/{grid} at {ref}: {proc.stderr.strip()}")
    return read_rows(proc.stdout.lstrip("﻿"))


def current_grid(grid):
    with open(os.path.join(BENCH, grid), encoding="utf-8-sig") as f:
        return read_rows(f.read())


def raw(row):
    return (row.get("observed_vba") or "").strip()


def ordered(x):
    bits = struct.unpack("<q", struct.pack("<d", x))[0]
    return bits if bits >= 0 else -(bits & 0x7FFFFFFFFFFFFFFF)


def ulps(a, b):
    return abs(ordered(a) - ordered(b))


def rel(value, reference):
    with mp.workdps(60):
        reference = mp.mpf(reference)
        return float(abs((mp.mpf(value) - reference) / reference))


def same(base, now):
    a, b = PP.observed(raw(base)), PP.observed(raw(now))
    if a is None or b is None:
        return raw(base) == raw(now)
    return struct.pack("<d", a) == struct.pack("<d", b)


def forward(record, row, predicted, bound):
    value = PP.observed(raw(row))
    out = {**record, "observed": raw(row), "predicted": repr(predicted), "bound": bound}
    if value is None:
        return {**out, "pass": False, "why": "no numeric observation"}
    distance, error = ulps(value, predicted), rel(value, row["reference"])
    out.update(ulps=distance, error=error)
    out["pass"] = distance <= PP.ULPS and error <= bound
    return out


def inverse(record, row, bound):
    value = PP.observed(raw(row))
    out = {**record, "observed": raw(row), "bound": bound}
    if value is None:
        return {**out, "pass": False, "why": "no numeric observation"}
    error = rel(value, row["reference"])
    return {**out, "error": error, "pass": error <= bound}


def check(baseline):
    with open(PREDICTIONS, encoding="utf-8") as f:
        recorded = json.load(f)["predictions"]
    grids = recorded["grids"]
    now = {g: current_grid(g) for g in grids}
    base = {g: baseline_grid(baseline, g) for g in grids}
    result = {"PT1": [], "PT2": [], "PT3": {}, "PT4": [], "not_reexported": []}
    predicted = set()
    for record in recorded["rows"]:
        grid, fn = record["grid"], record["function"]
        k = row_key(fn, record["args"], record["regime"])
        predicted.add((grid, k))
        row = now[grid].get(k)
        tag = {"grid": grid, "function": fn, "args": record["args"]}
        if row is None:
            result[record["prediction"]].append({**tag, "pass": False, "why": "row missing"})
        elif record["prediction"] == "PT1":
            result["PT1"].append(forward(tag, row, float(record["predicted"]), record["bound"]))
        else:
            result["PT2"].append(inverse(tag, row, record["bound"]))
    for grid in grids:
        changed, rows, broken = 0, 0, []
        for k, row in now[grid].items():
            fn = k[0]
            if not (fn.startswith("StudentT_") or fn.startswith("F_")) or (grid, k) in predicted:
                continue
            rows += 1
            old = base[grid].get(k)
            if old is None or not same(old, row):
                broken.append({"function": fn, "args": [repr(a) for a in k[1]],
                               "baseline": raw(old) if old else None, "observed": raw(row)})
        result["PT3"][grid] = {"rows": rows, "changed": broken, "pass": not broken}
        for record in recorded["rows"]:
            if record["grid"] == grid and record["prediction"] == "PT1":
                k = row_key(record["function"], record["args"], record["regime"])
                if k in base[grid] and not same(base[grid][k], now[grid][k]):
                    changed += 1
        if changed == 0:
            result["not_reexported"].append(grid)
    holdout_base = baseline_grid(baseline, HOLDOUT)
    for k, row in current_grid(HOLDOUT).items():
        fn = k[0]
        if not fn.startswith("StudentT_") or k in holdout_base and raw(holdout_base[k]):
            continue
        tag = {"grid": HOLDOUT, "function": fn, "args": [row["arg1"], row["arg2"]]}
        a, df = k[1]
        if fn == "StudentT_InverseCumulative":
            q = float(row["reference"])
            if not PP.reaches_route(abs(q), df):
                result["PT4"].append({**tag, "pass": False, "why": "quantile outside the route"})
                continue
            result["PT4"].append(inverse(tag, row, PP.quantile_bound(a, q, df)))
        else:
            if not PP.reaches_route(abs(a), df):
                result["PT4"].append({**tag, "pass": False, "why": "argument outside the route"})
                continue
            value = P.survival(a, df) if fn == "StudentT_Survival" else P.cumulative(a, df)
            result["PT4"].append(forward(tag, row, value, PP.B))
    return result


def summary(result):
    out = {}
    for name in ("PT1", "PT2", "PT4"):
        rows = result[name]
        errors = [r["error"] for r in rows if "error" in r]
        out[name] = {"rows": len(rows), "failed": sum(not r["pass"] for r in rows),
                     "worst_error": max(errors) if errors else None,
                     "worst_ulps": max((r["ulps"] for r in rows if "ulps" in r), default=None)}
    out["PT3"] = {g: {"rows": v["rows"], "changed": len(v["changed"])}
                  for g, v in result["PT3"].items()}
    return out


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    ap.add_argument("--baseline", required=True, help="commit holding the baseline observations")
    a = ap.parse_args(argv)
    sha = subprocess.run(["git", "-C", ROOT, "rev-parse", a.baseline], capture_output=True,
                         text=True, check=True).stdout.strip()
    result = check(sha)
    ok = (not result["not_reexported"]
          and all(r["pass"] for n in ("PT1", "PT2", "PT4") for r in result[n])
          and all(v["pass"] for v in result["PT3"].values())
          and len(result["PT4"]) == 44)
    record = {"issue": "#34", "design": "PRODUCTION_PREREGISTRATION.md §5",
              "baseline": sha, "summary": summary(result), "pass": ok, **result}
    with open(OUT, "w", encoding="utf-8", newline="\n") as f:
        json.dump(record, f, indent=1)
        f.write("\n")
    print(json.dumps(record["summary"], indent=1))
    if result["not_reexported"]:
        print("NOT RE-EXPORTED: " + ", ".join(result["not_reexported"]))
    for name in ("PT1", "PT2", "PT4"):
        for r in result[name]:
            if not r["pass"]:
                print(f"FAIL {name}: {r}")
    for grid, v in result["PT3"].items():
        for r in v["changed"][:20]:
            print(f"FAIL PT3 {grid}: {r}")
    print("PASS: every §5 prediction holds" if ok else "FAIL: a §5 prediction is falsified")
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
