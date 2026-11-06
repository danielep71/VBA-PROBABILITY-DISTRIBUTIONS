"""
Check the §7 predictions of PREREGISTRATION.md against an Excel export of the
repaired kernel (#72, #81), and write excel_wave_check.json.

The predictions are the ones results.json recorded before the export: the
predicted binary64 value of every P1 and P4 row, and the P2/P3 bound of every
standardised and interval row, in the main grid, survival_boundary and
holdout_older. This script reads them, and the Excel observations now in the
grids, and checks each row as §7 states:

  P1  NormalStandard Survival, Cumulative, Density: within 4 ulps of the
      predicted value, and relative error <= B.
  P2  Normal and Lognormal Cumulative and Survival: error <= the row's bound.
  P3  NormalStandard and Normal IntervalProbability: error <= the row's bound.
  P4  the six inverses: within 4 ulps of the predicted value (8 for
      Lognormal); the contract met unless the row was a predicted breach;
      split_boundary rows <= 1E-14.
  P5  the independent holdout, when exported: the P1-P4 formulas applied to its
      affected rows. The predictions are computed here, by the validator's own
      functions, because the study does not read the holdout (§9).
  unchanged  Normal_Density, Lognormal_Density and
      NormalStandard_InverseCumulativeFast: bit-identical to the observation
      at --baseline (the commit before the export).

Error is relative error against the committed reference at the row's binary64
arguments, as §7 defines it. A grid whose observations were not re-exported
(identical to --baseline for every affected row) is reported as not exported,
never as passing. Any row outside its prediction falsifies it: the script
exits 1, and the result must be recorded against #72 before any contract
changes.

    python check_excel_wave.py --baseline <commit before the export>
"""
import argparse
import csv
import hashlib
import io
import json
import os
import subprocess
import sys
from decimal import Decimal

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import validate_kernel as V  # noqa: E402

BENCH = V.BENCH
ROOT = os.path.dirname(BENCH)
OUT = os.path.join(HERE, "excel_wave_check.json")
HOLDOUT = "holdout/holdout_grid.csv"
ULPS = 4
ULPS_LOGNORMAL = 8
SPLIT_BOUNDARY = 1e-14


def key(fn, args):
    return fn, tuple(float(a) for a in args if a is not None)


def read_grid(text):
    rows = {}
    for row in csv.DictReader(io.StringIO(text)):
        args = [a for a in V.args_of(row) if a is not None]
        rows[key(row["function"], args)] = row
    return rows


def baseline_grid(ref, grid):
    proc = subprocess.run(["git", "-C", ROOT, "show", f"{ref}:benchmark/{grid}"],
                          capture_output=True, text=True, encoding="utf-8")
    if proc.returncode != 0:
        raise SystemExit(f"cannot read benchmark/{grid} at {ref}: {proc.stderr.strip()}")
    return read_grid(proc.stdout)


def current_grid(grid):
    with open(os.path.join(BENCH, grid), encoding="utf-8-sig") as f:
        return read_grid(f.read())


def check_row(rec, row):
    """The outcome of one predicted row against its Excel observation."""
    observed = V.parse_observed(row.get("observed_vba"))
    out = {"grid": rec["grid"], "function": rec["function"], "args": rec["args"],
           "prediction": rec["prediction"], "regime": rec.get("regime")}
    if observed is None:
        out.update(status="FAIL", problems=["no observation"])
        return out
    reference = Decimal(row["reference"].strip())
    error = V.rel(observed, reference)
    out["observed"] = repr(observed)
    out["error"] = error
    problems = []
    if rec["prediction"] in ("P1", "P4"):
        predicted = float(rec["predicted"])
        ulps = V.ulps_apart(observed, predicted)
        limit = ULPS_LOGNORMAL if rec["function"].startswith("Lognormal") else ULPS
        out["ulps_from_predicted"] = ulps
        if ulps > limit:
            problems.append(f"{ulps} ulps from the predicted {predicted!r} (limit {limit})")
    if "bound" in rec:
        out["bound"] = rec["bound"]
        if not error <= rec["bound"]:
            problems.append(f"error {error:.3g} above the bound {rec['bound']:.3g}")
    if rec["prediction"] == "P4":
        contract = rec.get("contract")
        out["contract"] = contract
        if contract is not None and not rec.get("predicted_breach") and not error <= contract:
            problems.append(f"error {error:.3g} above the contract {contract:.3g}")
    out["status"] = "FAIL" if problems else "PASS"
    out["problems"] = problems
    return out


def holdout_predictions():
    """P5: the P1-P4 predictions for the holdout's affected rows.

    The same per-row construction as validate_kernel.predictions(), applied to
    the holdout grid. That function also summarises each prediction class and
    so needs rows of every class, which the holdout does not have."""
    thresholds = V.contract_thresholds()
    rows = []
    with open(os.path.join(BENCH, HOLDOUT), encoding="utf-8-sig") as f:
        for row in csv.DictReader(f):
            fn = row["function"]
            if fn not in V.DIRECT + V.STANDARDISED + V.INTERVALS + V.INVERSES:
                continue
            args = V.args_of(row)
            reference = Decimal(row["reference"].strip())
            new, _ = V.km.public(fn, args, "candidate")
            rec = {"grid": HOLDOUT, "function": fn, "regime": row["regime"],
                   "args": [x for x in args if x is not None],
                   "predicted": repr(new), "predicted_error": V.rel(new, reference)}
            if fn in V.DIRECT:
                rec["prediction"] = "P1"
                rec["bound"] = V.B
            elif fn in V.STANDARDISED:
                rec["prediction"] = "P2"
                z = V.exact_z(fn, args)
                rec["bound"] = V.B + V.standardisation_term(z, reference, V.z_error(fn, args, z))
            elif fn in V.INTERVALS:
                rec["prediction"] = "P3"
                rec["bound"], _ = V.interval_bound(fn, args, reference)
            else:
                rec["prediction"] = "P4"
                limit = thresholds.get((fn, row["regime"]))
                rec["contract"] = limit
                rec["predicted_breach"] = limit is not None and rec["predicted_error"] > limit
                if row["regime"] == "split_boundary":
                    rec["bound"] = SPLIT_BOUNDARY
            rows.append(rec)
    return rows


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    ap.add_argument("--baseline", required=True,
                    help="the commit whose grids hold the observations before the export")
    ap.add_argument("--predictions", default=os.path.join(HERE, "results.json"),
                    help="the results.json that recorded the predictions (default: this folder's)")
    ap.add_argument("--no-holdout", action="store_true", help="skip P5")
    a = ap.parse_args(argv)

    with open(a.predictions, "rb") as f:
        raw = f.read()
    recorded = json.loads(raw)["predictions"]["rows"]
    grids = sorted({r["grid"] for r in recorded})
    if not a.no_holdout:
        holdout = holdout_predictions()
        for r in holdout:
            r["p5"] = True
        recorded = recorded + holdout
        grids.append(HOLDOUT)

    rows, grid_status = [], {}
    for grid in grids:
        now, before = current_grid(grid), baseline_grid(a.baseline, grid)
        mine = [r for r in recorded if r["grid"] == grid]
        changed = sum(1 for r in mine
                      if (now.get(key(r["function"], r["args"])) or {}).get("observed_vba")
                      != (before.get(key(r["function"], r["args"])) or {}).get("observed_vba"))
        unchanged_fns = [k for k in now if k[0] in V.UNCHANGED]
        if mine and changed == 0:
            grid_status[grid] = {"exported": False, "rows": len(mine)}
            continue
        results = []
        for rec in mine:
            row = now.get(key(rec["function"], rec["args"]))
            if row is None:
                results.append({"grid": grid, "function": rec["function"], "args": rec["args"],
                                "prediction": rec["prediction"], "status": "FAIL",
                                "problems": ["row not found in the grid"]})
            else:
                result = check_row(rec, row)
                if rec.get("p5"):
                    result["prediction"] = "P5/" + rec["prediction"]
                results.append(result)
        for k in unchanged_fns:
            old = (before.get(k) or {}).get("observed_vba")
            new = now[k].get("observed_vba")
            results.append({"grid": grid, "function": k[0], "args": list(k[1]),
                            "prediction": "unchanged", "status": "PASS" if old == new else "FAIL",
                            "problems": [] if old == new else [f"observation changed: {old} -> {new}"]})
        rows += results
        grid_status[grid] = {"exported": True, "rows": len(results),
                             "fail": sum(r["status"] == "FAIL" for r in results)}

    summary = {}
    for r in rows:
        s = summary.setdefault(r["prediction"], {"rows": 0, "fail": 0, "worst_error": 0.0,
                                                 "max_ulps_from_predicted": 0})
        s["rows"] += 1
        s["fail"] += r["status"] == "FAIL"
        s["worst_error"] = max(s["worst_error"], r.get("error") or 0.0)
        s["max_ulps_from_predicted"] = max(s["max_ulps_from_predicted"],
                                           r.get("ulps_from_predicted") or 0)
    report = {
        "issue": "#72 / #81 Excel wave: §7 predictions against the exported observations",
        "baseline": a.baseline,
        "predictions_sha256": hashlib.sha256(raw).hexdigest(),
        "grids": grid_status,
        "summary": summary,
        "falsified": [f"{r['prediction']} {r['grid']} {r['function']} {r['args']}: "
                      + "; ".join(r["problems"]) for r in rows if r["status"] == "FAIL"],
        "rows": rows,
    }
    with open(OUT, "w", encoding="utf-8", newline="\n") as f:
        json.dump(report, f, indent=1)
        f.write("\n")
    for grid, status in grid_status.items():
        print(f"{grid}: " + ("not exported" if not status["exported"]
                             else f"{status['rows']} rows, {status['fail']} outside their prediction"))
    for tag, s in sorted(summary.items()):
        print(f"  {tag}: {s['rows']} rows, {s['fail']} fail, worst error {s['worst_error']:.3g}, "
              f"max {s['max_ulps_from_predicted']} ulps from predicted")
    if not any(status["exported"] for status in grid_status.values()):
        print("NOT CHECKED: no grid holds a re-exported observation of an affected row")
        return 1
    if report["falsified"]:
        print("FALSIFIED:\n  - " + "\n  - ".join(report["falsified"]))
        return 1
    print("PASS: every exported affected row is within its preregistered prediction")
    return 0


if __name__ == "__main__":
    sys.exit(main())
