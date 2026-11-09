"""
Score the #23 Excel wave against the frozen design.

Part A, the study grid (gamma_series_small_shape_grid.csv; HOLDOUT_DESIGN.md):
  every row must echo exactly the binary64 inputs its reference was evaluated
  at, return a number, and meet 3E-15 relative. The Gamma and Chi-square routes
  of each holdout point reach the same kernel P(s, z), so their results are
  also compared bit for bit.

Part B, the fitting ladder (positive_ratio_subnormal/gamma_q_complement.csv)
  against its pre-fix export at BASELINE, rule by rule of MOVEMENT_MANIFEST.md:
  B1  x05_s0..x05_s10: series_p and cdf move, and both meet 3E-15;
  B2  x05_s11 (Shape = 0.5, a half-integer): reported; the post-selection
      explanation keeps exact half-integers on the unchanged table path;
  B3  xsub_s0..xsub_s10: the worst cdf error does not exceed the pre-fix worst;
  B4  x2_s0..x2_s11: public cdf and survival do not move;
  and, reported but not required (#24 owns it): the x05 survival error.

B4 was falsified by the 2026-11-07 wave: the continued-fraction survival moved
at every off-grid shape. PROB_TryGammaContinuedFractionQ forms its result with
the same PROB_TryGammaPrefactor as the series, so it consumes the corrected
off-grid PROB_StirlingError, the mechanism of the manifest's post-selection
explanation. The frozen rule is still evaluated and reported as falsified; the
manifest requires such movement to be explained, and B4X checks the
explanation's testable consequences:
  B4X public cdf at X = 2 does not move; the half-integer control x2_s11 does
      not move; only off-grid shapes move; and no moved survival error grows.

    python analyze_gamma_series_small_shape.py           print the result
    python analyze_gamma_series_small_shape.py --write   also write wave_results.json
    python analyze_gamma_series_small_shape.py --check   fail if wave_results.json
                                                         differs from a recomputation

Exit status: 0 when every requirement holds; 1 otherwise, or when the grid has
no observations yet.
"""
import argparse
import csv
import io
import json
import os
import subprocess
import sys

import mpmath as mp

HERE = os.path.dirname(os.path.abspath(__file__))
BENCH = os.path.dirname(HERE)
ROOT = os.path.dirname(BENCH)
GRID = os.path.join(HERE, "gamma_series_small_shape_grid.csv")
LADDER_PATH = "benchmark/positive_ratio_subnormal/gamma_q_complement.csv"
BASELINE = "5c884bc48e748e323b39dbee418f3aecb5d14a29"   # the pre-fix export (2026-08-14)
# The post-fix export this analysis scored: the #23 wave (run 42365758589).
# Later re-exports, such as #24's, are not the #23 observations.
AFTER = "e278318"
OUT = os.path.join(HERE, "wave_results.json")
THRESHOLD = 3e-15


def value(token):
    """A 'hi;lo' observation as the Double it encodes, or None for ERROR/blank."""
    token = token.strip()
    if not token or token in ("ERROR", "CVERR"):
        return None
    hi, lo = token.split(";")
    return float(hi) + float(lo)


def rel(observed, reference):
    with mp.workdps(60):
        reference = mp.mpf(reference)
        return float(abs((mp.mpf(observed) - reference) / reference))


def lower_p(shape, x):
    with mp.workdps(80):
        return mp.gammainc(mp.mpf(shape), 0, mp.mpf(x), regularized=True)


def upper_q(shape, x):
    with mp.workdps(80):
        return mp.gammainc(mp.mpf(shape), mp.mpf(x), mp.inf, regularized=True)


def part_a(path=GRID):
    with open(path, encoding="utf-8-sig") as f:
        rows = list(csv.DictReader(f))
    if not any(r["observed_vba"].strip() for r in rows):
        return None
    results, problems = [], []
    for r in rows:
        a1, a2 = float(r["arg1"]), float(r["arg2"])
        label = f"{r['function']}({r['arg1']}, {r['arg2']})"
        echo = (value(r["echo_arg1"]), value(r["echo_arg2"]))
        if echo != (a1, a2):
            problems.append(f"{label}: echoed inputs {echo} differ from the reference inputs")
        observed = value(r["observed_vba"])
        error = None if observed is None else rel(observed, r["reference"])
        if error is None:
            problems.append(f"{label}: no numeric result ({r['observed_vba']})")
        elif error > THRESHOLD:
            problems.append(f"{label}: relative error {error:.3e} exceeds {THRESHOLD:.0e}")
        results.append({"function": r["function"], "arg1": r["arg1"], "arg2": r["arg2"],
                        "regime": r["regime"], "observed": r["observed_vba"],
                        "rel_error": error})
    # Route agreement: Gamma(z, s, 1) and ChiSquare(2z, 2s) reach P(s, z).
    gamma = {(float(r["arg1"]), float(r["arg2"])): r["observed_vba"] for r in rows
             if r["function"] == "Gamma_Cumulative"}
    pairs = identical = 0
    for r in rows:
        if r["function"] == "ChiSquare_Cumulative":
            key = (float(r["arg1"]) / 2, float(r["arg2"]) / 2)
            if key in gamma:
                pairs += 1
                identical += value(gamma[key]) == value(r["observed_vba"])
    by_regime = {}
    for item in results:
        group = by_regime.setdefault(item["regime"], {"rows": 0, "worst": 0.0, "worst_at": None})
        group["rows"] += 1
        if item["rel_error"] is not None and item["rel_error"] >= group["worst"]:
            group["worst"], group["worst_at"] = item["rel_error"], \
                f"{item['function']}({item['arg1']}, {item['arg2']})"
    return {"rows": len(rows), "threshold": THRESHOLD, "by_regime": by_regime,
            "route_pairs": pairs, "route_pairs_bit_identical": identical,
            "problems": problems, "pass": not problems, "rows_detail": results}


def ladder(text):
    out = {}
    for r in csv.DictReader(io.StringIO(text)):
        out[r["point_id"]] = r
    return out


def git_show(ref, path):
    return subprocess.run(["git", "-C", ROOT, "show", f"{ref}:{path}"], check=True,
                          capture_output=True, text=True).stdout


def part_b(after_text=None, before_text=None):
    if after_text is None:
        after_text = git_show(AFTER, LADDER_PATH)
    if before_text is None:
        before_text = git_show(BASELINE, LADDER_PATH)
    after, before = ladder(after_text), ladder(before_text)
    if sorted(after) != sorted(before):
        return {"pass": False, "problems": ["the ladder's point set changed"]}
    problems, rows = [], {}
    for pid, new in after.items():
        old = before[pid]
        shape, x = value(new["shape"]), value(new["x"])
        entry = {"shape": shape, "x": x, "branch": new["branch"],
                 "status_before": {s: old[s + "_status"] for s in ("series_p", "cdf", "survival")},
                 "status_after": {s: new[s + "_status"] for s in ("series_p", "cdf", "survival")}}
        for surface in ("series_p", "cdf", "survival"):
            entry[surface + "_moved"] = new[surface] != old[surface]
        p_ref, q_ref = lower_p(shape, x), upper_q(shape, x)
        for when, row in (("before", old), ("after", new)):
            series, cdf, surv = value(row["series_p"]), value(row["cdf"]), value(row["survival"])
            entry[f"series_p_error_{when}"] = None if series is None else rel(series, p_ref)
            entry[f"cdf_error_{when}"] = None if cdf is None else rel(cdf, p_ref)
            entry[f"survival_error_{when}"] = None if surv is None else rel(surv, q_ref)
        rows[pid] = entry

    def ids(prefix, last):
        return [f"{prefix}_s{i}" for i in range(last + 1) if f"{prefix}_s{i}" in rows]

    def missing(rule, pids, surfaces):
        """A rule compares numeric results, so an unsuccessful status or a
        missing value on either side is a failure, never a zero error."""
        for pid in pids:
            for surface in surfaces:
                for when in ("before", "after"):
                    status = rows[pid][f"status_{when}"][surface]
                    if status != "OK" or rows[pid][f"{surface}_error_{when}"] is None:
                        problems.append(f"{rule} {pid}: {surface} {when} the fix has no "
                                        f"numeric result (status {status})")

    b1 = ids("x05", 10)
    for pid in b1:
        e = rows[pid]
        if not (e["series_p_moved"] and e["cdf_moved"]):
            problems.append(f"B1 {pid}: series_p or cdf did not move")
        for surface in ("series_p", "cdf"):
            error = e[f"{surface}_error_after"]
            if error is not None and error > THRESHOLD:
                problems.append(f"B1 {pid}: {surface} error {error:.3e} exceeds {THRESHOLD:.0e}")
    missing("B1", ids("x05", 10), ("series_p", "cdf"))
    b3 = ids("xsub", 10)
    missing("B3", b3, ("cdf",))
    worst_before = max(rows[p]["cdf_error_before"] or 0.0 for p in b3) if b3 else 0.0
    worst_after = max(rows[p]["cdf_error_after"] or 0.0 for p in b3) if b3 else 0.0
    if worst_after > worst_before:
        problems.append(f"B3 xsub: worst cdf error rose from {worst_before:.3e} to {worst_after:.3e}")
    b4 = ids("x2", 11)
    missing("B4X", b4, ("cdf", "survival"))
    b4_moved = [pid for pid in b4 if rows[pid]["cdf_moved"] or rows[pid]["survival_moved"]]
    explanation = []
    for pid in b4:
        e = rows[pid]
        if e["cdf_moved"]:
            explanation.append(f"B4X {pid}: public cdf moved")
        half_integer = (2 * e["shape"]).is_integer()
        if half_integer and e["survival_moved"]:
            explanation.append(f"B4X {pid}: the half-integer shape {e['shape']} moved")
        if e["survival_moved"] and (e["survival_error_after"] or 0) > (e["survival_error_before"] or 0):
            explanation.append(f"B4X {pid}: survival error grew from {e['survival_error_before']:.3e} "
                               f"to {e['survival_error_after']:.3e}")
    if "x2_s11" not in rows or rows["x2_s11"]["survival_moved"]:
        explanation.append("B4X: the half-integer control x2_s11 is missing or moved")
    problems.extend(explanation)
    return {"baseline": BASELINE, "threshold": THRESHOLD,
            "B1_rows": len(b1), "B1_worst_cdf_error": max(rows[p]["cdf_error_after"] or 0 for p in b1),
            "B1_worst_series_p_error": max(rows[p]["series_p_error_after"] or 0 for p in b1),
            "B2_x05_s11": rows.get("x05_s11"),
            "B3_rows": len(b3), "B3_worst_cdf_error_before": worst_before,
            "B3_worst_cdf_error_after": worst_after,
            "B4_rows": len(b4), "B4_moved": len(b4_moved),
            "B4_frozen_prediction": "held" if not b4_moved else "falsified",
            "B4X_explained": not explanation,
            "B4_worst_survival_error_before": max(rows[p]["survival_error_before"] or 0 for p in b4),
            "B4_worst_survival_error_after": max(rows[p]["survival_error_after"] or 0 for p in b4),
            "x05_survival_worst_after": max(rows[p]["survival_error_after"] or 0 for p in b1),
            "x05_survival_worst_before": max(rows[p]["survival_error_before"] or 0 for p in b1),
            "problems": problems, "pass": not problems, "rows": rows}


def analyze():
    a = part_a()
    if a is None:
        return None
    b = part_b()
    return {"issue": "#23", "design": "HOLDOUT_DESIGN.md", "movement": "MOVEMENT_MANIFEST.md",
            "study_grid": a, "fitting_ladder": b, "pass": a["pass"] and b["pass"]}


def summary(result):
    a, b = result["study_grid"], result["fitting_ladder"]
    lines = [f"Part A, study grid: {a['rows']} rows, threshold {a['threshold']:.0e}"]
    for regime, g in a["by_regime"].items():
        lines.append(f"  {regime}: {g['rows']} rows, worst {g['worst']:.3e} at {g['worst_at']}")
    lines.append(f"  Gamma/Chi-square routes bit-identical at {a['route_pairs_bit_identical']} "
                 f"of {a['route_pairs']} pairs")
    lines.append(f"Part B, fitting ladder against {b['baseline'][:7]}:")
    lines.append(f"  B1 x05 s0-s10: {b['B1_rows']} rows, worst cdf error {b['B1_worst_cdf_error']:.3e}, "
                 f"worst series_p error {b['B1_worst_series_p_error']:.3e}")
    s11 = b["B2_x05_s11"]
    if s11:
        lines.append(f"  B2 x05_s11: cdf moved {s11['cdf_moved']}, cdf error {s11['cdf_error_after']:.3e}")
    lines.append(f"  B3 xsub: worst cdf error {b['B3_worst_cdf_error_before']:.3e} before, "
                 f"{b['B3_worst_cdf_error_after']:.3e} after")
    lines.append(f"  B4 x2: {b['B4_moved']} of {b['B4_rows']} rows moved; frozen prediction "
                 f"{b['B4_frozen_prediction'].upper()}; explanation (B4X) "
                 f"{'holds' if b['B4X_explained'] else 'FAILS'}; worst survival error "
                 f"{b['B4_worst_survival_error_before']:.3e} before, "
                 f"{b['B4_worst_survival_error_after']:.3e} after (#26)")
    lines.append(f"  x05 survival (#24, reported only): worst {b['x05_survival_worst_before']:.3e} "
                 f"before, {b['x05_survival_worst_after']:.3e} after")
    for problem in a["problems"] + b["problems"]:
        lines.append("  FAIL " + problem)
    lines.append("RESULT: " + ("PASS" if result["pass"] else "FAIL"))
    return "\n".join(lines)


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    mode = ap.add_mutually_exclusive_group()
    mode.add_argument("--write", action="store_true")
    mode.add_argument("--check", action="store_true")
    a = ap.parse_args(argv)
    result = analyze()
    if result is None:
        print("no observations yet: export the grid first (ci/Export-ExcelGrids.ps1, "
              "export_studies=gss_small_shape)")
        return 1
    fresh = json.loads(json.dumps(result))
    if a.check:
        with open(OUT, encoding="utf-8") as f:
            if json.load(f) != fresh:
                print("FAIL: wave_results.json differs from a recomputation")
                return 1
        print("PASS: wave_results.json reproduces")
    print(summary(result))
    if a.write:
        with open(OUT, "w", encoding="utf-8", newline="\n") as f:
            json.dump(result, f, indent=1)
            f.write("\n")
    return 0 if result["pass"] else 1


if __name__ == "__main__":
    sys.exit(main())
