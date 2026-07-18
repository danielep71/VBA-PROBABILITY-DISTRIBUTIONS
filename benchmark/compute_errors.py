#!/usr/bin/env python3
"""
compute_errors.py
============================================================================
Phase 3 of the accuracy harness. Reads probability_accuracy_grid.csv AFTER the
observed_vba column has been filled by the Excel export macro, computes the
absolute and relative error of each observed value against the mpmath
reference, locates the maximum error per function, checks each function against
its published claim, and writes accuracy_summary.md.

Rows whose observed_vba cell is still empty are reported as "not measured" and
excluded from the pass/fail check, so a partial run degrades honestly.

Usage:
    python compute_errors.py
    python compute_errors.py --grid probability_accuracy_grid.csv --out accuracy_summary.md
"""
import argparse
import csv
import datetime as _dt
import math
import re


def parse_claim(claim):
    # e.g. "rel<6.1E-14" or "abs<=3E-17" -> (metric, threshold)
    m = re.match(r"(rel|abs)\s*<?=?\s*([0-9.eE+-]+)", claim)
    if not m:
        return None, None
    return m.group(1), float(m.group(2))


def errors(observed, reference, metric):
    o, r = float(observed), float(reference)
    abs_e = abs(o - r)
    rel_e = abs_e / abs(r) if r != 0 else (0.0 if o == 0 else math.inf)
    return abs_e, (abs_e if metric == "abs" else rel_e)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--grid", default="probability_accuracy_grid.csv")
    ap.add_argument("--out", default="accuracy_summary.md")
    args = ap.parse_args()

    rows = list(csv.DictReader(open(args.grid)))
    by_fn = {}
    for r in rows:
        by_fn.setdefault(r["function"], []).append(r)

    lines = []
    lines.append("# Accuracy summary")
    lines.append("")
    lines.append(f"Generated {_dt.date.today().isoformat()} by `compute_errors.py` "
                 f"from `{args.grid}`.")
    lines.append("")
    lines.append("Reference values are mpmath at 50 digits (see "
                 "`generate_reference_values.py`). Observed values are produced by the "
                 "VBA library via `M_STATS_PROBDIST_ACCURACY_EXPORT.bas`. Each function is "
                 "checked against the accuracy claim published in its source comment.")
    lines.append("")
    lines.append("| Function | Claim | Metric | Max error | At input | Points | Verdict |")
    lines.append("|---|---|---|---:|---|---:|---|")

    all_pass = True
    any_measured = False
    for fn in sorted(by_fn):
        grp = by_fn[fn]
        claim = grp[0]["claim"]
        cmetric, threshold = parse_claim(claim)
        measured = [r for r in grp if r["observed_vba"].strip() != ""]
        n_meas = len(measured)
        if n_meas == 0:
            lines.append(f"| {fn} | {claim} | {cmetric or ''} | — | not measured | "
                         f"0/{len(grp)} | ⏳ pending |")
            continue
        any_measured = True
        worst_e, worst_at = -1.0, ""
        for r in measured:
            try:
                _, e = errors(r["observed_vba"], r["reference"], cmetric)
            except (ValueError, ZeroDivisionError):
                e = math.inf
            if e > worst_e:
                worst_e = e
                args_str = ", ".join(a for a in (r["arg1"], r["arg2"], r["arg3"]) if a)
                worst_at = args_str
        ok = threshold is not None and worst_e <= threshold
        all_pass = all_pass and ok
        verdict = "✅ pass" if ok else "❌ FAIL"
        lines.append(f"| {fn} | {claim} | {cmetric} | {worst_e:.2e} | "
                     f"`{worst_at}` | {n_meas}/{len(grp)} | {verdict} |")

    lines.append("")
    if not any_measured:
        lines.append("> **No observed values present yet.** Run the export macro in Excel to "
                     "fill the `observed_vba` column, then re-run `compute_errors.py`.")
    elif all_pass:
        lines.append("> All measured functions meet their published accuracy claims.")
    else:
        lines.append("> **One or more functions fail their published claim.** "
                     "Investigate the flagged inputs before updating any source comment.")

    with open(args.out, "w") as f:
        f.write("\n".join(lines) + "\n")
    print(f"wrote {args.out} ({sum(1 for r in rows if r['observed_vba'].strip())} "
          f"observed of {len(rows)} rows)")


if __name__ == "__main__":
    main()
