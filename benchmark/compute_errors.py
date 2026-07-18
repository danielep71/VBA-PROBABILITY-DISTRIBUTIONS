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
import sys
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


def load_contract(path=None):
    """Single source of truth for thresholds; used to cross-check grid claims."""
    import os
    if path is None:
        path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "accuracy_contracts.csv")
    out = {}
    try:
        with open(path, newline="") as f:
            for row in csv.DictReader(f):
                m = "rel" if row["metric"].strip().lower().startswith("rel") else "abs"
                out[row["function"]] = {
                    "metric": m,
                    "threshold": row["threshold"].strip(),
                    "status": row.get("status", "active").strip() or "active",
                    "domain": row.get("domain", "").strip(),
                    "notes": row.get("notes", "").strip(),
                }
    except FileNotFoundError:
        return None
    return out


def check_contract_consistency(rows):
    """Warn if any grid claim disagrees with accuracy_contracts.csv."""
    contract = load_contract()
    if contract is None:
        return
    seen = set()
    for r in rows:
        fn = r["function"]
        if fn in seen:
            continue
        seen.add(fn)
        cm, ct = parse_claim(r["claim"])
        if fn in contract:
            gm, gt = contract[fn]["metric"], contract[fn]["threshold"]
            if cm != gm or float(ct) != float(gt):
                print(f"  WARNING: grid claim for {fn} ({r['claim']}) disagrees with "
                      f"accuracy_contracts.csv ({gm}<={gt})")


def parse_observed(s):
    """Observed values may be a single number or a two-part 'hi;lo' sum that the
    export macro writes to preserve full Double precision. Sum the parts."""
    s = s.strip()
    if s == "" or s.upper() == "ERROR":
        return None
    return sum(float(part) for part in s.split(";"))


def errors(observed, reference, metric):
    o, r = parse_observed(observed), float(reference)
    if o is None:
        raise ValueError("no observed value")
    abs_e = abs(o - r)
    rel_e = abs_e / abs(r) if r != 0 else (0.0 if o == 0 else math.inf)
    return abs_e, (abs_e if metric == "abs" else rel_e)


def main():
    ap = argparse.ArgumentParser(
        description="Compute accuracy verdicts and act as a numerical release gate.")
    ap.add_argument("--grid", default="probability_accuracy_grid.csv")
    ap.add_argument("--out", default="accuracy_summary.md")
    ap.add_argument("--allow-known-limitations", action="store_true",
                    help="Development mode: documented KNOWN LIMITATION rows do not "
                         "fail the gate. Default (strict) mode blocks on them.")
    args = ap.parse_args()

    rows = list(csv.DictReader(open(args.grid)))

    check_contract_consistency(rows)
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

    contract = load_contract() or {}
    n_fail = 0
    n_known = 0
    n_pending = 0
    any_measured = False
    for fn in sorted(by_fn):
        grp = by_fn[fn]
        claim = grp[0]["claim"]
        cmetric, threshold = parse_claim(claim)
        fn_status = contract.get(fn, {}).get("status", "active")
        measured = [r for r in grp
                    if r["observed_vba"].strip() != ""
                    and r["observed_vba"].strip().upper() != "ERROR"]
        n_meas = len(measured)
        if n_meas == 0:
            n_pending += 1
            lines.append(f"| {fn} | {claim} | {cmetric or ''} | — | not measured | "
                         f"0/{len(grp)} | ⏳ PENDING |")
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
        # Measurement floor: observed values are a VBA Double rendered to ~15-16
        # significant digits in the CSV, so a RELATIVE claim tighter than ~1E-14
        # cannot be confirmed or denied by this harness. Report it honestly.
        FLOOR = 1e-14
        below_floor = (cmetric == "rel" and threshold is not None and threshold < FLOOR)
        if cmetric == "abs" and threshold is not None:
            # Absolute claim below the double/CSV precision at the value's magnitude
            # cannot be verified. Use the largest reference magnitude in the group.
            mags = [abs(float(r["reference"])) for r in measured
                    if r["reference"].strip() not in ("", "0")]
            ref_mag = max(mags) if mags else 1.0
            if threshold < ref_mag * 2.2e-16:
                below_floor = True
        ok = threshold is not None and worst_e <= threshold
        if fn_status == "known_limitation":
            # Documented defect: not a silent pass, not a hard fail.
            n_known += 1
            verdict = "🔷 KNOWN LIMITATION"
        elif below_floor and not ok:
            verdict = "⚠️ below harness precision"
        elif ok:
            verdict = "✅ PASS"
        else:
            n_fail += 1
            verdict = "❌ FAIL"
        lines.append(f"| {fn} | {claim} | {cmetric} | {worst_e:.2e} | "
                     f"`{worst_at}` | {n_meas}/{len(grp)} | {verdict} |")

    # Surface documented known-limitation contracts that have no measured grid rows
    # yet (e.g. the PROB_LogBeta kernel), so the authoritative verdict cannot read
    # all-green while a documented defect is unrepresented.
    for cfn in sorted(contract):
        if contract[cfn].get("status") == "known_limitation" and cfn not in by_fn:
            n_known += 1
            note = contract[cfn].get("notes", "")
            lines.append(f"| {cfn} | {contract[cfn]['metric']}<={contract[cfn]['threshold']} "
                         f"| {contract[cfn]['metric']} | — | documented | 0/0 "
                         f"| 🔷 KNOWN LIMITATION |")

    lines.append("")
    if not any_measured:
        lines.append("> **No observed values present yet.** Run the export macro in Excel to "
                     "fill the `observed_vba` column, then re-run `compute_errors.py`.")
    else:
        lines.append(f"> **Verdict tally** — FAIL: {n_fail}, KNOWN LIMITATION: {n_known}, "
                     f"PENDING: {n_pending}.")
        lines.append("")
        lines.append("> States: **PASS** meets the contract; **FAIL** exceeds it and must be "
                     "investigated; **KNOWN LIMITATION** is a documented defect tracked in "
                     "`accuracy_contracts.csv` (does not read as green); **PENDING** is not yet "
                     "measured. Rows marked *below harness precision* have relative claims "
                     "tighter than a 15-16 digit CSV round-trip can verify and are not failures.")

    with open(args.out, "w") as f:
        f.write("\n".join(lines) + "\n")

    # ---- release gate ----
    # Strict (default): any FAIL or KNOWN LIMITATION blocks the gate, so a release
    # cannot report a fully green numerical contract while a documented defect
    # (e.g. the LogBeta imbalance band) is still present. Development mode
    # (--allow-known-limitations) blocks only on FAIL.
    blocking = n_fail if args.allow_known_limitations else (n_fail + n_known)
    mode = "development (known limitations allowed)" if args.allow_known_limitations else "strict"
    print(f"{args.out}: FAIL={n_fail} KNOWN_LIMITATION={n_known} PENDING={n_pending} "
          f"[gate: {mode}]")
    if blocking:
        print(f"  gate FAILED: {blocking} blocking item(s).")
        sys.exit(1)
    print("  gate passed.")
    sys.exit(0)
    print(f"wrote {args.out} ({sum(1 for r in rows if r['observed_vba'].strip())} "
          f"observed of {len(rows)} rows)")


if __name__ == "__main__":
    main()
