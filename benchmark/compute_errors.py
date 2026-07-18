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
from decimal import Decimal, getcontext, InvalidOperation
import csv
import datetime as _dt
import math
import re


getcontext().prec = 50  # exact-enough for Double reconstruction and error ratios


def parse_claim(claim):
    # e.g. "rel<6.1E-14" or "abs<=3E-17" -> (metric, threshold)
    m = re.match(r"(rel|abs)\s*<?=?\s*([0-9.eE+-]+)", claim)
    if not m:
        return None, None
    try:
        return m.group(1), Decimal(m.group(2))
    except InvalidOperation:
        return m.group(1), None
    # (unreachable legacy line below retained for clarity)


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
    # Sum the two-part 'hi;lo' export in Decimal so the full Double is preserved
    # and the later obs-ref subtraction does not cancel in binary float.
    total = Decimal(0)
    for part in s.split(";"):
        total += Decimal(part.strip())
    return total


def errors(observed, reference, metric):
    o = parse_observed(observed)
    if o is None:
        raise ValueError("no observed value")
    r = Decimal(str(reference).strip())
    abs_e = abs(o - r)
    if r != 0:
        rel_e = abs_e / abs(r)
    else:
        rel_e = Decimal(0) if o == 0 else Decimal("Infinity")
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
    n_char = 0
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
        worst_e, worst_at = Decimal(-1), ""
        for r in measured:
            try:
                _, e = errors(r["observed_vba"], r["reference"], cmetric)
            except (ValueError, ZeroDivisionError, InvalidOperation):
                e = Decimal("Infinity")
            if e > worst_e:
                worst_e = e
                args_str = ", ".join(a for a in (r["arg1"], r["arg2"], r["arg3"]) if a)
                worst_at = args_str
        # A measured error is judged directly against the contract threshold in
        # Decimal. There is no blanket precision-floor exemption: the two-part
        # hi;lo export preserves the full Double, so a miss is a miss unless the
        # contract explicitly classifies the function otherwise.
        ok = threshold is not None and worst_e <= threshold
        if fn_status == "known_limitation":
            # Documented defect: not a silent pass, not a hard fail.
            n_known += 1
            verdict = "🔷 KNOWN LIMITATION"
        elif fn_status == "characterization_only":
            # Measured for the record, not held to a pass/fail claim.
            n_char += 1
            verdict = "🧪 CHARACTERIZATION ONLY"
        elif ok:
            verdict = "✅ PASS"
        else:
            n_fail += 1
            verdict = "❌ FAIL"
        lines.append(f"| {fn} | {claim} | {cmetric} | {float(worst_e):.2e} | "
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
                     f"CHARACTERIZATION ONLY: {n_char}, PENDING: {n_pending}.")
        lines.append("")
        lines.append("> States: **PASS** meets the contract; **FAIL** exceeds it and must be "
                     "investigated; **KNOWN LIMITATION** is a documented defect tracked in "
                     "`accuracy_contracts.csv` (does not read as green); **CHARACTERIZATION "
                     "ONLY** is measured for the record but not held to a pass/fail claim; "
                     "**PENDING** is not yet measured. Errors are computed in Decimal from the "
                     "two-part hi;lo export, so a miss is a real miss (no precision-floor "
                     "exemption).")

    with open(args.out, "w") as f:
        f.write("\n".join(lines) + "\n")

    # ---- release gate (granular exit codes) ----
    #   exit 1 : a hard FAIL, or (strict mode) an unresolved KNOWN LIMITATION.
    #   exit 2 : required observations are missing (PENDING) with nothing worse.
    #   exit 0 : all active contracts pass.
    # CHARACTERIZATION ONLY never blocks. Development mode
    # (--allow-known-limitations) does not block on KNOWN LIMITATION, but FAIL
    # and PENDING still apply.
    mode = "development (known limitations allowed)" if args.allow_known_limitations else "strict"
    print(f"{args.out}: FAIL={n_fail} KNOWN_LIMITATION={n_known} "
          f"CHARACTERIZATION_ONLY={n_char} PENDING={n_pending} [gate: {mode}]")

    fail_block = n_fail + (0 if args.allow_known_limitations else n_known)
    if fail_block:
        print(f"  gate FAILED (exit 1): {fail_block} blocking item(s) "
              f"[FAIL={n_fail}"
              f"{'' if args.allow_known_limitations else f', KNOWN LIMITATION={n_known}'}].")
        sys.exit(1)
    if n_pending:
        print(f"  gate INCOMPLETE (exit 2): {n_pending} function(s) not yet measured.")
        sys.exit(2)
    print("  gate passed (exit 0).")
    sys.exit(0)
    print(f"wrote {args.out} ({sum(1 for r in rows if r['observed_vba'].strip())} "
          f"observed of {len(rows)} rows)")


if __name__ == "__main__":
    main()
