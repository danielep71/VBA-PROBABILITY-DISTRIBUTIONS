"""
Validate the older-families contracts against the fresh consolidated holdout.

Presence-based selection (function + regime present in holdout_older_grid.csv),
metric-aware error (absolute when the contract metric is 'absolute', else
relative). Reports worst holdout error vs threshold with margin, writes
holdout_older_summary.md, and exits non-zero if any contract fails - so passing
is the evidence for flipping provenance to 'validated and frozen'.

Since #327 each contract is scored against its guarantee (accuracy_tiers.py);
the registry threshold is the measured target. A contract inside its guarantee
but above its target passes as "PASS · target missed", and its rows are listed
under "Rows above the measured target", which never fails the run.

The summary also lists every failing row with its arguments, reference,
observation and contract, and every grid row that no thresholded contract
claims, so a failure or a silently unscored row is visible in the committed
record rather than only in this script's console output (#164).

An unscored row has one accepted disposition: a contract for its function and
regime whose status is `characterization_only`, the registry's existing marker
for evidence that is measured but not claimed. Such a row is listed with that
contract. Any other unscored row is listed with no disposition, and
check_holdout_older_summary.py fails on it (#171).

Usage, from this directory:

    python analyze_holdout_older.py          rewrite holdout_older_summary.md

check_holdout_older_summary.py renders the same text in memory and fails when
the committed summary differs, so a registry or grid change that is not
followed by a re-run fails the evidence tooling instead of going unnoticed.
"""
import argparse
import csv
import os
import sys
from decimal import Decimal, getcontext

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from accuracy_tiers import derive as derive_guarantee  # noqa: E402

getcontext().prec = 50

HERE = os.path.dirname(os.path.abspath(__file__))
GRID = os.path.join(HERE, "holdout_older_grid.csv")
CONTRACTS = os.path.join(HERE, "..", "accuracy_contracts.csv")
SUMMARY = os.path.join(HERE, "holdout_older_summary.md")

CHARACTERIZATION = "characterization_only"
TARGET_MISSED = "PASS · target missed"

# Failing contracts whose cause is tracked elsewhere. The note is printed next
# to each failing row of that contract. A contract that fails and is not listed
# here still fails, and its rows are still listed, with no note.
TRACKED = {
    "Beta_Survival.balanced.output_rel": "tracked in #163",
    "Gamma_Cumulative.all.output":
        "the observation predates the 2026-07-29 and 2026-08-27 prefactor changes; see README (#164)",
    "Gamma_Survival.all.output":
        "the observation predates the 2026-07-29 and 2026-08-27 prefactor changes; see README (#164)",
}


def parse(s):
    s = s.strip()
    return None if (not s or s.upper() == "ERROR") else sum(Decimal(p) for p in s.split(";"))


def row_error(r, absolute):
    """Error of one grid row under the contract metric, or None without an observation."""
    o = parse(r["observed_vba"])
    if o is None:
        return None
    ref = Decimal(r["reference"])
    ae = abs(o - ref)
    return ae if absolute else (ae / abs(ref) if ref != 0 else Decimal(0))


def evaluate(rows, contracts):
    """Score every thresholded contract present in the grid.

    Returns (results, failing, unclaimed, above_target): one (contract, worst,
    margin, verdict) per contract, one (contract, row, error) per row above its
    guarantee, one (contract, row, error) per row above its measured target
    but within its guarantee, and
    one (row, contract) per grid row no thresholded contract claims, where the
    contract is the `characterization_only` contract for that function and
    regime, or None when the row has no disposition.
    """
    grid_by = {}
    for r in rows:
        grid_by.setdefault((r["function"], r["regime"]), []).append(r)

    results, failing, above_target, claimed = [], [], [], set()
    for c in sorted(contracts, key=lambda c: c["contract_id"]):
        matched = grid_by.get((c["function"], c["regime"]), [])
        if not matched or not c["threshold"].strip():
            continue
        claimed.add((c["function"], c["regime"]))
        absolute = c["metric"].strip().lower() == "absolute"
        thr = Decimal(derive_guarantee(c)[0])
        target = Decimal(c["threshold"])
        worst = Decimal(-1)
        for r in matched:
            e = row_error(r, absolute)
            if e is None:
                continue
            if e > thr:
                failing.append((c, r, e))
            elif e > target:
                above_target.append((c, r, e))
            if e > worst:
                worst = e
        if worst < 0:
            results.append((c, None, None, "NO OBS"))
            continue
        margin = float(thr / worst) if worst > 0 else float("inf")
        verdict = ("FAIL" if worst > thr else
                   TARGET_MISSED if worst > target else "PASS")
        results.append((c, worst, margin, verdict))
    characterized = {}
    for c in sorted(contracts, key=lambda c: c["contract_id"]):
        if c.get("status", "").strip() == CHARACTERIZATION:
            characterized.setdefault((c["function"], c["regime"]), c)
    unclaimed = [(r, characterized.get((r["function"], r["regime"])))
                 for r in rows if (r["function"], r["regime"]) not in claimed]
    return results, failing, unclaimed, above_target


def undisposed(unclaimed):
    """The unscored rows that are not marked characterization only."""
    return [r for r, c in unclaimed if c is None]


def args_text(r):
    return ", ".join(r[k] for k in ("arg1", "arg2", "arg3") if r.get(k, "").strip())


def _row_table(entries, label, limit_of):
    out = [f"| Contract | Function (arguments) | Regime | Reference | Observed | Error | "
           f"{label} | Error / {label.lower()} | Note |",
           "|---|---|---|---:|---:|---:|---:|---:|---|"]
    for c, r, e in entries:
        limit = limit_of(c)
        ratio = float(e / Decimal(limit))
        out.append(f"| {c['contract_id']} | {r['function']}({args_text(r)}) | {r['regime']} | "
                   f"{r['reference']} | {r['observed_vba']} | {float(e):.3e} | {limit} | "
                   f"{ratio:.2f}x | {TRACKED.get(c['contract_id'], '')} |")
    return out


def render(rows, contracts):
    """The summary text, the per-contract results, the failing rows and the
    unscored rows, from the inputs alone."""
    results, failing, unclaimed, above_target = evaluate(rows, contracts)
    md = ["# Older-families holdout summary", "",
          "Fresh, off-compliance-grid validation of the contracts that predated the "
          "independent-holdout discipline. Error is absolute where the contract metric "
          "is absolute, else relative. Passing here is the evidence for flipping "
          "provenance to `validated and frozen`.", "",
          "Generated by `analyze_holdout_older.py` from `holdout_older_grid.csv` and "
          "`../accuracy_contracts.csv`; `check_holdout_older_summary.py` fails when "
          "this file differs from a re-run.", "",
          "| Contract | Metric | Guarantee | Measured target | Holdout worst | Margin | "
          "Provenance | Verdict |",
          "|---|---|---|---|---:|---:|---|---|"]
    for c, worst, margin, verdict in results:
        wt = f"{float(worst):.2e}" if worst is not None else "-"
        mt = f"{margin:.1f}x" if margin is not None else "-"
        md.append(f"| {c['contract_id']} | {c['metric']} | {derive_guarantee(c)[0]} | "
                  f"{c['threshold']} | {wt} | {mt} | {c['provenance']} | {verdict} |")
    npass = sum(1 for _, _, _, v in results if v in ("PASS", TARGET_MISSED))
    nfail = sum(1 for _, _, _, v in results if v == "FAIL")
    nmiss = sum(1 for _, _, _, v in results if v == TARGET_MISSED)
    md += ["", f"> {npass} pass, {nfail} fail across {len(results)} contract(s) with fresh "
           f"holdout points; {nmiss} passing contract(s) exceed their measured target "
           "(reported, not gating). Margin = guarantee / holdout worst."]

    md += ["", "## Failing rows", ""]
    if failing:
        md += ["Every row whose error exceeds its contract's guarantee.", ""]
        md += _row_table(failing, "Guarantee", lambda c: derive_guarantee(c)[0])
    else:
        md.append("None.")

    md += ["", "## Rows above the measured target", ""]
    if above_target:
        md += ["Rows within their guarantee but above the registry threshold, the "
               "measured target. Reported, not gating (#327).", ""]
        md += _row_table(above_target, "Target", lambda c: c["threshold"])
    else:
        md.append("None.")

    md += ["", "## Rows no contract scores", ""]
    if unclaimed:
        md += ["No contract with a threshold has this function and regime, so these rows "
               "are not scored above. A row is accepted only when a `characterization_only` "
               "contract covers it; `check_holdout_older_summary.py` fails on a row with no "
               "disposition.", "",
               "| Function (arguments) | Regime | Disposition |", "|---|---|---|"]
        for r, c in unclaimed:
            disposition = (f"characterization only ({c['contract_id']})" if c is not None
                           else "none")
            md.append(f"| {r['function']}({args_text(r)}) | {r['regime']} | {disposition} |")
    else:
        md.append("None.")
    return "\n".join(md) + "\n", results, failing, unclaimed, above_target


def load(grid, contracts):
    with open(grid, encoding="utf-8", newline="") as f:
        rows = list(csv.DictReader(f))
    with open(contracts, encoding="utf-8", newline="") as f:
        regs = list(csv.DictReader(f))
    return rows, regs


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--grid", default=GRID)
    ap.add_argument("--contracts", default=CONTRACTS)
    ap.add_argument("--out", default=SUMMARY)
    a = ap.parse_args()
    text, results, failing, unclaimed, _ = render(*load(a.grid, a.contracts))
    with open(a.out, "w", encoding="utf-8", newline="\n") as f:
        f.write(text)

    tested = sum(1 for _, w, _, _ in results if w is not None)
    npass = sum(1 for _, _, _, v in results if v in ("PASS", TARGET_MISSED))
    nfail = sum(1 for _, _, _, v in results if v == "FAIL")
    print(f"tested {tested} older contracts on fresh holdout: {npass} pass, {nfail} fail")
    print(f"wrote {a.out}")
    for c, worst, margin, verdict in results:
        if verdict not in ("PASS", TARGET_MISSED):
            print(f"  {verdict}: {c['contract_id']} (worst {float(worst):.2e} vs {c['threshold']})" if worst else f"  {verdict}: {c['contract_id']}")
    for r in undisposed(unclaimed):
        print(f"  UNSCORED: {r['function']}({args_text(r)}) in regime {r['regime']} "
              "has no contract and is not marked characterization only")
    if nfail:
        raise SystemExit(1)


if __name__ == "__main__":
    main()
