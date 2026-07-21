"""
Validate the older-families contracts against the fresh consolidated holdout.

Presence-based selection (function + regime present in holdout_older_grid.csv),
metric-aware error (absolute when the contract metric is 'absolute', else
relative). Reports worst holdout error vs threshold with margin, writes
holdout_older_summary.md, and exits non-zero if any contract fails - so passing
is the evidence for flipping provenance to 'validated and frozen'.
"""
import argparse, csv, os
from decimal import Decimal, getcontext
getcontext().prec = 50

def parse(s):
    s = s.strip()
    return None if (not s or s.upper() == "ERROR") else sum(Decimal(p) for p in s.split(";"))

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--grid", default="holdout_older_grid.csv")
    ap.add_argument("--contracts", default=os.path.join("..", "accuracy_contracts.csv"))
    a = ap.parse_args()
    rows = list(csv.DictReader(open(a.grid)))
    contracts = list(csv.DictReader(open(a.contracts)))

    grid_by = {}
    for r in rows:
        grid_by.setdefault((r["function"], r["regime"]), []).append(r)

    results = []
    tested = nfail = 0
    for c in sorted(contracts, key=lambda c: c["contract_id"]):
        matched = grid_by.get((c["function"], c["regime"]), [])
        if not matched or not c["threshold"].strip():
            continue
        absolute = c["metric"].strip().lower() == "absolute"
        worst = Decimal(-1); at = ""
        for r in matched:
            o = parse(r["observed_vba"])
            if o is None:
                continue
            ref = Decimal(r["reference"])
            ae = abs(o - ref)
            e = ae if absolute else (ae / abs(ref) if ref != 0 else Decimal(0))
            if e > worst:
                worst = e; at = f"a1={r['arg1']} a2={r['arg2']} a3={r['arg3']}"
        if worst < 0:
            results.append((c, None, None, "NO OBS")); continue
        thr = Decimal(c["threshold"]); ok = worst <= thr; tested += 1
        margin = float(thr / worst) if worst > 0 else float("inf")
        if not ok: nfail += 1
        results.append((c, worst, margin, "PASS" if ok else "FAIL"))

    md = ["# Older-families holdout summary", "",
          "Fresh, off-compliance-grid validation of the contracts that predated the "
          "independent-holdout discipline. Error is absolute where the contract metric "
          "is absolute, else relative. Passing here is the evidence for flipping "
          "provenance to `validated and frozen`.", "",
          "| Contract | Metric | Threshold | Holdout worst | Margin | Provenance | Verdict |",
          "|---|---|---|---:|---:|---|---|"]
    for c, worst, margin, verdict in results:
        wt = f"{float(worst):.2e}" if worst is not None else "-"
        mt = f"{margin:.1f}x" if margin is not None else "-"
        md.append(f"| {c['contract_id']} | {c['metric']} | {c['threshold']} | {wt} | {mt} | {c['provenance']} | {verdict} |")
    npass = sum(1 for _,_,_,v in results if v == "PASS")
    md += ["", f"> {npass} pass, {nfail} fail across {len(results)} contract(s) with fresh holdout points."]
    with open("holdout_older_summary.md", "w", encoding="utf-8") as f:
        f.write("\n".join(md) + "\n")

    print(f"tested {tested} older contracts on fresh holdout: {npass} pass, {nfail} fail")
    print("wrote holdout_older_summary.md")
    for c, worst, margin, verdict in results:
        if verdict != "PASS":
            print(f"  {verdict}: {c['contract_id']} (worst {float(worst):.2e} vs {c['threshold']})" if worst else f"  {verdict}: {c['contract_id']}")
    if nfail:
        raise SystemExit(1)

if __name__ == "__main__":
    main()
