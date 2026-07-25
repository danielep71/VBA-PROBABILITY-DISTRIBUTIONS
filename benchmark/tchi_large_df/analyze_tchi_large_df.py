"""
Score the large-df t / Chi-square study once the VBA observations are exported,
and report where (if anywhere) the public kernels diverge from the true value.

Answers the P2-02 question: in the accepted-but-not-contracted regime, does the
kernel return the correct (normal-limit) value, or a silently wrong one? For each
(function, df) it prints the worst relative error of observed vs reference. Run
with no observations to see the reference-side structure (which df are genuinely
non-normal vs already at the normal limit).
"""
import csv
import os
from collections import defaultdict
from decimal import Decimal, InvalidOperation


def parse_obs(s):
    s = (s or "").strip()
    if not s or s.upper() == "ERROR":
        return None
    try:
        return sum(Decimal(p.strip()) for p in s.split(";"))
    except (InvalidOperation, ValueError):
        return "BAD"


def rel(obs, ref):
    ref = Decimal(ref)
    if ref == 0:
        return abs(obs)
    return abs((obs - ref) / ref)


def main():
    path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "tchi_large_df_grid.csv")
    rows = list(csv.DictReader(open(path)))
    by = defaultdict(list)
    for r in rows:
        by[(r["function"], r["arg2"])].append(r)

    any_obs = any((r["observed_vba"] or "").strip() for r in rows)
    print(f"{'function':<30}{'df':>10}{'n':>5}{'worst_rel':>14}  note")
    worst_overall = Decimal(0)
    for (fn, df), rs in sorted(by.items(), key=lambda k: (k[0][0], float(k[0][1]))):
        notes = {r["regime"] for r in rs}
        band = "normal-limit" if all("normal_limit" in n or "normal" in n for n in notes) else \
               ("meaningful" if any(n == "meaningful" for n in notes) else "mixed/inverse")
        if not any_obs:
            print(f"{fn:<30}{float(df):>10.0e}{len(rs):>5}{'(no obs yet)':>14}  {band}")
            continue
        worst = Decimal(-1); nbad = nmiss = 0
        for r in rs:
            o = parse_obs(r["observed_vba"])
            if o is None:
                nmiss += 1; continue
            if o == "BAD":
                nbad += 1; continue
            e = rel(o, r["reference"])
            if e > worst:
                worst = e
        if worst > worst_overall:
            worst_overall = worst
        flag = ""
        if worst >= 0:
            if worst > Decimal("1e-6"):
                flag = "  <-- DIVERGES from true value"
            elif worst > Decimal("2e-10"):
                flag = "  (outside contract-grade)"
        miss = f" [{nmiss} unobserved, {nbad} unparseable]" if (nmiss or nbad) else ""
        wtxt = f"{float(worst):.2e}" if worst >= 0 else "n/a"
        print(f"{fn:<30}{float(df):>10.0e}{len(rs):>5}{wtxt:>14}  {band}{flag}{miss}")

    if any_obs:
        print(f"\nworst relative error across the study: {float(worst_overall):.2e}")
        print("Interpretation: errors that stay near Double epsilon mean the kernel tracks")
        print("the true (normal-limit) value; a jump to O(1e-6)+ marks a silent-divergence")
        print("boundary and is the evidence needed to justify enforcing an envelope there.")
    else:
        print("\nNo observations yet. Export the VBA values into observed_vba, then re-run.")


if __name__ == "__main__":
    main()
