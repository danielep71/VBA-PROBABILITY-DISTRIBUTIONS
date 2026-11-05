"""
F accuracy-envelope study.

Finds the df boundary where the public F CDF/survival cross the 1.1E-10 accuracy
contract. F_CDF(x; d1, d2) = I_y(d1/2, d2/2), y = d1*x / (d1*x + d2), so the
degrading quantity is the incomplete-beta shape parameter max(d1/2, d2/2). This
sweeps that parameter finely around 1E7 in BOTH orientations (large first vs
large second beta parameter), because the continued fraction can behave
differently on each side. References are the 50-digit continued-fraction
incomplete beta, validated self-consistent to ~1E-41 in this region, evaluated
at the exact binary64 argument each row stores, with the survival computed
directly (#159; see _fenv_common.py).

The committed grid holds Excel observations. Re-reference it, changing
reference cells only, with

    python generate_f_envelope.py --merge-references          # report
    python generate_f_envelope.py --merge-references --write  # patch
"""
import argparse, csv
import mpmath as mp
from _fenv_common import FIELDS, merge, rows_for
mp.mp.dps = 50

# incomplete-beta shape parameter targets (= df/2), fine around 1E7
BETA_PARAMS = [1e6, 3e6, 6e6, 1e7, 1.5e7, 2e7, 3e7, 5e7, 8e7,
               1.5e8, 3e8, 6e8, 1e9, 3e9, 5e9]
XS = ["0.3", "0.6", "1.0", "1.7", "3.0"]   # span the mass for large-df F

def build():
    rows = []
    for B in BETA_PARAMS:
        df = 2 * B
        # orientation 1: large SECOND beta param (d1 small, d2 = 2B)
        for x in XS:
            rows += rows_for(x, 1.0, df, "d2big")
        # orientation 2: large FIRST beta param (d1 = 2B, d2 small)
        for x in XS:
            rows += rows_for(x, df, 1.0, "d1big")
    return rows

def main():
    ap = argparse.ArgumentParser(); ap.add_argument("--out", default="f_envelope_grid.csv")
    ap.add_argument("--merge-references", action="store_true",
                    help="re-reference the committed grid, keeping every observation")
    ap.add_argument("--write", action="store_true",
                    help="with --merge-references: patch the committed grid")
    a = ap.parse_args()
    rows = build()
    if a.merge_references:
        merge("f_envelope_grid.csv", rows, a.write)
        return
    with open(a.out, "w", newline="") as f:
        w = csv.DictWriter(f, fieldnames=FIELDS, lineterminator="\n"); w.writeheader(); w.writerows(rows)
    print(f"wrote {a.out}: {len(rows)} rows ({len(BETA_PARAMS)} beta params x 2 orientations x {len(XS)} x x 2 functions)")

if __name__ == "__main__":
    main()
