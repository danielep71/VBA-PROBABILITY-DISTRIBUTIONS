"""
F accuracy-envelope study - GAP sweep.

Finds the df boundary where the public F CDF/survival cross the 1.1E-10 accuracy
contract. F_CDF(x; d1, d2) = I_y(d1/2, d2/2), y = d1*x / (d1*x + d2), so the
degrading quantity is the incomplete-beta shape parameter max(d1/2, d2/2). This
sweep fills the region below the base sweep (generate_f_envelope.py): that
parameter from ~50 to 1E6, in BOTH orientations (large first vs large second
beta parameter), because the continued fraction can behave differently on each
side. References are the 50-digit continued-fraction incomplete beta, as in the
base sweep, evaluated at the exact binary64 argument each row stores, with the
survival computed directly (#159; see _fenv_common.py).

The committed grid holds Excel observations. Re-reference it, changing
reference cells only, with

    python generate_f_envelope_gap.py --merge-references          # report
    python generate_f_envelope_gap.py --merge-references --write  # patch
"""
import argparse, csv
import mpmath as mp
from _fenv_common import FIELDS, merge, rows_for
mp.mp.dps = 50

# incomplete-beta shape parameter targets (= df/2), ~50 to 1E6, below the base sweep
BETA_PARAMS = [50, 100, 200, 500, 1e3, 2e3, 5e3, 1e4, 2e4, 5e4, 1e5, 2e5, 5e5, 1e6]
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
    ap = argparse.ArgumentParser(); ap.add_argument("--out", default="f_envelope_gap_grid.csv")
    ap.add_argument("--merge-references", action="store_true",
                    help="re-reference the committed grid, keeping every observation")
    ap.add_argument("--write", action="store_true",
                    help="with --merge-references: patch the committed grid")
    a = ap.parse_args()
    rows = build()
    if a.merge_references:
        merge("f_envelope_gap_grid.csv", rows, a.write)
        return
    with open(a.out, "w", newline="") as f:
        w = csv.DictWriter(f, fieldnames=FIELDS, lineterminator="\n"); w.writeheader(); w.writerows(rows)
    print(f"wrote {a.out}: {len(rows)} rows ({len(BETA_PARAMS)} beta params x 2 orientations x {len(XS)} x x 2 functions)")

if __name__ == "__main__":
    main()
