"""
F accuracy-envelope study.

Finds the df boundary where the public F CDF/survival cross the 1.1E-10 accuracy
contract. F_CDF(x; d1, d2) = I_y(d1/2, d2/2), y = d1*x / (d1*x + d2), so the
degrading quantity is the incomplete-beta shape parameter max(d1/2, d2/2). This
sweeps that parameter finely around 1E7 in BOTH orientations (large first vs
large second beta parameter), because the continued fraction can behave
differently on each side. References are the 50-digit continued-fraction
incomplete beta, validated self-consistent to ~1E-41 in this region.
"""
import argparse, csv
import mpmath as mp
import os as _os, sys as _sys
# Single-sourced reference helper: benchmark/_ibeta.py is the only copy.
_sys.path.insert(0, _os.path.dirname(_os.path.dirname(_os.path.abspath(__file__))))
from _ibeta import ibeta
mp.mp.dps = 50

# incomplete-beta shape parameter targets (= df/2), fine around 1E7
BETA_PARAMS = [1e6, 3e6, 6e6, 1e7, 1.5e7, 2e7, 3e7, 5e7, 8e7,
               1.5e8, 3e8, 6e8, 1e9, 3e9, 5e9]
XS = ["0.3", "0.6", "1.0", "1.7", "3.0"]   # span the mass for large-df F

def f_cdf(x, d1, d2):
    x, d1, d2 = mp.mpf(x), mp.mpf(d1), mp.mpf(d2)
    return ibeta(d1 * x / (d1 * x + d2), d1 / 2, d2 / 2)

def row(fn, x, d1, d2, ref, orient):
    return {"function": fn, "vba_kernel": f"K_STATS_{fn}", "claim": "rel<=1.1E-10",
            "metric": "rel", "arg1": x, "arg2": mp.nstr(mp.mpf(d1), 17),
            "arg3": mp.nstr(mp.mpf(d2), 17), "reference": mp.nstr(ref, 30),
            "observed_vba": "", "regime": "extreme_df", "evidence_set": f"f_envelope:{orient}"}

def build():
    rows = []
    for B in BETA_PARAMS:
        df = 2 * B
        # orientation 1: large SECOND beta param (d1 small, d2 = 2B)
        for x in XS:
            c = f_cdf(x, 1.0, df)
            rows.append(row("F_Cumulative", x, 1.0, df, c, "d2big"))
            rows.append(row("F_Survival",   x, 1.0, df, 1 - c, "d2big"))
        # orientation 2: large FIRST beta param (d1 = 2B, d2 small)
        for x in XS:
            c = f_cdf(x, df, 1.0)
            rows.append(row("F_Cumulative", x, df, 1.0, c, "d1big"))
            rows.append(row("F_Survival",   x, df, 1.0, 1 - c, "d1big"))
    return rows

def main():
    ap = argparse.ArgumentParser(); ap.add_argument("--out", default="f_envelope_grid.csv")
    a = ap.parse_args()
    rows = build()
    fields = ["function","vba_kernel","claim","metric","arg1","arg2","arg3","reference","observed_vba","regime","evidence_set"]
    with open(a.out, "w", newline="") as f:
        w = csv.DictWriter(f, fieldnames=fields); w.writeheader(); w.writerows(rows)
    print(f"wrote {a.out}: {len(rows)} rows ({len(BETA_PARAMS)} beta params x 2 orientations x {len(XS)} x x 2 functions)")

if __name__ == "__main__":
    main()
