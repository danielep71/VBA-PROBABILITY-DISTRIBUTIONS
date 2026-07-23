"""
Inverse-function study for the unbalanced Beta/F regime.

Inverse solvers amplify the LogBeta normalization error differently from the
forward functions, so they are measured separately. Two metrics are assessed by
the analyzer (this generator supplies the reference quantile x_ref):

  quantile error       abs/rel error in x;
  forward residual     push x_VBA back through the TRUE (mpmath) CDF and compare
                       the recovered probability to the target p, plus a
                       tail-relative residual  |I_xVBA - p| / min(p, 1-p).

Only the validated regime (incomplete-beta shape params < ~1E7) is referenced;
extreme-df F inverse falls under the separate incomplete-beta convergence limit.
"""
import argparse, csv
import mpmath as mp
import os as _os, sys as _sys
# Single-sourced reference helper: benchmark/_ibeta.py is the only copy.
_sys.path.insert(0, _os.path.dirname(_os.path.dirname(_os.path.abspath(__file__))))
from _ibeta import beta_invcdf, f_invcdf
mp.mp.dps = 50

BETA_CASES = [("0.7","1000"),("2.5","1000000"),("100000","2.5"),
              ("10.25","68"),("0.8","10000"),("1000","0.8")]
F_CASES = [("1","10000"),("1000000","3"),("2.5","1000000"),("5","100000")]
PROBS = ["0.5","0.9","0.99","0.01","0.001","0.999"]
F_PROBS = ["0.5","0.9","0.95","0.99"]

def add(rows, fn, kernel, p, a2, a3, ref):
    rows.append({"function": fn, "vba_kernel": kernel, "arg1": p,
                 "arg2": mp.nstr(mp.mpf(a2),17), "arg3": mp.nstr(mp.mpf(a3),17),
                 "reference": mp.nstr(ref,30), "observed_vba": ""})

def build():
    rows = []
    for a, b in BETA_CASES:
        for p in PROBS:
            add(rows, "Beta_InverseCumulative", "K_STATS_Beta_InverseCumulative",
                p, a, b, beta_invcdf(p, mp.mpf(a), mp.mpf(b)))
    for d1, d2 in F_CASES:
        for p in F_PROBS:
            add(rows, "F_InverseCumulative", "K_STATS_F_InverseCumulative",
                p, d1, d2, f_invcdf(p, mp.mpf(d1), mp.mpf(d2)))
    return rows

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default="beta_f_inverse_grid.csv")
    a = ap.parse_args()
    rows = build()
    with open(a.out, "w", newline="") as f:
        w = csv.DictWriter(f, fieldnames=["function","vba_kernel","arg1","arg2","arg3","reference","observed_vba"])
        w.writeheader(); w.writerows(rows)
    print(f"wrote {a.out}: {len(rows)} rows")

if __name__ == "__main__":
    main()
