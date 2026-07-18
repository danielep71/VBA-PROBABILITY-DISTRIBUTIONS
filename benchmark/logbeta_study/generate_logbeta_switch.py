"""
Unbalanced-Beta switch study for PROB_LogBeta.

PROB_LogBeta chooses between:
  - the asymptotic branch  LogGamma(Small) - Small*Log(Large), used when
    Small/Large <= PROB_EPS (1E-15);
  - the general identity    LogGamma(A) + LogGamma(B) - LogGamma(A+B).

The general identity loses precision by cancellation as A and B become
unbalanced (error ~ macheps * Large/Small). This grid measures PROB_LogBeta
across the ratio range so we can see empirically where the general case
degrades and whether the 1E-15 switch is well placed.

Reference values are mpmath at 50+ digits. Output schema matches the main
accuracy grid so compute_errors.py can parse the hi;lo observations.
"""
import argparse, csv
import mpmath as mp
mp.mp.dps = 60

def log_beta(a, b):
    return mp.log(mp.beta(mp.mpf(a), mp.mpf(b)))

def build():
    rows = []
    # Sweep ratio Small/Large from balanced (1E-1) to deep-unbalanced (1E-18),
    # at several Small values (avoiding 0.5, which hits the half-integer shortcut).
    smalls = [mp.mpf("0.8"), mp.mpf(1), mp.mpf("1.5"), mp.mpf(3), mp.mpf(10)]
    exps = list(range(1, 19))  # ratio = 1E-1 .. 1E-18
    for small in smalls:
        for e in exps:
            ratio = mp.mpf(10) ** (-e)
            large = small / ratio
            # arg1 = A (large), arg2 = B (small); PROB_LogBeta is symmetric
            ref = log_beta(large, small)
            rows.append({
                "function": "LogBeta",
                "vba_kernel": "PROB_LogBeta",
                "claim": "rel<=5E-15",          # the inherited Beta accuracy claim
                "metric": "rel",
                "arg1": mp.nstr(large, 17),
                "arg2": mp.nstr(small, 17),
                "arg3": "",
                "reference": mp.nstr(ref, 30),
                "observed_vba": "",
            })
    return rows

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default="logbeta_switch_grid.csv")
    args = ap.parse_args()
    rows = build()
    fields = ["function","vba_kernel","claim","metric","arg1","arg2","arg3","reference","observed_vba"]
    with open(args.out, "w", newline="") as f:
        w = csv.DictWriter(f, fieldnames=fields)
        w.writeheader()
        w.writerows(rows)
    print(f"wrote {args.out}: {len(rows)} reference rows at 50+ digits")

if __name__ == "__main__":
    main()
