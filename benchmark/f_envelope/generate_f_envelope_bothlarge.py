"""
F envelope study - BOTH-LARGE orientation.

The one-large sweep (one df large, the other ~1) is not the worst case. When BOTH
incomplete-beta parameters are large the continued fraction degrades earlier (an
existing test already used a loose 1E-6 tolerance for F_cdf(1, 1E5, 1E5)). This
sweep measures d1 = d2 = df and d1 = df, d2 = 3*df across a df range, placing x at
fixed CDF levels via the (fast) beta inverse so no reference is degenerate.
The strict F envelope must be the MINIMUM boundary over all orientations.
"""
import argparse, csv
import mpmath as mp
from scipy.special import betaincinv
from _ibeta import ibeta
mp.mp.dps = 40

DF_VALUES = [100, 300, 1e3, 3e3, 1e4, 3e4, 1e5, 3e5, 1e6]
PROBS = [0.1, 0.25, 0.5, 0.75, 0.9]

def ref_cdf(x, d1, d2):
    x, d1, d2 = mp.mpf(x), mp.mpf(d1), mp.mpf(d2)
    return ibeta(d1 * x / (d1 * x + d2), d1 / 2, d2 / 2)

def rowdict(fn, xs, d1, d2, ref, tag):
    return {"function": fn, "vba_kernel": f"K_STATS_{fn}", "claim": "rel<=1.1E-10",
            "metric": "rel", "arg1": xs, "arg2": mp.nstr(mp.mpf(d1), 17),
            "arg3": mp.nstr(mp.mpf(d2), 17), "reference": mp.nstr(ref, 30),
            "observed_vba": "", "regime": "extreme_df", "evidence_set": f"f_envelope:{tag}"}

def build():
    rows = []
    for df in DF_VALUES:
        for d1, d2, tag in ((df, df, "botheq"), (df, 3 * df, "bothbig")):
            a, b = d1 / 2, d2 / 2
            for p in PROBS:
                y = betaincinv(a, b, p)
                x = d2 * y / (d1 * (1 - y))
                xs = mp.nstr(mp.mpf(x), 17)
                c = ref_cdf(xs, d1, d2)
                rows.append(rowdict("F_Cumulative", xs, d1, d2, c, tag))
                rows.append(rowdict("F_Survival", xs, d1, d2, 1 - c, tag))
    return rows

def main():
    ap = argparse.ArgumentParser(); ap.add_argument("--out", default="f_envelope_bothlarge_grid.csv")
    a = ap.parse_args()
    rows = build()
    fields = ["function","vba_kernel","claim","metric","arg1","arg2","arg3","reference","observed_vba","regime","evidence_set"]
    with open(a.out, "w", newline="") as f:
        w = csv.DictWriter(f, fieldnames=fields); w.writeheader(); w.writerows(rows)
    print(f"wrote {a.out}: {len(rows)} rows")

if __name__ == "__main__":
    main()
