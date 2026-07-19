"""
Step-6 ground-truth study: relative error of the PUBLIC Beta/F functions at
strongly disparate shapes / degrees of freedom, now that PROB_LogBeta uses the
stable log-gamma difference. This measures what actually reaches users, rather
than the LogBeta proxy.

References: mpmath at 60 digits.
  Beta_Density(x,a,b)    = exp((a-1)ln x + (b-1)ln(1-x) - ln Beta(a,b))
  Beta_Cumulative(x,a,b) = regularized incomplete beta I_x(a,b)
  Beta_Survival          = 1 - CDF
  F_Cumulative(x,d1,d2)  = I_{d1 x/(d1 x+d2)}(d1/2, d2/2)
  F_Survival             = 1 - CDF

X is chosen near each distribution's mass (mean / median) so the value is
representable and the relative error is meaningful.
"""
import argparse, csv
import mpmath as mp
mp.mp.dps = 60

def beta_pdf(x, a, b):
    x, a, b = mp.mpf(x), mp.mpf(a), mp.mpf(b)
    return mp.e ** ((a - 1) * mp.log(x) + (b - 1) * mp.log(1 - x) - mp.log(mp.beta(a, b)))

def beta_cdf(x, a, b):
    return mp.betainc(mp.mpf(a), mp.mpf(b), 0, mp.mpf(x), regularized=True)

def f_cdf(x, d1, d2):
    x, d1, d2 = mp.mpf(x), mp.mpf(d1), mp.mpf(d2)
    y = d1 * x / (d1 * x + d2)
    return mp.betainc(d1 / 2, d2 / 2, 0, y, regularized=True)

# strongly disparate Beta shapes (small, large) and a representative X near the mean
BETA_CASES = [
    ("0.7", "1000"), ("2.5", "1000000"), ("10.25", "68"),
    ("0.8", "10000"), ("1.3", "100000"), ("5.75", "1000000"),
    ("1000", "0.8"), ("100000", "2.5"),
]
# strongly asymmetric F degrees of freedom
F_CASES = [("1", "10000"), ("2.5", "100000000"), ("10", "10000000000"), ("1000000", "3")]

def add(rows, fn, kernel, a1, a2, a3, ref):
    rows.append({"function": fn, "vba_kernel": kernel, "arg1": mp.nstr(a1, 17),
                 "arg2": mp.nstr(a2, 17), "arg3": mp.nstr(a3, 17),
                 "reference": mp.nstr(ref, 30), "observed_vba": ""})

def build():
    rows = []
    for a, b in BETA_CASES:
        A, B = mp.mpf(a), mp.mpf(b)
        mean = A / (A + B)
        for x in (mean, (A + mp.mpf("0.5")) / (A + B + 1)):  # mean and a mode-ish point
            if not (0 < x < 1):
                continue
            add(rows, "Beta_Density",    "K_STATS_Beta_Density",    x, A, B, beta_pdf(x, A, B))
            add(rows, "Beta_Cumulative", "K_STATS_Beta_Cumulative", x, A, B, beta_cdf(x, A, B))
            add(rows, "Beta_Survival",   "K_STATS_Beta_Survival",   x, A, B, 1 - beta_cdf(x, A, B))
    for d1, d2 in F_CASES:
        D1, D2 = mp.mpf(d1), mp.mpf(d2)
        for x in (mp.mpf(1), D2 / (D2 - 2) if D2 > 2 else mp.mpf(1)):  # x=1 and near the mean
            add(rows, "F_Cumulative", "K_STATS_F_Cumulative", x, D1, D2, f_cdf(x, D1, D2))
            add(rows, "F_Survival",   "K_STATS_F_Survival",   x, D1, D2, 1 - f_cdf(x, D1, D2))
    return rows

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default="beta_f_unbalanced_grid.csv")
    a = ap.parse_args()
    rows = build()
    with open(a.out, "w", newline="") as f:
        w = csv.DictWriter(f, fieldnames=["function","vba_kernel","arg1","arg2","arg3","reference","observed_vba"])
        w.writeheader(); w.writerows(rows)
    print(f"wrote {a.out}: {len(rows)} rows")

if __name__ == "__main__":
    main()
