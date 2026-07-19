"""
Independent holdout for the regime-specific contracts.

Every point here is FRESH: shapes, ratios and probabilities that were NOT used to
set any threshold. If the provisional thresholds hold on this unseen data, they
generalise and can be frozen. Covers:

  Beta density/CDF/survival (unbalanced)     -> output_error
  Beta inverse (unbalanced)                  -> quantile_error + tail residual
  F CDF/survival (validated range)           -> output_error
  F inverse (validated range)                -> quantile_error + tail residual
  PROB_LogBeta (all, dispatched)             -> log_absolute_error

Grid schema matches the main grid so `compute_errors.py --grid holdout_grid.csv`
verdicts it directly against the frozen contract.
References: mpmath / continued-fraction incomplete beta at 50 digits.
"""
import argparse, csv
import mpmath as mp
from _ibeta import ibeta, beta_invcdf, f_cdf, f_invcdf
mp.mp.dps = 50

# FRESH unbalanced Beta shapes (min/max < 0.1), none in the fitting set
BETA_UNBAL = [("0.55","3000"),("1.9","50000"),("3.3","40000"),("0.42","700"),
              ("7.7","200"),("250","1.15"),("0.9","250000"),("4.4","90000")]
# FRESH F df within the validated incomplete-beta range (param < ~1E7)
F_VAL = [("3","5000"),("1.5","200000"),("7","50000"),("500000","4")]
# FRESH probabilities (extra tails + central), distinct from the fitting set
PROBS = ["0.0001","0.005","0.25","0.75","0.995","0.9999"]
F_PROBS = ["0.25","0.75","0.98","0.995"]
# FRESH Small values + near-seam / between-decade ratios for PROB_LogBeta
LB_SMALL = ["0.42","1.9","3.3","7.7"]
LB_RATIO = ["0.3","0.15","0.11","0.101","0.099","0.09","0.075","0.03","0.003","3E-4","3E-6","3E-9"]

def beta_pdf(x,a,b):
    x,a,b=mp.mpf(x),mp.mpf(a),mp.mpf(b)
    return mp.e**((a-1)*mp.log(x)+(b-1)*mp.log(1-x)-mp.log(mp.beta(a,b)))

def logbeta(a,b):
    a,b=mp.mpf(a),mp.mpf(b)
    return mp.loggamma(a)+mp.loggamma(b)-mp.loggamma(a+b)

def row(fn,kernel,a1,a2,a3,ref,regime):
    return {"function":fn,"vba_kernel":kernel,"claim":"","metric":"",
            "arg1":a1,"arg2":a2,"arg3":a3,"reference":mp.nstr(ref,30),
            "observed_vba":"","regime":regime,"evidence_set":"holdout"}

def build():
    rows=[]
    # Beta forward unbalanced (density/CDF/survival), X at the mass
    for a,b in BETA_UNBAL:
        A,B=mp.mpf(a),mp.mpf(b); x=A/(A+B)
        rows.append(row("Beta_Density","K_STATS_Beta_Density",mp.nstr(x,17),a,b,beta_pdf(x,A,B),"unbalanced"))
        rows.append(row("Beta_Cumulative","K_STATS_Beta_Cumulative",mp.nstr(x,17),a,b,ibeta(x,A,B),"unbalanced"))
        rows.append(row("Beta_Survival","K_STATS_Beta_Survival",mp.nstr(x,17),a,b,1-ibeta(x,A,B),"unbalanced"))
    # Beta inverse unbalanced
    for a,b in BETA_UNBAL[:5]:
        for p in PROBS:
            rows.append(row("Beta_InverseCumulative","K_STATS_Beta_InverseCumulative",p,a,b,
                            beta_invcdf(p,mp.mpf(a),mp.mpf(b)),"unbalanced"))
    # F forward validated
    for d1,d2 in F_VAL:
        for p in ("0.5",):  # x at median-ish via invcdf
            x=f_invcdf("0.5",mp.mpf(d1),mp.mpf(d2))
        for xval in (mp.mpf(1), x):
            rows.append(row("F_Cumulative","K_STATS_F_Cumulative",mp.nstr(xval,17),d1,d2,f_cdf(xval,mp.mpf(d1),mp.mpf(d2)),"validated"))
            rows.append(row("F_Survival","K_STATS_F_Survival",mp.nstr(xval,17),d1,d2,1-f_cdf(xval,mp.mpf(d1),mp.mpf(d2)),"validated"))
    # F inverse validated
    for d1,d2 in F_VAL:
        for p in F_PROBS:
            rows.append(row("F_InverseCumulative","K_STATS_F_InverseCumulative",p,d1,d2,
                            f_invcdf(p,mp.mpf(d1),mp.mpf(d2)),"validated"))
    # PROB_LogBeta dispatched (fresh Small x near-seam/between-decade ratios)
    for sm in LB_SMALL:
        for r in LB_RATIO:
            small=mp.mpf(sm); large=small/mp.mpf(r)
            rows.append(row("PROB_LogBeta","PROB_LogBeta",mp.nstr(large,17),mp.nstr(small,17),"",
                            logbeta(large,small),"all"))
    return rows

def main():
    ap=argparse.ArgumentParser(); ap.add_argument("--out",default="holdout_grid.csv"); a=ap.parse_args()
    rows=build()
    fields=["function","vba_kernel","claim","metric","arg1","arg2","arg3","reference","observed_vba","regime","evidence_set"]
    with open(a.out,"w",newline="") as f:
        w=csv.DictWriter(f,fieldnames=fields); w.writeheader(); w.writerows(rows)
    print(f"wrote {a.out}: {len(rows)} rows")

if __name__=="__main__":
    main()
