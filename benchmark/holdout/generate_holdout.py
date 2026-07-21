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


# --- Discrete references (fresh holdout; none of these points set any threshold) ---
def _b_pmf(k,n,pr): k,n,pr=mp.mpf(k),mp.mpf(n),mp.mpf(pr); return mp.binomial(n,k)*pr**k*(1-pr)**(n-k)
def _b_log(k,n,pr): k,n,pr=mp.mpf(k),mp.mpf(n),mp.mpf(pr); return mp.log(mp.binomial(n,k))+k*mp.log(pr)+(n-k)*mp.log(1-pr)
def _b_cdf(k,n,pr):
    k,n,pr=mp.mpf(k),mp.mpf(n),mp.mpf(pr)
    return mp.mpf(1) if k>=n else ibeta(1-pr, n-k, k+1)
def _b_sf(k,n,pr):
    k,n,pr=mp.mpf(k),mp.mpf(n),mp.mpf(pr)
    return mp.mpf(0) if k>=n else ibeta(pr, k+1, n-k)
def _b_inv(prob,n,pr):
    prob=mp.mpf(prob); lo,hi=-1,int(n)
    while hi-lo>1:
        m=(lo+hi)//2
        if _b_cdf(m,n,pr)>=prob: hi=m
        else: lo=m
    return mp.mpf(hi)
def _p_pmf(k,lam): k,lam=mp.mpf(k),mp.mpf(lam); return mp.e**(k*mp.log(lam)-lam-mp.loggamma(k+1))
def _p_log(k,lam): k,lam=mp.mpf(k),mp.mpf(lam); return k*mp.log(lam)-lam-mp.loggamma(k+1)
def _p_cdf(k,lam): return mp.gammainc(mp.mpf(k)+1, mp.mpf(lam), mp.inf, regularized=True)
def _p_sf(k,lam):  return mp.gammainc(mp.mpf(k)+1, 0, mp.mpf(lam), regularized=True)
def _p_inv(prob,lam):
    prob=mp.mpf(prob); lo=-1; hi=int(mp.floor(mp.mpf(lam)+12*mp.sqrt(mp.mpf(lam))+40))
    while hi-lo>1:
        m=(lo+hi)//2
        if _p_cdf(m,lam)>=prob: hi=m
        else: lo=m
    return mp.mpf(hi)
def _g_pmf(k,pr): k,pr=mp.mpf(k),mp.mpf(pr); return pr*(1-pr)**k
def _g_log(k,pr): k,pr=mp.mpf(k),mp.mpf(pr); return mp.log(pr)+k*mp.log(1-pr)
def _g_cdf(k,pr): k,pr=mp.mpf(k),mp.mpf(pr); return 1-(1-pr)**(k+1)
def _g_sf(k,pr):  k,pr=mp.mpf(k),mp.mpf(pr); return (1-pr)**(k+1)
def _g_inv(prob,pr):
    prob,pr=mp.mpf(prob),mp.mpf(pr); k=int(mp.ceil(mp.log(1-prob)/mp.log(1-pr)-1))
    if k<0: k=0
    while k>0 and _g_cdf(k-1,pr)>=prob: k-=1
    while _g_cdf(k,pr)<prob: k+=1
    return mp.mpf(k)

def _discrete_holdout_rows():
    out=[]
    def R(fn,kern,a1,a2,a3,ref): out.append(row(fn,kern,a1,a2,a3,ref,"all"))
    import math as _m
    # Binomial: fresh (n,p) not in the main grid {20,1000,1e5,1e6,1e7}x{.02,.5,.9}
    for n,pr in [(50,0.1),(50,0.75),(5000,0.35),(500000,0.1),(5000000,0.75)]:
        sd=_m.sqrt(n*pr*(1-pr)); k=int(min(n, _m.floor(n*pr+2*sd)))
        R("Binomial_PMF","K_STATS_Binomial_PMF",k,n,pr,_b_pmf(k,n,pr))
        R("Binomial_LogPMF","K_STATS_Binomial_LogPMF",k,n,pr,_b_log(k,n,pr))
        R("Binomial_Cumulative","K_STATS_Binomial_Cumulative",k,n,pr,_b_cdf(k,n,pr))
        R("Binomial_Survival","K_STATS_Binomial_Survival",k,n,pr,_b_sf(k,n,pr))
        for prob in [0.1,0.9]:
            R("Binomial_InverseCumulative","K_STATS_Binomial_InverseCumulative",prob,n,pr,_b_inv(prob,n,pr))
        R("Binomial_Mean","K_STATS_Binomial_Mean",n,pr,"",mp.mpf(n)*pr)
        R("Binomial_Variance","K_STATS_Binomial_Variance",n,pr,"",mp.mpf(n)*pr*(1-pr))
        R("Binomial_StdDev","K_STATS_Binomial_StdDev",n,pr,"",mp.sqrt(mp.mpf(n)*pr*(1-pr)))
    # Poisson: fresh mean not in {3,50,1000,1e6}
    for lam in [10,200,100000]:
        sd=_m.sqrt(lam); k=int(_m.floor(lam+2*sd))
        R("Poisson_PMF","K_STATS_Poisson_PMF",k,lam,"",_p_pmf(k,lam))
        R("Poisson_LogPMF","K_STATS_Poisson_LogPMF",k,lam,"",_p_log(k,lam))
        R("Poisson_Cumulative","K_STATS_Poisson_Cumulative",k,lam,"",_p_cdf(k,lam))
        R("Poisson_Survival","K_STATS_Poisson_Survival",k,lam,"",_p_sf(k,lam))
        for prob in [0.1,0.9]:
            R("Poisson_InverseCumulative","K_STATS_Poisson_InverseCumulative",prob,lam,"",_p_inv(prob,lam))
        R("Poisson_Mean","K_STATS_Poisson_Mean",lam,"","",mp.mpf(lam))
        R("Poisson_Variance","K_STATS_Poisson_Variance",lam,"","",mp.mpf(lam))
        R("Poisson_StdDev","K_STATS_Poisson_StdDev",lam,"","",mp.sqrt(mp.mpf(lam)))
    # Geometric: fresh p not in {.5,.05,.001,1e-6}
    for pr in [0.2,0.01,1e-4]:
        mean=(1-pr)/pr; k=int(_m.floor(mean))
        R("Geometric_PMF","K_STATS_Geometric_PMF",k,pr,"",_g_pmf(k,pr))
        R("Geometric_LogPMF","K_STATS_Geometric_LogPMF",k,pr,"",_g_log(k,pr))
        R("Geometric_Cumulative","K_STATS_Geometric_Cumulative",k,pr,"",_g_cdf(k,pr))
        R("Geometric_Survival","K_STATS_Geometric_Survival",k,pr,"",_g_sf(k,pr))
        for prob in [0.25,0.9]:
            R("Geometric_InverseCumulative","K_STATS_Geometric_InverseCumulative",prob,pr,"",_g_inv(prob,pr))
        R("Geometric_Mean","K_STATS_Geometric_Mean",pr,"","",(1-mp.mpf(pr))/mp.mpf(pr))
        R("Geometric_Variance","K_STATS_Geometric_Variance",pr,"","",(1-mp.mpf(pr))/mp.mpf(pr)**2)
        R("Geometric_StdDev","K_STATS_Geometric_StdDev",pr,"","",mp.sqrt(1-mp.mpf(pr))/mp.mpf(pr))
    return out

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
    rows += _discrete_holdout_rows()
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
