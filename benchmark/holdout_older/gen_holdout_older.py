"""
Consolidated retroactive holdout for the older families (F-02).

Fresh, off-compliance-grid inputs for every family whose contracts predate the
independent-holdout discipline. References are mpmath at 40 dps, matching each
function's EXACT parameterization (Gamma/Weibull = scale; Exponential = rate;
survival = upper tail; inverse = quantile). The point of the fresh inputs is that
they were never used to calibrate the thresholds, so passing here is genuine
out-of-sample validation before flipping provenance to "validated and frozen".

Binary64 arguments (#159). Every reference is evaluated at the exact Double its
argument token parses to (M below), which is what the exporter passes, and
every survival is computed directly as an upper tail, never as 1 - CDF. Before
#159 the references were taken at the decimal token ("1.35" as an exact
decimal), which cost the survival rows 15 to 17 of their 30 printed digits.
The token written to the grid is unchanged: it is still the decimal design value.

The committed grid holds Excel observations. Re-reference it, changing
reference cells only, with

    python gen_holdout_older.py --merge-references          # report
    python gen_holdout_older.py --merge-references --write  # patch

The merge holds back every row of a function that the #72 normal-tail study
reads from this grid (benchmark/normal_tail_study/validate_kernel.py: its P1 to
P4 functions). That study's record binds this grid and records each row's error
against the committed reference, so changing those references changes recorded
study values and is a decision for that study, not for this script.
"""
import argparse
import ast
import csv
import os
import sys

import mpmath as mp

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.dirname(HERE))
from _binary64 import exact_binary64_mpf  # noqa: E402
from _ibeta import f_sf, ibeta, ibeta_upper  # noqa: E402

mp.mp.dps = 40


def M(t):
    """The exact Double an argument token denotes; other values pass through."""
    return exact_binary64_mpf(t) if isinstance(t, str) else mp.mpf(t)


def ncdf(z): return mp.ncdf(z)
def npdf(z): return mp.npdf(z)
def ninv(p): return mp.sqrt(2) * mp.erfinv(2*M(p) - 1)

def betainc_reg(a, b, x): return mp.betainc(a, b, 0, x, regularized=True)
def gammainc_reg(a, x): return mp.gammainc(a, 0, x, regularized=True)
def gammainc_upper(a, x): return mp.gammainc(a, x, mp.inf, regularized=True)

R = []  # rows
def add(fn, regime, a1, a2, a3, ref):
    R.append({"function": fn, "vba_kernel": f"K_STATS_{fn}", "claim": "holdout", "metric": "rel",
              "arg1": mp.nstr(mp.mpf(a1), 17), "arg2": "" if a2 is None else mp.nstr(mp.mpf(a2), 17),
              "arg3": "" if a3 is None else mp.nstr(mp.mpf(a3), 17),
              "reference": mp.nstr(ref, 30), "observed_vba": "", "regime": regime,
              "evidence_set": "holdout_older"})

# The committed grid has no survival row at these two deep-tail points: they were
# removed in ecf3e45 (z = 5.1 and 3.96, beyond the survival contracts' central
# domain; see SurvivalTailRel in numerical_limitations.csv). The cumulative rows
# remain.
RETIRED_SURVIVAL = {("Normal_Survival", "5.2"), ("Lognormal_Survival", "12.0")}

# ---- Normal (Mean, StdDev); fresh mu/sigma ----
for x, mu, sd in [("1.35","0.4","1.7"), ("-2.1","0.4","1.7"), ("5.2","1.1","0.8")]:
    z=(M(x)-M(mu))/M(sd)
    add("Normal_Cumulative","all",x,mu,sd, ncdf(z))
    if ("Normal_Survival", x) not in RETIRED_SURVIVAL:
        add("Normal_Survival","all",x,mu,sd, ncdf(-z))
add("Normal_Density","all","1.35","0.4","1.7", npdf((M('1.35')-M('0.4'))/M('1.7'))/M('1.7'))
for p, mu, sd in [("0.137","0.4","1.7"), ("0.9973","1.1","0.8")]:
    add("Normal_InverseCumulative","all",p,mu,sd, M(mu)+M(sd)*ninv(p))
    add("Normal_InverseSurvival","all",p,mu,sd, M(mu)+M(sd)*ninv(1-M(p)))
add("Normal_ZScore","all","5.2","1.1","0.8", (M('5.2')-M('1.1'))/M('0.8'))

# ---- NormalStandard (Z) ----
for z in ["0.73","-1.85","3.1"]:
    add("NormalStandard_Cumulative","all",z,None,None, ncdf(M(z)))
    add("NormalStandard_Survival","all",z,None,None, ncdf(-M(z)))
    add("NormalStandard_Density","all",z,None,None, npdf(M(z)))
for p in ["0.137","0.9973","0.0021"]:
    add("NormalStandard_InverseCumulative","all",p,None,None, ninv(p))
    add("NormalStandard_InverseCumulativeFast","all",p,None,None, ninv(p))
    add("NormalStandard_InverseSurvival","all",p,None,None, ninv(1-M(p)))
add("NormalStandard_IntervalProbability","all","-1.85","0.73",None, ncdf(M('0.73'))-ncdf(M('-1.85')))

# ---- Lognormal (MeanLog, StdDevLog) ----
for x, mu, sd in [("2.3","0.2","0.6"), ("0.45","0.2","0.6"), ("12.0","0.9","0.4")]:
    z=(mp.log(M(x))-M(mu))/M(sd)
    add("Lognormal_Cumulative","all",x,mu,sd, ncdf(z))
    if ("Lognormal_Survival", x) not in RETIRED_SURVIVAL:
        add("Lognormal_Survival","all",x,mu,sd, ncdf(-z))
add("Lognormal_Density","all","2.3","0.2","0.6",
    npdf((mp.log(M('2.3'))-M('0.2'))/M('0.6'))/(M('2.3')*M('0.6')))
for p, mu, sd in [("0.137","0.2","0.6"), ("0.9973","0.9","0.4")]:
    add("Lognormal_InverseCumulative","all",p,mu,sd, mp.e**(M(mu)+M(sd)*ninv(p)))
    add("Lognormal_InverseSurvival","all",p,mu,sd, mp.e**(M(mu)+M(sd)*ninv(1-M(p))))
add("Lognormal_Mean","all","0.2","0.6",None, mp.e**(M('0.2')+M('0.6')**2/2))
add("Lognormal_Variance","all","0.2","0.6",None, (mp.e**(M('0.6')**2)-1)*mp.e**(2*M('0.2')+M('0.6')**2))
add("Lognormal_StdDev","all","0.2","0.6",None, mp.sqrt((mp.e**(M('0.6')**2)-1))*mp.e**(M('0.2')+M('0.6')**2/2))

# ---- StudentT (df) ----
def t_cdf(x, df):
    x, df = M(x), M(df)
    ib = betainc_reg(df/2, mp.mpf(1)/2, df/(df+x*x))
    return 1 - ib/2 if x > 0 else ib/2
def t_sf(x, df):
    """Upper tail, direct: I_z(df/2, 1/2) / 2 for x > 0."""
    x, df = M(x), M(df)
    h = ibeta(df/(df+x*x), df/2, mp.mpf(1)/2)/2
    return h if x > 0 else 1 - h
def t_pdf(x, df):
    x, df = M(x), M(df)
    return mp.gamma((df+1)/2)/(mp.sqrt(df*mp.pi)*mp.gamma(df/2))*(1+x*x/df)**(-(df+1)/2)
for x, df in [("1.4","7"), ("-2.3","12"), ("3.1","4")]:
    add("StudentT_Cumulative","all",x,df,None, t_cdf(x,df))
    add("StudentT_Survival","all",x,df,None, t_sf(x,df))
add("StudentT_Density","all","1.4","7",None, t_pdf("1.4","7"))
add("StudentT_InverseCumulative","all","0.9973","7",None, mp.findroot(lambda t: t_cdf(t,7)-M('0.9973'), 2))

# ---- ChiSquare (df) ----
def chi_cdf(x, df): return gammainc_reg(M(df)/2, M(x)/2)
def chi_sf(x, df): return gammainc_upper(M(df)/2, M(x)/2)
for x, df in [("6.3","5"), ("18.7","10"), ("2.1","3")]:
    add("ChiSquare_Cumulative","all",x,df,None, chi_cdf(x,df))
    add("ChiSquare_Survival","all",x,df,None, chi_sf(x,df))
add("ChiSquare_InverseCumulative","all","0.9973","5",None, mp.findroot(lambda t: chi_cdf(t,5)-M('0.9973'), 10))

# ---- F (df1, df2) small-df validated regime ----
def f_cdf(x, d1, d2):
    x,d1,d2=M(x),M(d1),M(d2)
    return betainc_reg(d1/2, d2/2, d1*x/(d1*x+d2))
for x,d1,d2 in [("2.1","6","14"), ("0.7","9","9")]:
    add("F_Cumulative","validated",x,d1,d2, f_cdf(x,d1,d2))
    add("F_Survival","validated",x,d1,d2, f_sf(M(x),M(d1),M(d2)))

# ---- Gamma (Shape, Scale) ----
def g_cdf(x, k, th): return gammainc_reg(M(k), M(x)/M(th))
def g_sf(x, k, th): return gammainc_upper(M(k), M(x)/M(th))
def g_pdf(x, k, th):
    x,k,th=M(x),M(k),M(th)
    return x**(k-1)*mp.e**(-x/th)/(th**k*mp.gamma(k))
for x,k,th in [("3.4","2.7","1.3"), ("0.6","2.7","1.3"), ("9.0","4.3","0.8")]:
    add("Gamma_Cumulative","all",x,k,th, g_cdf(x,k,th))
    add("Gamma_Survival","all",x,k,th, g_sf(x,k,th))
add("Gamma_Density","all","3.4","2.7","1.3", g_pdf("3.4","2.7","1.3"))
add("Gamma_InverseCumulative","all","0.9973","2.7","1.3", mp.findroot(lambda t: g_cdf(t,"2.7","1.3")-M('0.9973'), 5))
add("Gamma_Mean","all","2.7","1.3",None, M('2.7')*M('1.3'))
add("Gamma_Variance","all","2.7","1.3",None, M('2.7')*M('1.3')**2)
add("Gamma_StdDev","all","2.7","1.3",None, mp.sqrt(M('2.7'))*M('1.3'))

# ---- Beta (Alpha, Beta) balanced ----
def b_cdf(x,a,b): return betainc_reg(M(a),M(b),M(x))
def b_sf(x,a,b): return ibeta_upper(M(x),M(a),M(b))
def b_pdf(x,a,b):
    x,a,b=M(x),M(a),M(b)
    return x**(a-1)*(1-x)**(b-1)/mp.beta(a,b)
for x,a,b in [("0.35","2.7","3.4"), ("0.82","2.7","3.4"), ("0.15","4.3","1.9")]:
    add("Beta_Cumulative","balanced",x,a,b, b_cdf(x,a,b))
    add("Beta_Survival","balanced",x,a,b, b_sf(x,a,b))
add("Beta_Density","balanced","0.35","2.7","3.4", b_pdf("0.35","2.7","3.4"))
add("Beta_InverseCumulative","balanced","0.9973","2.7","3.4", mp.findroot(lambda t: b_cdf(t,"2.7","3.4")-M('0.9973'), mp.mpf('0.7')))
add("Beta_Mean","all","2.7","3.4",None, M('2.7')/(M('2.7')+M('3.4')))
add("Beta_Variance","all","2.7","3.4",None, M('2.7')*M('3.4')/((M('2.7')+M('3.4'))**2*(M('2.7')+M('3.4')+1)))
add("Beta_StdDev","all","2.7","3.4",None, mp.sqrt(M('2.7')*M('3.4')/((M('2.7')+M('3.4'))**2*(M('2.7')+M('3.4')+1))))

# ---- Exponential (rate Lambda) ----
for x, lam in [("0.9","1.5"), ("4.2","0.3")]:
    add("Exponential_Cumulative","all",x,lam,None, 1-mp.e**(-M(lam)*M(x)))
    add("Exponential_Survival","all",x,lam,None, mp.e**(-M(lam)*M(x)))
add("Exponential_Density","all","0.9","1.5",None, M('1.5')*mp.e**(-M('1.5')*M('0.9')))
add("Exponential_InverseCumulative","all","0.9973","1.5",None, -mp.log(1-M('0.9973'))/M('1.5'))

# ---- Weibull (Shape, Scale) ----
def w_cdf(x,k,s): return 1-mp.e**(-(M(x)/M(s))**M(k))
for x,k,s in [("1.8","1.6","2.2"), ("5.0","1.6","2.2")]:
    add("Weibull_Cumulative","all",x,k,s, w_cdf(x,k,s))
    add("Weibull_Survival","all",x,k,s, mp.e**(-(M(x)/M(s))**M(k)))
add("Weibull_Density","all","1.8","1.6","2.2",
    (M('1.6')/M('2.2'))*(M('1.8')/M('2.2'))**(M('1.6')-1)*mp.e**(-(M('1.8')/M('2.2'))**M('1.6')))
add("Weibull_InverseCumulative","all","0.9973","1.6","2.2", M('2.2')*(-mp.log(1-M('0.9973')))**(1/M('1.6')))
add("Weibull_Mean","all","1.6","2.2",None, M('2.2')*mp.gamma(1+1/M('1.6')))
add("Weibull_Variance","all","1.6","2.2",None, M('2.2')**2*(mp.gamma(1+2/M('1.6'))-mp.gamma(1+1/M('1.6'))**2))
add("Weibull_StdDev","all","1.6","2.2",None, M('2.2')*mp.sqrt(mp.gamma(1+2/M('1.6'))-mp.gamma(1+1/M('1.6'))**2))

# ---- Uniform (a, b) ----
for x,a,b in [("2.5","1.0","4.0"), ("3.7","1.0","4.0")]:
    add("Uniform_Cumulative","all",x,a,b, (M(x)-M(a))/(M(b)-M(a)))
    add("Uniform_Survival","all",x,a,b, (M(b)-M(x))/(M(b)-M(a)))
add("Uniform_Density","all","2.5","1.0","4.0", 1/(M('4.0')-M('1.0')))
add("Uniform_InverseCumulative","all","0.137","1.0","4.0", M('1.0')+M('0.137')*(M('4.0')-M('1.0')))


# ---- Special-function kernels (PROB_*) ----
def add_sf(fn, a1, a2, ref):
    R.append({"function": fn, "vba_kernel": f"PROB_{fn}", "claim": "holdout", "metric": "rel",
              "arg1": mp.nstr(mp.mpf(a1),17), "arg2": "" if a2 is None else mp.nstr(mp.mpf(a2),17),
              "arg3": "", "reference": mp.nstr(ref,30), "observed_vba": "", "regime": "all",
              "evidence_set": "holdout_older"})
for z in ["3.7","12.3","0.65"]:
    add_sf("LogGamma", z, None, mp.loggamma(M(z)))
for n,k in [("40","12"),("100","7")]:
    add_sf("LogChoose", n, k, mp.loggamma(M(n)+1)-mp.loggamma(M(k)+1)-mp.loggamma(M(n)-M(k)+1))
for z in ["3.7","250000"]:
    add_sf("LogGammaHalfDiff", z, None, mp.loggamma(M(z)+mp.mpf('0.5'))-mp.loggamma(M(z)))
for n in ["2.3","7.8"]:
    nn=M(n)
    add_sf("StirlingError", n, None, mp.loggamma(nn+1)-(nn*mp.log(nn)-nn+mp.mpf('0.5')*mp.log(2*mp.pi*nn)))

fields=["function","vba_kernel","claim","metric","arg1","arg2","arg3","reference","observed_vba","regime","evidence_set"]


def normal_tail_functions():
    """The functions the #72 normal-tail record reads from this grid.

    Parsed from validate_kernel.py rather than imported, so this generator does
    not depend on that study's modules; P1 to P4 are DIRECT, STANDARDISED,
    INTERVALS and INVERSES."""
    path = os.path.join(os.path.dirname(HERE), "normal_tail_study", "validate_kernel.py")
    with open(path, encoding="utf-8") as fh:
        tree = ast.parse(fh.read())
    names = {}
    for node in tree.body:
        if isinstance(node, ast.Assign) and len(node.targets) == 1 \
                and isinstance(node.targets[0], ast.Name):
            names[node.targets[0].id] = node.value
    out = set()
    for key in ("DIRECT", "STANDARDISED", "INTERVALS"):
        out |= set(ast.literal_eval(names[key]))
    for fam in ("NormalStandard", "Normal", "Lognormal"):     # INVERSES
        for s in ("InverseCumulative", "InverseSurvival"):
            out.add(f"{fam}_{s}")
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--merge-references", action="store_true",
                    help="re-reference the committed grid, keeping every observation")
    ap.add_argument("--write", action="store_true",
                    help="with --merge-references: patch the committed grid")
    a = ap.parse_args()
    grid = os.path.join(HERE, "holdout_older_grid.csv")
    if a.merge_references:
        from _study_merge import merge_references, print_report, study_describe, study_key
        bound = normal_tail_functions()

        def hold(row):
            if row["function"] in bound:
                return ("read by the #72 normal-tail record "
                        "(normal_tail_study/results.json binds this grid)")
            return None

        changes, held = merge_references(grid, R, study_key, {"observed_vba"},
                                         hold=hold, write=a.write, describe=study_describe)
        print_report("holdout_older_grid.csv", changes, held, a.write)
        return
    with open(grid,"w",newline="") as f:
        w=csv.DictWriter(f,fieldnames=fields,lineterminator="\n"); w.writeheader(); w.writerows(R)
    print(f"wrote holdout_older_grid.csv: {len(R)} rows across {len(set(r['function'] for r in R))} functions")


if __name__ == "__main__":
    main()
