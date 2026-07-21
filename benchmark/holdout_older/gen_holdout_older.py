"""
Consolidated retroactive holdout for the older families (F-02).

Fresh, off-compliance-grid inputs for every family whose contracts predate the
independent-holdout discipline. References are mpmath at 40 dps, matching each
function's EXACT parameterization (Gamma/Weibull = scale; Exponential = rate;
survival = upper tail; inverse = quantile). The point of the fresh inputs is that
they were never used to calibrate the thresholds, so passing here is genuine
out-of-sample validation before flipping provenance to "validated and frozen".
"""
import csv
import mpmath as mp
mp.mp.dps = 40

def ncdf(z): return mp.ncdf(z)
def npdf(z): return mp.npdf(z)
def ninv(p): return mp.sqrt(2) * mp.erfinv(2*mp.mpf(p) - 1)

def betainc_reg(a, b, x): return mp.betainc(a, b, 0, x, regularized=True)
def gammainc_reg(a, x): return mp.gammainc(a, 0, x, regularized=True)

R = []  # rows
def add(fn, regime, a1, a2, a3, ref):
    R.append({"function": fn, "vba_kernel": f"K_STATS_{fn}", "claim": "holdout", "metric": "rel",
              "arg1": mp.nstr(mp.mpf(a1), 17), "arg2": "" if a2 is None else mp.nstr(mp.mpf(a2), 17),
              "arg3": "" if a3 is None else mp.nstr(mp.mpf(a3), 17),
              "reference": mp.nstr(ref, 30), "observed_vba": "", "regime": regime,
              "evidence_set": "holdout_older"})

# ---- Normal (Mean, StdDev); fresh mu/sigma ----
for x, mu, sd in [("1.35","0.4","1.7"), ("-2.1","0.4","1.7"), ("5.2","1.1","0.8")]:
    z=(mp.mpf(x)-mp.mpf(mu))/mp.mpf(sd)
    add("Normal_Cumulative","all",x,mu,sd, ncdf(z))
    add("Normal_Survival","all",x,mu,sd, ncdf(-z))
add("Normal_Density","all","1.35","0.4","1.7", npdf((mp.mpf('1.35')-mp.mpf('0.4'))/mp.mpf('1.7'))/mp.mpf('1.7'))
for p, mu, sd in [("0.137","0.4","1.7"), ("0.9973","1.1","0.8")]:
    add("Normal_InverseCumulative","all",p,mu,sd, mp.mpf(mu)+mp.mpf(sd)*ninv(p))
    add("Normal_InverseSurvival","all",p,mu,sd, mp.mpf(mu)+mp.mpf(sd)*ninv(1-mp.mpf(p)))
add("Normal_ZScore","all","5.2","1.1","0.8", (mp.mpf('5.2')-mp.mpf('1.1'))/mp.mpf('0.8'))

# ---- NormalStandard (Z) ----
for z in ["0.73","-1.85","3.1"]:
    add("NormalStandard_Cumulative","all",z,None,None, ncdf(mp.mpf(z)))
    add("NormalStandard_Survival","all",z,None,None, ncdf(-mp.mpf(z)))
    add("NormalStandard_Density","all",z,None,None, npdf(mp.mpf(z)))
for p in ["0.137","0.9973","0.0021"]:
    add("NormalStandard_InverseCumulative","all",p,None,None, ninv(p))
    add("NormalStandard_InverseCumulativeFast","all",p,None,None, ninv(p))
    add("NormalStandard_InverseSurvival","all",p,None,None, ninv(1-mp.mpf(p)))
add("NormalStandard_IntervalProbability","all","-1.85","0.73",None, ncdf(mp.mpf('0.73'))-ncdf(mp.mpf('-1.85')))

# ---- Lognormal (MeanLog, StdDevLog) ----
for x, mu, sd in [("2.3","0.2","0.6"), ("0.45","0.2","0.6"), ("12.0","0.9","0.4")]:
    z=(mp.log(mp.mpf(x))-mp.mpf(mu))/mp.mpf(sd)
    add("Lognormal_Cumulative","all",x,mu,sd, ncdf(z))
    add("Lognormal_Survival","all",x,mu,sd, ncdf(-z))
add("Lognormal_Density","all","2.3","0.2","0.6",
    npdf((mp.log(mp.mpf('2.3'))-mp.mpf('0.2'))/mp.mpf('0.6'))/(mp.mpf('2.3')*mp.mpf('0.6')))
for p, mu, sd in [("0.137","0.2","0.6"), ("0.9973","0.9","0.4")]:
    add("Lognormal_InverseCumulative","all",p,mu,sd, mp.e**(mp.mpf(mu)+mp.mpf(sd)*ninv(p)))
    add("Lognormal_InverseSurvival","all",p,mu,sd, mp.e**(mp.mpf(mu)+mp.mpf(sd)*ninv(1-mp.mpf(p))))
add("Lognormal_Mean","all","0.2","0.6",None, mp.e**(mp.mpf('0.2')+mp.mpf('0.6')**2/2))
add("Lognormal_Variance","all","0.2","0.6",None, (mp.e**(mp.mpf('0.6')**2)-1)*mp.e**(2*mp.mpf('0.2')+mp.mpf('0.6')**2))
add("Lognormal_StdDev","all","0.2","0.6",None, mp.sqrt((mp.e**(mp.mpf('0.6')**2)-1))*mp.e**(mp.mpf('0.2')+mp.mpf('0.6')**2/2))

# ---- StudentT (df) ----
def t_cdf(x, df):
    x, df = mp.mpf(x), mp.mpf(df)
    ib = betainc_reg(df/2, mp.mpf(1)/2, df/(df+x*x))
    return 1 - ib/2 if x > 0 else ib/2
def t_pdf(x, df):
    x, df = mp.mpf(x), mp.mpf(df)
    return mp.gamma((df+1)/2)/(mp.sqrt(df*mp.pi)*mp.gamma(df/2))*(1+x*x/df)**(-(df+1)/2)
for x, df in [("1.4","7"), ("-2.3","12"), ("3.1","4")]:
    add("StudentT_Cumulative","all",x,df,None, t_cdf(x,df))
    add("StudentT_Survival","all",x,df,None, 1-t_cdf(x,df))
add("StudentT_Density","all","1.4","7",None, t_pdf("1.4","7"))
add("StudentT_InverseCumulative","all","0.9973","7",None, mp.findroot(lambda t: t_cdf(t,7)-mp.mpf('0.9973'), 2))

# ---- ChiSquare (df) ----
def chi_cdf(x, df): return gammainc_reg(mp.mpf(df)/2, mp.mpf(x)/2)
for x, df in [("6.3","5"), ("18.7","10"), ("2.1","3")]:
    add("ChiSquare_Cumulative","all",x,df,None, chi_cdf(x,df))
    add("ChiSquare_Survival","all",x,df,None, 1-chi_cdf(x,df))
add("ChiSquare_InverseCumulative","all","0.9973","5",None, mp.findroot(lambda t: chi_cdf(t,5)-mp.mpf('0.9973'), 10))

# ---- F (df1, df2) small-df validated regime ----
def f_cdf(x, d1, d2):
    x,d1,d2=mp.mpf(x),mp.mpf(d1),mp.mpf(d2)
    return betainc_reg(d1/2, d2/2, d1*x/(d1*x+d2))
for x,d1,d2 in [("2.1","6","14"), ("0.7","9","9")]:
    add("F_Cumulative","validated",x,d1,d2, f_cdf(x,d1,d2))
    add("F_Survival","validated",x,d1,d2, 1-f_cdf(x,d1,d2))

# ---- Gamma (Shape, Scale) ----
def g_cdf(x, k, th): return gammainc_reg(mp.mpf(k), mp.mpf(x)/mp.mpf(th))
def g_pdf(x, k, th):
    x,k,th=mp.mpf(x),mp.mpf(k),mp.mpf(th)
    return x**(k-1)*mp.e**(-x/th)/(th**k*mp.gamma(k))
for x,k,th in [("3.4","2.7","1.3"), ("0.6","2.7","1.3"), ("9.0","4.3","0.8")]:
    add("Gamma_Cumulative","all",x,k,th, g_cdf(x,k,th))
    add("Gamma_Survival","all",x,k,th, 1-g_cdf(x,k,th))
add("Gamma_Density","all","3.4","2.7","1.3", g_pdf("3.4","2.7","1.3"))
add("Gamma_InverseCumulative","all","0.9973","2.7","1.3", mp.findroot(lambda t: g_cdf(t,"2.7","1.3")-mp.mpf('0.9973'), 5))
add("Gamma_Mean","all","2.7","1.3",None, mp.mpf('2.7')*mp.mpf('1.3'))
add("Gamma_Variance","all","2.7","1.3",None, mp.mpf('2.7')*mp.mpf('1.3')**2)
add("Gamma_StdDev","all","2.7","1.3",None, mp.sqrt(mp.mpf('2.7'))*mp.mpf('1.3'))

# ---- Beta (Alpha, Beta) balanced ----
def b_cdf(x,a,b): return betainc_reg(mp.mpf(a),mp.mpf(b),mp.mpf(x))
def b_pdf(x,a,b):
    x,a,b=mp.mpf(x),mp.mpf(a),mp.mpf(b)
    return x**(a-1)*(1-x)**(b-1)/mp.beta(a,b)
for x,a,b in [("0.35","2.7","3.4"), ("0.82","2.7","3.4"), ("0.15","4.3","1.9")]:
    add("Beta_Cumulative","balanced",x,a,b, b_cdf(x,a,b))
    add("Beta_Survival","balanced",x,a,b, 1-b_cdf(x,a,b))
add("Beta_Density","balanced","0.35","2.7","3.4", b_pdf("0.35","2.7","3.4"))
add("Beta_InverseCumulative","balanced","0.9973","2.7","3.4", mp.findroot(lambda t: b_cdf(t,"2.7","3.4")-mp.mpf('0.9973'), mp.mpf('0.7')))
add("Beta_Mean","all","2.7","3.4",None, mp.mpf('2.7')/(mp.mpf('2.7')+mp.mpf('3.4')))
add("Beta_Variance","all","2.7","3.4",None, mp.mpf('2.7')*mp.mpf('3.4')/((mp.mpf('2.7')+mp.mpf('3.4'))**2*(mp.mpf('2.7')+mp.mpf('3.4')+1)))
add("Beta_StdDev","all","2.7","3.4",None, mp.sqrt(mp.mpf('2.7')*mp.mpf('3.4')/((mp.mpf('2.7')+mp.mpf('3.4'))**2*(mp.mpf('2.7')+mp.mpf('3.4')+1))))

# ---- Exponential (rate Lambda) ----
for x, lam in [("0.9","1.5"), ("4.2","0.3")]:
    add("Exponential_Cumulative","all",x,lam,None, 1-mp.e**(-mp.mpf(lam)*mp.mpf(x)))
    add("Exponential_Survival","all",x,lam,None, mp.e**(-mp.mpf(lam)*mp.mpf(x)))
add("Exponential_Density","all","0.9","1.5",None, mp.mpf('1.5')*mp.e**(-mp.mpf('1.5')*mp.mpf('0.9')))
add("Exponential_InverseCumulative","all","0.9973","1.5",None, -mp.log(1-mp.mpf('0.9973'))/mp.mpf('1.5'))

# ---- Weibull (Shape, Scale) ----
def w_cdf(x,k,s): return 1-mp.e**(-(mp.mpf(x)/mp.mpf(s))**mp.mpf(k))
for x,k,s in [("1.8","1.6","2.2"), ("5.0","1.6","2.2")]:
    add("Weibull_Cumulative","all",x,k,s, w_cdf(x,k,s))
    add("Weibull_Survival","all",x,k,s, mp.e**(-(mp.mpf(x)/mp.mpf(s))**mp.mpf(k)))
add("Weibull_Density","all","1.8","1.6","2.2",
    (mp.mpf('1.6')/mp.mpf('2.2'))*(mp.mpf('1.8')/mp.mpf('2.2'))**(mp.mpf('1.6')-1)*mp.e**(-(mp.mpf('1.8')/mp.mpf('2.2'))**mp.mpf('1.6')))
add("Weibull_InverseCumulative","all","0.9973","1.6","2.2", mp.mpf('2.2')*(-mp.log(1-mp.mpf('0.9973')))**(1/mp.mpf('1.6')))
add("Weibull_Mean","all","1.6","2.2",None, mp.mpf('2.2')*mp.gamma(1+1/mp.mpf('1.6')))
add("Weibull_Variance","all","1.6","2.2",None, mp.mpf('2.2')**2*(mp.gamma(1+2/mp.mpf('1.6'))-mp.gamma(1+1/mp.mpf('1.6'))**2))
add("Weibull_StdDev","all","1.6","2.2",None, mp.mpf('2.2')*mp.sqrt(mp.gamma(1+2/mp.mpf('1.6'))-mp.gamma(1+1/mp.mpf('1.6'))**2))

# ---- Uniform (a, b) ----
for x,a,b in [("2.5","1.0","4.0"), ("3.7","1.0","4.0")]:
    add("Uniform_Cumulative","all",x,a,b, (mp.mpf(x)-mp.mpf(a))/(mp.mpf(b)-mp.mpf(a)))
    add("Uniform_Survival","all",x,a,b, (mp.mpf(b)-mp.mpf(x))/(mp.mpf(b)-mp.mpf(a)))
add("Uniform_Density","all","2.5","1.0","4.0", 1/(mp.mpf('4.0')-mp.mpf('1.0')))
add("Uniform_InverseCumulative","all","0.137","1.0","4.0", mp.mpf('1.0')+mp.mpf('0.137')*(mp.mpf('4.0')-mp.mpf('1.0')))


# ---- Special-function kernels (PROB_*) ----
def add_sf(fn, a1, a2, ref):
    R.append({"function": fn, "vba_kernel": f"PROB_{fn}", "claim": "holdout", "metric": "rel",
              "arg1": mp.nstr(mp.mpf(a1),17), "arg2": "" if a2 is None else mp.nstr(mp.mpf(a2),17),
              "arg3": "", "reference": mp.nstr(ref,30), "observed_vba": "", "regime": "all",
              "evidence_set": "holdout_older"})
for z in ["3.7","12.3","0.65"]:
    add_sf("LogGamma", z, None, mp.loggamma(mp.mpf(z)))
for n,k in [("40","12"),("100","7")]:
    add_sf("LogChoose", n, k, mp.loggamma(mp.mpf(n)+1)-mp.loggamma(mp.mpf(k)+1)-mp.loggamma(mp.mpf(n)-mp.mpf(k)+1))
for z in ["3.7","250000"]:
    add_sf("LogGammaHalfDiff", z, None, mp.loggamma(mp.mpf(z)+mp.mpf('0.5'))-mp.loggamma(mp.mpf(z)))
for n in ["2.3","7.8"]:
    nn=mp.mpf(n)
    add_sf("StirlingError", n, None, mp.loggamma(nn+1)-(nn*mp.log(nn)-nn+mp.mpf('0.5')*mp.log(2*mp.pi*nn)))

fields=["function","vba_kernel","claim","metric","arg1","arg2","arg3","reference","observed_vba","regime","evidence_set"]
with open("holdout_older_grid.csv","w",newline="") as f:
    w=csv.DictWriter(f,fieldnames=fields); w.writeheader(); w.writerows(R)
print(f"wrote holdout_older_grid.csv: {len(R)} rows across {len(set(r['function'] for r in R))} functions")
