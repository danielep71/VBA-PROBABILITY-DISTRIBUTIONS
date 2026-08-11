"""
Build benchmark/excel_comparison/excel_comparison_grid.csv.

WHY THIS STUDY EXISTS
    The README asserts that this library is more accurate than Excel's native
    functions in the deep tails and at large shapes. That claim was never
    demonstrated with numbers. This grid measures Excel and the library side by
    side against 50-digit references, so the claim can be shown rather than
    asserted - or corrected, if Excel turns out to be fine on a point.

WHAT IS AND IS NOT BEING CLAIMED
    Excel's native statistical functions are good across the range most users
    work in, and this grid is NOT a list of Excel bugs. The points below are
    chosen where a structural difference is expected:

      * deep tails, where Excel returns a probability but the API forces it
        through 1 - p, or where no direct survival function exists;
      * large shape or df, where the underlying series lose accuracy;
      * inputs where Excel returns an error and the library returns a value,
        or vice versa.

    Points where the two agree are kept in the grid deliberately. A comparison
    that only showed the wins would not be evidence.

REFERENCES
    mpmath at 50 digits. Continued fractions and series are evaluated directly
    where mpmath's own incomplete functions do not converge at these arguments.
"""
import csv
import mpmath as mp

mp.mp.dps = 50
NONE = "-"          # Excel has no equivalent function
EPS = mp.mpf(10) ** -40


def norm_sf(z):
    return mp.erfc(mp.mpf(z) / mp.sqrt(2)) / 2


def norm_cdf(z):
    return mp.erfc(-mp.mpf(z) / mp.sqrt(2)) / 2


def gser(A, X):
    A = mp.mpf(A); X = mp.mpf(X); Ap = A; S = 1 / A; D = S; n = 0
    while n < 3000000:
        Ap += 1; D *= X / Ap; S += D; n += 1
        if abs(D) <= abs(S) * EPS:
            break
    return S * mp.e ** (-X + A * mp.log(X) - mp.loggamma(A))


def gcf(A, X):
    A = mp.mpf(A); X = mp.mpf(X); tiny = mp.mpf(10) ** -300
    b = X + 1 - A; c = 1 / tiny; d = 1 / b; h = d
    for i in range(1, 300000):
        an = -mp.mpf(i) * (mp.mpf(i) - A); b += 2
        d = an * d + b
        if abs(d) < tiny: d = tiny
        c = b + an / c
        if abs(c) < tiny: c = tiny
        d = 1 / d; de = d * c; h *= de
        if abs(de - 1) <= EPS:
            break
    return h * mp.e ** (-X + A * mp.log(X) - mp.loggamma(A))


def chi_cdf(x, df):
    A = mp.mpf(df) / 2; X = mp.mpf(x) / 2
    return gser(A, X) if X < A + 1 else 1 - gcf(A, X)


def chi_sf(x, df):
    A = mp.mpf(df) / 2; X = mp.mpf(x) / 2
    return gcf(A, X) if X >= A + 1 else 1 - gser(A, X)


def bcf(A, B, X):
    A = mp.mpf(A); B = mp.mpf(B); X = mp.mpf(X); tiny = mp.mpf(10) ** -300
    qab = A + B; qap = A + 1; qam = A - 1
    c = mp.mpf(1); d = 1 - qab * X / qap
    if abs(d) < tiny: d = tiny
    d = 1 / d; h = d
    for m in range(1, 300000):
        m2 = 2 * m
        aa = m * (B - m) * X / ((qam + m2) * (A + m2))
        d = 1 + aa * d
        if abs(d) < tiny: d = tiny
        c = 1 + aa / c
        if abs(c) < tiny: c = tiny
        d = 1 / d; h *= d * c
        aa = -(A + m) * (qab + m) * X / ((A + m2) * (qap + m2))
        d = 1 + aa * d
        if abs(d) < tiny: d = tiny
        c = 1 + aa / c
        if abs(c) < tiny: c = tiny
        d = 1 / d; de = d * c; h *= de
        if abs(de - 1) <= EPS:
            break
    return h


def beta_I(A, B, X):
    A = mp.mpf(A); B = mp.mpf(B); X = mp.mpf(X)
    if X <= 0: return mp.mpf(0)
    if X >= 1: return mp.mpf(1)
    lf = (A * mp.log(X) + B * mp.log1p(-X)
          - (mp.loggamma(A) + mp.loggamma(B) - mp.loggamma(A + B)))
    fr = mp.e ** lf
    if X < (A + 1) / (A + B + 2):
        return fr * bcf(A, B, X) / A
    return 1 - fr * bcf(B, A, 1 - X) / B


def t_sf(t, df):
    df = mp.mpf(df); t = mp.mpf(t)
    x = df / (df + t * t)
    return beta_I(df / 2, mp.mpf("0.5"), x) / 2


def f_sf(f, d1, d2):
    d1 = mp.mpf(d1); d2 = mp.mpf(d2); f = mp.mpf(f)
    x = d1 * f / (d1 * f + d2)
    return 1 - beta_I(d1 / 2, d2 / 2, x)



def gam_cdf(x, k, theta):
    return mp.gammainc(mp.mpf(k), 0, mp.mpf(x) / mp.mpf(theta), regularized=True)


def gam_sf(x, k, theta):
    return mp.gammainc(mp.mpf(k), mp.mpf(x) / mp.mpf(theta), mp.inf, regularized=True)


def beta_cdf(x, a, b):
    return mp.betainc(mp.mpf(a), mp.mpf(b), 0, mp.mpf(x), regularized=True)


def binom_sf(k, n, p):
    """P(X > k) via the beta identity, avoiding a long summation."""
    return beta_cdf(p, mp.mpf(k) + 1, mp.mpf(n) - mp.mpf(k))


def pois_sf(k, lam):
    """P(X > k) = P(gamma(k+1) < lam)."""
    return mp.gammainc(mp.mpf(k) + 1, 0, mp.mpf(lam), regularized=True)


def negbinom_cdf(k, r, p):
    return beta_cdf(mp.mpf(p), mp.mpf(r), mp.mpf(k) + 1)


def hyper_pmf(k, n, K, N):
    k = mp.mpf(k); n = mp.mpf(n); K = mp.mpf(K); N = mp.mpf(N)
    return mp.binomial(K, k) * mp.binomial(N - K, n - k) / mp.binomial(N, n)


CASES = [
    # ---- Normal: no .RT variant in Excel -------------------------------------
    ("normal cdf, body", "NORM.S.DIST(1.5,TRUE)",
     "K_STATS_NormalStandard_Cumulative(1.5)", norm_cdf(1.5),
     "body: both should be exact"),
    ("normal survival, z = 8", "1-NORM.S.DIST(8,TRUE)",
     "K_STATS_NormalStandard_Survival(8)", norm_sf(8),
     "no direct survival in Excel; 1-CDF loses the tail"),
    ("normal survival, z = 15", "1-NORM.S.DIST(15,TRUE)",
     "K_STATS_NormalStandard_Survival(15)", norm_sf(15),
     "1-CDF is exactly 0 in Double here"),
    ("normal inverse survival", NONE,
     "K_STATS_NormalStandard_InverseSurvival(1E-15)",
     -mp.sqrt(2) * mp.erfinv(2 * mp.mpf("1e-15") - 1),
     "Excel has no inverse survival; NORM.S.INV(1-1E-15) loses the argument"),
    # ---- Lognormal: no .RT ---------------------------------------------------
    ("lognormal cdf, body", "LOGNORM.DIST(2,0,1,TRUE)",
     "K_STATS_Lognormal_Cumulative(2,0,1)", norm_cdf(mp.log(2)),
     "body"),
    ("lognormal survival, deep", "1-LOGNORM.DIST(2981,0,1,TRUE)",
     "K_STATS_Lognormal_Survival(2981,0,1)", norm_sf(mp.log(2981)),
     "no direct survival; log(2981) is about 8"),
    # ---- Student t: Excel HAS .RT -------------------------------------------
    ("t survival, body", "T.DIST.RT(1.5,10)",
     "K_STATS_StudentT_Survival(1.5,10)", t_sf(1.5, 10),
     "body; Excel has a direct upper tail here"),
    ("t survival, deep tail", "T.DIST.RT(30,10)",
     "K_STATS_StudentT_Survival(30,10)", t_sf(30, 10),
     "deep tail with a direct Excel function"),
    ("t survival, large df", "T.DIST.RT(2,1E7)",
     "K_STATS_StudentT_Survival(2,1E7)", t_sf(2, 1e7),
     "large df: known library weakness, incomplete-beta CF conditioning"),
    # ---- Chi-square: Excel HAS .RT ------------------------------------------
    ("chi-square sf, deep tail", "CHISQ.DIST.RT(200,3)",
     "K_STATS_ChiSquare_Survival(200,3)", chi_sf(200, 3),
     "deep tail with a direct Excel function"),
    ("chi-square cdf, large df", "CHISQ.DIST(1E6,1E6,TRUE)",
     "K_STATS_ChiSquare_Cumulative(1E6,1E6)", chi_cdf(1e6, 1e6),
     "large df"),
    # ---- F: Excel HAS .RT ----------------------------------------------------
    ("F survival, deep tail", "F.DIST.RT(100,5,10)",
     "K_STATS_F_Survival(100,5,10)", f_sf(100, 5, 10),
     "deep tail with a direct Excel function"),
    ("F survival, extreme ratio", "F.DIST.RT(1,1E6,3)",
     "K_STATS_F_Survival(1,1E6,3)", f_sf(1, 1e6, 3),
     "extreme degree ratio"),
    # ---- Gamma: no .RT -------------------------------------------------------
    ("gamma cdf, body", "GAMMA.DIST(3,2,1,TRUE)",
     "K_STATS_Gamma_Cumulative(3,2,1)", gam_cdf(3, 2, 1),
     "body"),
    ("gamma survival, deep", "1-GAMMA.DIST(60,2,1,TRUE)",
     "K_STATS_Gamma_Survival(60,2,1)", gam_sf(60, 2, 1),
     "no direct survival; true value is about 5E-25"),
    ("gamma cdf, large shape", "GAMMA.DIST(1E6,1E6,1,TRUE)",
     "K_STATS_Gamma_Cumulative(1E6,1E6,1)", gam_cdf(1e6, 1e6, 1),
     "large shape"),
    # ---- Beta: no .RT --------------------------------------------------------
    ("beta cdf, body", "BETA.DIST(0.3,2,3,TRUE)",
     "K_STATS_Beta_Cumulative(0.3,2,3)", beta_cdf(0.3, 2, 3),
     "body"),
    ("beta survival, near 1", "1-BETA.DIST(0.999,2,3,TRUE)",
     "K_STATS_Beta_Survival(0.999,2,3)", 1 - beta_cdf(mp.mpf("0.999"), 2, 3),
     "no direct survival; true value is about 4E-9"),
    # ---- Exponential: no .RT -------------------------------------------------
    ("exponential survival, deep", "1-EXPON.DIST(50,1,TRUE)",
     "K_STATS_Exponential_Survival(50,1)", mp.e ** mp.mpf(-50),
     "no direct survival; true value e^-50, about 2E-22"),
    # ---- Weibull: no .RT -----------------------------------------------------
    ("weibull survival, deep", "1-WEIBULL.DIST(10,2,1,TRUE)",
     "K_STATS_Weibull_Survival(10,2,1)", mp.e ** mp.mpf(-100),
     "no direct survival; true value e^-100, about 4E-44"),
    # ---- Uniform: Excel has nothing -----------------------------------------
    ("uniform cdf", NONE, "K_STATS_Uniform_Cumulative(0.3,0,1)", mp.mpf("0.3"),
     "Excel has no uniform distribution function"),
    # ---- Binomial: partial ---------------------------------------------------
    ("binomial cdf, body", "BINOM.DIST(5,10,0.5,TRUE)",
     "K_STATS_Binomial_Cumulative(5,10,0.5)",
     sum(mp.binomial(10, i) * mp.mpf("0.5") ** 10 for i in range(6)),
     "body"),
    ("binomial survival, deep", "1-BINOM.DIST(90,100,0.5,TRUE)",
     "K_STATS_Binomial_Survival(90,100,0.5)", binom_sf(90, 100, mp.mpf("0.5")),
     "no direct survival; true value about 2E-17"),
    ("binomial log-pmf", NONE, "K_STATS_Binomial_LogPMF(900,1000,0.5)",
     mp.log(mp.binomial(1000, 900)) + 1000 * mp.log(mp.mpf("0.5")),
     "Excel has no log-mass; BINOM.DIST underflows to 0 here"),
    # ---- Poisson: no .RT -----------------------------------------------------
    ("poisson cdf, body", "POISSON.DIST(3,2,TRUE)",
     "K_STATS_Poisson_Cumulative(3,2)",
     sum(mp.e ** -2 * mp.mpf(2) ** i / mp.factorial(i) for i in range(4)),
     "body"),
    ("poisson survival, deep", "1-POISSON.DIST(60,2,TRUE)",
     "K_STATS_Poisson_Survival(60,2)", pois_sf(60, 2),
     "no direct survival; true value about 6E-67"),
    # ---- Negative binomial: no .RT ------------------------------------------
    ("negative binomial cdf", "NEGBINOM.DIST(5,3,0.5,TRUE)",
     "K_STATS_NegativeBinomial_Cumulative(5,3,0.5)",
     negbinom_cdf(5, 3, mp.mpf("0.5")), "body"),
    # ---- Hypergeometric: no .RT ---------------------------------------------
    ("hypergeometric pmf", "HYPGEOM.DIST(2,5,10,50,FALSE)",
     "K_STATS_Hypergeometric_PMF(2,5,10,50)", hyper_pmf(2, 5, 10, 50),
     "body"),
    # ---- Discrete uniform: Excel has nothing ---------------------------------
    ("discrete uniform cdf", NONE,
     "K_STATS_DiscreteUniform_Cumulative(3,1,10)", mp.mpf(3) / 10,
     "Excel has no discrete uniform distribution function"),
]


if __name__ == "__main__":
    # PIPE-delimited, not comma. Every interesting field here contains commas -
    # "normal survival, z = 8", "CHISQ.DIST(1E6,1E6,TRUE)" - and the VBA exporter
    # splits on the delimiter without quote handling, so a comma delimiter wrote
    # observations over the reference column. A delimiter that cannot appear in
    # the data is simpler and safer than teaching VBA to parse quoted CSV.
    fields = ["label", "excel_formula", "kstats_call", "reference",
              "observed_excel", "observed_kstats", "why"]
    with open("excel_comparison_grid.csv", "w", newline="\n", encoding="utf-8") as f:
        f.write("|".join(fields) + "\n")
        for label, xl, ks, ref, why in CASES:
            for field in (label, xl, ks, why):
                assert "|" not in field, f"pipe in field: {field}"
            f.write("|".join([label, xl, ks, mp.nstr(ref, 34), "", "", why]) + "\n")
    print(f"wrote excel_comparison_grid.csv: {len(CASES)} cases")
    for label, xl, ks, ref, why in CASES:
        print(f"  {label:<28} true = {mp.nstr(ref, 12)}")
