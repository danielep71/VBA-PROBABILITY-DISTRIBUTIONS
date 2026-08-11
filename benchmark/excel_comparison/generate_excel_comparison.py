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


CASES = [
    # (label, excel_formula, k_stats_call, reference, why this point)
    ("normal survival, z = 8", "1-NORM.S.DIST(8,TRUE)",
     "K_STATS_NormalStandard_Survival(8)", norm_sf(8),
     "Excel has no direct standard-normal survival; 1-CDF loses the tail"),
    ("normal survival, z = 10", "1-NORM.S.DIST(10,TRUE)",
     "K_STATS_NormalStandard_Survival(10)", norm_sf(10),
     "same, one decade deeper"),
    ("normal survival, z = 15", "1-NORM.S.DIST(15,TRUE)",
     "K_STATS_NormalStandard_Survival(15)", norm_sf(15),
     "1-CDF is exactly 0 in Double here"),
    ("normal cdf, z = -8", "NORM.S.DIST(-8,TRUE)",
     "K_STATS_NormalStandard_Cumulative(-8)", norm_cdf(-8),
     "direct lower tail: both should be accurate"),
    ("chi-square cdf, df 1E6", "CHISQ.DIST(1E6,1E6,TRUE)",
     "K_STATS_ChiSquare_Cumulative(1E6,1E6)", chi_cdf(1e6, 1e6),
     "large df, near the median"),
    ("chi-square sf, df 1E6", "CHISQ.DIST.RT(1010000,1E6)",
     "K_STATS_ChiSquare_Survival(1010000,1E6)", chi_sf(1.01e6, 1e6),
     "large df, upper tail"),
    ("chi-square cdf, df 1E8", "CHISQ.DIST(1E8,1E8,TRUE)",
     "K_STATS_ChiSquare_Cumulative(1E8,1E8)", chi_cdf(1e8, 1e8),
     "at this library's chi envelope"),
    ("t survival, t=30 df=10", "T.DIST.RT(30,10)",
     "K_STATS_StudentT_Survival(30,10)", t_sf(30, 10),
     "moderate df, deep tail"),
    ("t survival, t=50 df=5", "T.DIST.RT(50,5)",
     "K_STATS_StudentT_Survival(50,5)", t_sf(50, 5),
     "small df, deep tail"),
    ("t survival, t=2 df=1E7", "T.DIST.RT(2,1E7)",
     "K_STATS_StudentT_Survival(2,1E7)", t_sf(2, 1e7),
     "df beyond Excel's documented range"),
    ("F survival, f=1 df 1E6", "F.DIST.RT(1,1E6,1E6)",
     "K_STATS_F_Survival(1,1E6,1E6)", f_sf(1, 1e6, 1e6),
     "large balanced df"),
    ("F survival, f=1 df 1E6,3", "F.DIST.RT(1,1E6,3)",
     "K_STATS_F_Survival(1,1E6,3)", f_sf(1, 1e6, 3),
     "extreme df ratio"),
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
