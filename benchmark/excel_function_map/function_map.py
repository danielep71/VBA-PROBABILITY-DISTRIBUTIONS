"""
#330: the map from Excel's distribution functions to the library's.

One source for three things: the table of docs/PUBLIC_API.md
(render_map.py), the grid Excel evaluates (build_grid.py) and the tests.

A ROW is one Excel function in one of its modes:

  excel     the Excel call, with argument names as Microsoft documents them;
  legacy    older Excel names with the same meaning, kept for compatibility;
  kstats    the library expression that replaces it, in the same names;
  notes     what differs: argument order, parameterization, tail;
  example   (Excel formula, library expression) with numbers. Excel
            evaluates both; the library side is also computed here from
            120-digit arithmetic.

A PROBE is a pair of formulas that settles a question the documentation
leaves open, by measurement: what each side does with a fractional degree
of freedom or count. `truncated` is the library expression at the argument
rounded down, to recognise an Excel that truncates.

An EDGE is an input where the replacement of a row stops applying. Both
formulas are evaluated as they stand, to show what a reader who copied the
row would get.
"""

ROWS = [
    # --- Normal -----------------------------------------------------------
    dict(family="Normal", excel="NORM.S.DIST(z, TRUE)", legacy="NORMSDIST(z)",
         kstats="K_STATS_NormalStandard_Cumulative(z)", notes="Same argument.",
         example=("NORM.S.DIST(1.5,TRUE)", "K_STATS_NormalStandard_Cumulative(1.5)")),
    dict(family="Normal", excel="NORM.S.DIST(z, FALSE)", legacy="PHI(z)",
         kstats="K_STATS_NormalStandard_Density(z)", notes="Same argument.",
         example=("NORM.S.DIST(1.5,FALSE)", "K_STATS_NormalStandard_Density(1.5)")),
    dict(family="Normal", excel="1 - NORM.S.DIST(z, TRUE)", legacy="",
         kstats="K_STATS_NormalStandard_Survival(z)",
         notes="Excel has no upper-tail normal function. The subtraction loses the tail "
               "beyond about z = 8.",
         example=("1-NORM.S.DIST(9,TRUE)", "K_STATS_NormalStandard_Survival(9)")),
    dict(family="Normal", excel="NORM.S.INV(p)", legacy="NORMSINV(p)",
         kstats="K_STATS_NormalStandard_InverseCumulative(p)", notes="Same argument.",
         example=("NORM.S.INV(0.975)", "K_STATS_NormalStandard_InverseCumulative(0.975)")),
    dict(family="Normal", excel="NORM.DIST(x, mean, sd, TRUE)", legacy="NORMDIST(x, mean, sd, TRUE)",
         kstats="K_STATS_Normal_Cumulative(x, mean, sd)", notes="Same order; no cumulative flag.",
         example=("NORM.DIST(12,10,2,TRUE)", "K_STATS_Normal_Cumulative(12,10,2)")),
    dict(family="Normal", excel="NORM.DIST(x, mean, sd, FALSE)", legacy="NORMDIST(x, mean, sd, FALSE)",
         kstats="K_STATS_Normal_Density(x, mean, sd)", notes="Same order; no cumulative flag.",
         example=("NORM.DIST(12,10,2,FALSE)", "K_STATS_Normal_Density(12,10,2)")),
    dict(family="Normal", excel="NORM.INV(p, mean, sd)", legacy="NORMINV(p, mean, sd)",
         kstats="K_STATS_Normal_InverseCumulative(p, mean, sd)", notes="Same order.",
         example=("NORM.INV(0.975,10,2)", "K_STATS_Normal_InverseCumulative(0.975,10,2)")),
    dict(family="Normal", excel="STANDARDIZE(x, mean, sd)", legacy="",
         kstats="K_STATS_Normal_ZScore(x, mean, sd)", notes="Same order.",
         example=("STANDARDIZE(12,10,2)", "K_STATS_Normal_ZScore(12,10,2)")),
    dict(family="Normal", excel="GAUSS(z)", legacy="",
         kstats="K_STATS_NormalStandard_IntervalProbability(0, z)",
         notes="The probability between the mean and z, for z of zero or more. The library "
               "takes its two bounds in order and refuses an upper bound below the lower.",
         example=("GAUSS(1.5)", "K_STATS_NormalStandard_IntervalProbability(0,1.5)")),
    dict(family="Normal", excel="GAUSS(z), z below zero", legacy="",
         kstats="-K_STATS_NormalStandard_IntervalProbability(z, 0)",
         notes="Excel returns a negative number for a negative z; put the bounds in order and "
               "change the sign.",
         example=("GAUSS(-1.5)", "-K_STATS_NormalStandard_IntervalProbability(-1.5,0)")),
    # --- Lognormal --------------------------------------------------------
    dict(family="Lognormal", excel="LOGNORM.DIST(x, mean, sd, TRUE)", legacy="LOGNORMDIST(x, mean, sd)",
         kstats="K_STATS_Lognormal_Cumulative(x, mean, sd)",
         notes="Same order. Both take the mean and standard deviation of ln(x).",
         example=("LOGNORM.DIST(2,0.5,0.8,TRUE)", "K_STATS_Lognormal_Cumulative(2,0.5,0.8)")),
    dict(family="Lognormal", excel="LOGNORM.DIST(x, mean, sd, FALSE)", legacy="",
         kstats="K_STATS_Lognormal_Density(x, mean, sd)", notes="Same order.",
         example=("LOGNORM.DIST(2,0.5,0.8,FALSE)", "K_STATS_Lognormal_Density(2,0.5,0.8)")),
    dict(family="Lognormal", excel="LOGNORM.INV(p, mean, sd)", legacy="LOGINV(p, mean, sd)",
         kstats="K_STATS_Lognormal_InverseCumulative(p, mean, sd)", notes="Same order.",
         example=("LOGNORM.INV(0.9,0.5,0.8)", "K_STATS_Lognormal_InverseCumulative(0.9,0.5,0.8)")),
    # --- Student t --------------------------------------------------------
    dict(family="Student t", excel="T.DIST(x, df, TRUE)", legacy="",
         kstats="K_STATS_StudentT_Cumulative(x, df)", notes="Same order; lower tail.",
         example=("T.DIST(1.5,10,TRUE)", "K_STATS_StudentT_Cumulative(1.5,10)")),
    dict(family="Student t", excel="T.DIST(x, df, FALSE)", legacy="",
         kstats="K_STATS_StudentT_Density(x, df)", notes="Same order.",
         example=("T.DIST(1.5,10,FALSE)", "K_STATS_StudentT_Density(1.5,10)")),
    dict(family="Student t", excel="T.DIST.RT(x, df)", legacy="TDIST(x, df, 1)",
         kstats="K_STATS_StudentT_Survival(x, df)",
         notes="Same order; upper tail. `TDIST` refuses a negative x.",
         example=("T.DIST.RT(1.5,10)", "K_STATS_StudentT_Survival(1.5,10)")),
    dict(family="Student t", excel="T.DIST.2T(x, df)", legacy="TDIST(x, df, 2)",
         kstats="2 * K_STATS_StudentT_Survival(x, df)",
         notes="Two tails, for x of zero or more. The library has no two-tailed function.",
         example=("T.DIST.2T(1.5,10)", "2*K_STATS_StudentT_Survival(1.5,10)")),
    dict(family="Student t", excel="T.INV(p, df)", legacy="",
         kstats="K_STATS_StudentT_InverseCumulative(p, df)", notes="Same order; lower tail.",
         example=("T.INV(0.975,10)", "K_STATS_StudentT_InverseCumulative(0.975,10)")),
    dict(family="Student t", excel="T.INV.2T(p, df)", legacy="TINV(p, df)",
         kstats="-K_STATS_StudentT_InverseCumulative(p / 2, df)",
         notes="p is the probability in both tails together.",
         example=("T.INV.2T(0.05,10)", "-K_STATS_StudentT_InverseCumulative(0.025,10)")),
    # --- Chi-square -------------------------------------------------------
    dict(family="Chi-square", excel="CHISQ.DIST(x, df, TRUE)", legacy="",
         kstats="K_STATS_ChiSquare_Cumulative(x, df)", notes="Same order; lower tail.",
         example=("CHISQ.DIST(3,4,TRUE)", "K_STATS_ChiSquare_Cumulative(3,4)")),
    dict(family="Chi-square", excel="CHISQ.DIST(x, df, FALSE)", legacy="",
         kstats="K_STATS_ChiSquare_Density(x, df)", notes="Same order.",
         example=("CHISQ.DIST(3,4,FALSE)", "K_STATS_ChiSquare_Density(3,4)")),
    dict(family="Chi-square", excel="CHISQ.DIST.RT(x, df)", legacy="CHIDIST(x, df)",
         kstats="K_STATS_ChiSquare_Survival(x, df)", notes="Same order; upper tail.",
         example=("CHISQ.DIST.RT(3,4)", "K_STATS_ChiSquare_Survival(3,4)")),
    dict(family="Chi-square", excel="CHISQ.INV(p, df)", legacy="",
         kstats="K_STATS_ChiSquare_InverseCumulative(p, df)", notes="Same order; lower tail.",
         example=("CHISQ.INV(0.95,4)", "K_STATS_ChiSquare_InverseCumulative(0.95,4)")),
    dict(family="Chi-square", excel="CHISQ.INV.RT(p, df)", legacy="CHIINV(p, df)",
         kstats="K_STATS_ChiSquare_InverseCumulative(1 - p, df)",
         notes="The library has no upper-tail chi-square inverse. `1 - p` is exact for p of "
               "0.5 or more and loses digits of a p near zero.",
         example=("CHISQ.INV.RT(0.05,4)", "K_STATS_ChiSquare_InverseCumulative(1-0.05,4)")),
    # --- F ----------------------------------------------------------------
    dict(family="F", excel="F.DIST(x, df1, df2, TRUE)", legacy="",
         kstats="K_STATS_F_Cumulative(x, df1, df2)", notes="Same order; lower tail.",
         example=("F.DIST(2.5,5,10,TRUE)", "K_STATS_F_Cumulative(2.5,5,10)")),
    dict(family="F", excel="F.DIST(x, df1, df2, FALSE)", legacy="",
         kstats="K_STATS_F_Density(x, df1, df2)", notes="Same order.",
         example=("F.DIST(2.5,5,10,FALSE)", "K_STATS_F_Density(2.5,5,10)")),
    dict(family="F", excel="F.DIST.RT(x, df1, df2)", legacy="FDIST(x, df1, df2)",
         kstats="K_STATS_F_Survival(x, df1, df2)", notes="Same order; upper tail.",
         example=("F.DIST.RT(2.5,5,10)", "K_STATS_F_Survival(2.5,5,10)")),
    dict(family="F", excel="F.INV(p, df1, df2)", legacy="",
         kstats="K_STATS_F_InverseCumulative(p, df1, df2)", notes="Same order; lower tail.",
         example=("F.INV(0.95,5,10)", "K_STATS_F_InverseCumulative(0.95,5,10)")),
    dict(family="F", excel="F.INV.RT(p, df1, df2)", legacy="FINV(p, df1, df2)",
         kstats="K_STATS_F_InverseCumulative(1 - p, df1, df2)",
         notes="The library has no upper-tail F inverse. `1 - p` is exact for p of 0.5 or "
               "more and loses digits of a p near zero.",
         example=("F.INV.RT(0.05,5,10)", "K_STATS_F_InverseCumulative(1-0.05,5,10)")),
    # --- Gamma ------------------------------------------------------------
    dict(family="Gamma", excel="GAMMA.DIST(x, alpha, beta, TRUE)", legacy="GAMMADIST(x, alpha, beta, TRUE)",
         kstats="K_STATS_Gamma_Cumulative(x, alpha, beta)",
         notes="Same order. In both, alpha is the shape and beta the scale, not the rate.",
         example=("GAMMA.DIST(3,2,1.5,TRUE)", "K_STATS_Gamma_Cumulative(3,2,1.5)")),
    dict(family="Gamma", excel="GAMMA.DIST(x, alpha, beta, FALSE)", legacy="GAMMADIST(x, alpha, beta, FALSE)",
         kstats="K_STATS_Gamma_Density(x, alpha, beta)", notes="Same order; beta is the scale.",
         example=("GAMMA.DIST(3,2,1.5,FALSE)", "K_STATS_Gamma_Density(3,2,1.5)")),
    dict(family="Gamma", excel="GAMMA.INV(p, alpha, beta)", legacy="GAMMAINV(p, alpha, beta)",
         kstats="K_STATS_Gamma_InverseCumulative(p, alpha, beta)", notes="Same order; beta is the scale.",
         example=("GAMMA.INV(0.9,2,1.5)", "K_STATS_Gamma_InverseCumulative(0.9,2,1.5)")),
    # --- Beta -------------------------------------------------------------
    dict(family="Beta", excel="BETA.DIST(x, alpha, beta, TRUE)", legacy="BETADIST(x, alpha, beta)",
         kstats="K_STATS_Beta_Cumulative(x, alpha, beta)", notes="Same order, on the interval 0 to 1.",
         example=("BETA.DIST(0.3,2,3,TRUE)", "K_STATS_Beta_Cumulative(0.3,2,3)")),
    dict(family="Beta", excel="BETA.DIST(x, alpha, beta, FALSE)", legacy="",
         kstats="K_STATS_Beta_Density(x, alpha, beta)", notes="Same order, on the interval 0 to 1.",
         example=("BETA.DIST(0.3,2,3,FALSE)", "K_STATS_Beta_Density(0.3,2,3)")),
    dict(family="Beta", excel="BETA.DIST(x, alpha, beta, TRUE, A, B)", legacy="BETADIST(x, alpha, beta, A, B)",
         kstats="K_STATS_Beta_Cumulative((x - A) / (B - A), alpha, beta)",
         notes="The library works on 0 to 1; move x there first. A density is divided by B - A as well.",
         example=("BETA.DIST(14,2,3,TRUE,10,20)", "K_STATS_Beta_Cumulative((14-10)/(20-10),2,3)")),
    dict(family="Beta", excel="BETA.INV(p, alpha, beta)", legacy="BETAINV(p, alpha, beta)",
         kstats="K_STATS_Beta_InverseCumulative(p, alpha, beta)", notes="Same order, on the interval 0 to 1.",
         example=("BETA.INV(0.4,2,3)", "K_STATS_Beta_InverseCumulative(0.4,2,3)")),
    dict(family="Beta", excel="BETA.INV(p, alpha, beta, A, B)", legacy="BETAINV(p, alpha, beta, A, B)",
         kstats="A + (B - A) * K_STATS_Beta_InverseCumulative(p, alpha, beta)",
         notes="Move the quantile from 0 to 1 back to A to B.",
         example=("BETA.INV(0.4,2,3,10,20)", "10+(20-10)*K_STATS_Beta_InverseCumulative(0.4,2,3)")),
    # --- Exponential and Weibull -----------------------------------------
    dict(family="Exponential", excel="EXPON.DIST(x, lambda, TRUE)", legacy="EXPONDIST(x, lambda, TRUE)",
         kstats="K_STATS_Exponential_Cumulative(x, lambda)", notes="Same order. In both, lambda is the rate.",
         example=("EXPON.DIST(2,0.7,TRUE)", "K_STATS_Exponential_Cumulative(2,0.7)")),
    dict(family="Exponential", excel="EXPON.DIST(x, lambda, FALSE)", legacy="EXPONDIST(x, lambda, FALSE)",
         kstats="K_STATS_Exponential_Density(x, lambda)", notes="Same order; lambda is the rate.",
         example=("EXPON.DIST(2,0.7,FALSE)", "K_STATS_Exponential_Density(2,0.7)")),
    dict(family="Weibull", excel="WEIBULL.DIST(x, alpha, beta, TRUE)", legacy="WEIBULL(x, alpha, beta, TRUE)",
         kstats="K_STATS_Weibull_Cumulative(x, alpha, beta)",
         notes="Same order. In both, alpha is the shape and beta the scale.",
         example=("WEIBULL.DIST(1.2,2,1.5,TRUE)", "K_STATS_Weibull_Cumulative(1.2,2,1.5)")),
    dict(family="Weibull", excel="WEIBULL.DIST(x, alpha, beta, FALSE)", legacy="WEIBULL(x, alpha, beta, FALSE)",
         kstats="K_STATS_Weibull_Density(x, alpha, beta)", notes="Same order.",
         example=("WEIBULL.DIST(1.2,2,1.5,FALSE)", "K_STATS_Weibull_Density(1.2,2,1.5)")),
    # --- Binomial ---------------------------------------------------------
    dict(family="Binomial", excel="BINOM.DIST(k, n, p, TRUE)", legacy="BINOMDIST(k, n, p, TRUE)",
         kstats="K_STATS_Binomial_Cumulative(k, n, p)", notes="Same order.",
         example=("BINOM.DIST(6,20,0.3,TRUE)", "K_STATS_Binomial_Cumulative(6,20,0.3)")),
    dict(family="Binomial", excel="BINOM.DIST(k, n, p, FALSE)", legacy="BINOMDIST(k, n, p, FALSE)",
         kstats="K_STATS_Binomial_PMF(k, n, p)", notes="Same order.",
         example=("BINOM.DIST(6,20,0.3,FALSE)", "K_STATS_Binomial_PMF(6,20,0.3)")),
    dict(family="Binomial", excel="BINOM.DIST.RANGE(n, p, k1, k2)", legacy="",
         kstats="K_STATS_Binomial_Cumulative(k2, n, p) - K_STATS_Binomial_Cumulative(k1 - 1, n, p)",
         notes="**Different order**: Excel puts n and p first. For k1 of one or more. With k2 "
               "omitted it is `K_STATS_Binomial_PMF(k1, n, p)`.",
         example=("BINOM.DIST.RANGE(20,0.3,4,8)",
                  "K_STATS_Binomial_Cumulative(8,20,0.3)-K_STATS_Binomial_Cumulative(3,20,0.3)")),
    dict(family="Binomial", excel="BINOM.DIST.RANGE(n, p, 0, k2)", legacy="",
         kstats="K_STATS_Binomial_Cumulative(k2, n, p)",
         notes="With a lower bound of zero the range is the cumulative probability; the "
               "difference above would ask the library for a count of -1.",
         example=("BINOM.DIST.RANGE(20,0.3,0,8)", "K_STATS_Binomial_Cumulative(8,20,0.3)")),
    dict(family="Binomial", excel="BINOM.INV(n, p, alpha)", legacy="CRITBINOM(n, p, alpha)",
         kstats="K_STATS_Binomial_InverseCumulative(alpha, n, p)",
         notes="**Different order**: the library puts the probability first, as in every inverse.",
         example=("BINOM.INV(20,0.3,0.9)", "K_STATS_Binomial_InverseCumulative(0.9,20,0.3)")),
    # --- Negative binomial, Poisson, hypergeometric -----------------------
    dict(family="Negative binomial", excel="NEGBINOM.DIST(f, s, p, TRUE)", legacy="",
         kstats="K_STATS_NegativeBinomial_Cumulative(f, s, p)",
         notes="Same order: f failures before the s-th success.",
         example=("NEGBINOM.DIST(5,3,0.4,TRUE)", "K_STATS_NegativeBinomial_Cumulative(5,3,0.4)")),
    dict(family="Negative binomial", excel="NEGBINOM.DIST(f, s, p, FALSE)", legacy="NEGBINOMDIST(f, s, p)",
         kstats="K_STATS_NegativeBinomial_PMF(f, s, p)", notes="Same order.",
         example=("NEGBINOM.DIST(5,3,0.4,FALSE)", "K_STATS_NegativeBinomial_PMF(5,3,0.4)")),
    dict(family="Poisson", excel="POISSON.DIST(k, mean, TRUE)", legacy="POISSON(k, mean, TRUE)",
         kstats="K_STATS_Poisson_Cumulative(k, mean)", notes="Same order.",
         example=("POISSON.DIST(4,2.5,TRUE)", "K_STATS_Poisson_Cumulative(4,2.5)")),
    dict(family="Poisson", excel="POISSON.DIST(k, mean, FALSE)", legacy="POISSON(k, mean, FALSE)",
         kstats="K_STATS_Poisson_PMF(k, mean)", notes="Same order.",
         example=("POISSON.DIST(4,2.5,FALSE)", "K_STATS_Poisson_PMF(4,2.5)")),
    dict(family="Hypergeometric",
         excel="HYPGEOM.DIST(sample_s, number_sample, population_s, number_pop, TRUE)", legacy="",
         kstats="K_STATS_Hypergeometric_Cumulative(sample_s, number_sample, population_s, number_pop)",
         notes="Same order.",
         example=("HYPGEOM.DIST(2,10,8,40,TRUE)", "K_STATS_Hypergeometric_Cumulative(2,10,8,40)")),
    dict(family="Hypergeometric",
         excel="HYPGEOM.DIST(sample_s, number_sample, population_s, number_pop, FALSE)",
         legacy="HYPGEOMDIST(sample_s, number_sample, population_s, number_pop)",
         kstats="K_STATS_Hypergeometric_PMF(sample_s, number_sample, population_s, number_pop)",
         notes="Same order.",
         example=("HYPGEOM.DIST(2,10,8,40,FALSE)", "K_STATS_Hypergeometric_PMF(2,10,8,40)")),
]

# What each side does with a fractional degree of freedom or count.
PROBES = [
    dict(topic="Student t, df = 10.6", excel="T.DIST(1.5,10.6,TRUE)",
         kstats="K_STATS_StudentT_Cumulative(1.5,10.6)", truncated="K_STATS_StudentT_Cumulative(1.5,10)"),
    dict(topic="Student t inverse, df = 10.6", excel="T.INV(0.975,10.6)",
         kstats="K_STATS_StudentT_InverseCumulative(0.975,10.6)",
         truncated="K_STATS_StudentT_InverseCumulative(0.975,10)"),
    dict(topic="Chi-square, df = 4.6", excel="CHISQ.DIST(3,4.6,TRUE)",
         kstats="K_STATS_ChiSquare_Cumulative(3,4.6)", truncated="K_STATS_ChiSquare_Cumulative(3,4)"),
    dict(topic="Chi-square inverse, df = 4.6", excel="CHISQ.INV(0.95,4.6)",
         kstats="K_STATS_ChiSquare_InverseCumulative(0.95,4.6)",
         truncated="K_STATS_ChiSquare_InverseCumulative(0.95,4)"),
    dict(topic="F, df1 = 5.6", excel="F.DIST(2.5,5.6,10,TRUE)",
         kstats="K_STATS_F_Cumulative(2.5,5.6,10)", truncated="K_STATS_F_Cumulative(2.5,5,10)"),
    dict(topic="Binomial, k = 6.6", excel="BINOM.DIST(6.6,20,0.3,TRUE)",
         kstats="K_STATS_Binomial_Cumulative(6.6,20,0.3)", truncated="K_STATS_Binomial_Cumulative(6,20,0.3)"),
    dict(topic="Poisson, k = 4.6", excel="POISSON.DIST(4.6,2.5,TRUE)",
         kstats="K_STATS_Poisson_Cumulative(4.6,2.5)", truncated="K_STATS_Poisson_Cumulative(4,2.5)"),
    dict(topic="Negative binomial, f = 5.6", excel="NEGBINOM.DIST(5.6,3,0.4,TRUE)",
         kstats="K_STATS_NegativeBinomial_Cumulative(5.6,3,0.4)",
         truncated="K_STATS_NegativeBinomial_Cumulative(5,3,0.4)"),
]

# Where the replacement of a row above stops applying: the Excel formula and
# the library expression that row would give, both evaluated as they stand.
EDGES = [
    dict(topic="GAUSS at a negative z, bounds not reordered", excel="GAUSS(-1.5)",
         kstats="K_STATS_NormalStandard_IntervalProbability(0,-1.5)"),
    dict(topic="Binomial range from zero, as a difference", excel="BINOM.DIST.RANGE(20,0.3,0,8)",
         kstats="K_STATS_Binomial_Cumulative(8,20,0.3)-K_STATS_Binomial_Cumulative(-1,20,0.3)"),
]

# Excel functions with no counterpart in the library, and why.
WITHOUT_COUNTERPART = [
    ("CHISQ.TEST, F.TEST, T.TEST, Z.TEST", "Hypothesis tests on data ranges. The library holds "
     "distributions; the tail probability of a statistic you have computed is a survival function."),
    ("CONFIDENCE.NORM, CONFIDENCE.T", "Interval half-widths: a quantile times a standard error."),
    ("GAMMA, GAMMALN, GAMMALN.PRECISE, ERF, ERFC, ERF.PRECISE, ERFC.PRECISE",
     "Special functions. The library's own are internal and not part of its public surface."),
    ("FISHER, FISHERINV, PROB, PERCENTILE.*, QUARTILE.*", "Transformations and statistics of data, "
     "not distribution functions."),
]

# Library families and functions Excel has no function for.
LIBRARY_ONLY = [
    ("Survival functions", "`K_STATS_<Family>_Survival`, for every family: the upper tail computed "
     "directly, where Excel has an `.RT` function only for Student t, chi-square and F."),
    ("Inverse survival", "`K_STATS_NormalStandard_InverseSurvival`, `K_STATS_Normal_InverseSurvival`, "
     "`K_STATS_Lognormal_InverseSurvival`."),
    ("Logarithm of a probability mass", "`K_STATS_<Family>_LogPMF`, for the discrete families."),
    ("Inverses Excel lacks", "`K_STATS_Weibull_InverseCumulative`, "
     "`K_STATS_Exponential_InverseCumulative`, `K_STATS_Poisson_InverseCumulative`, "
     "`K_STATS_NegativeBinomial_InverseCumulative`, `K_STATS_Hypergeometric_InverseCumulative`."),
    ("Families", "Uniform, discrete uniform and geometric. A geometric probability is "
     "`NEGBINOM.DIST(f, 1, p, ...)` in Excel."),
    ("Moments", "`K_STATS_<Family>_Mean`, `_Variance`, `_StdDev`."),
]
