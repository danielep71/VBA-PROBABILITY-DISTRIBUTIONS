# 6. Distribution families

This chapter lists every distribution shipped in v1.0.0. "Name" is the name used inside function names. Parameter names are the argument names of the functions. In the kind tables, "Inv." means `InverseCumulative`, "Inv. surv." means `InverseSurvival`, and "Moments" means the three functions `Mean`, `Variance` and `StdDev`.

The mathematical support is not the same as the numerical domain in which the library was tested; see section 6.5 and chapter 10.

## 6.1 Normal family

| Distribution | Name | Parameters | Support |
|---|---|---|---|
| Standard Normal | `NormalStandard` | none | all real numbers |
| Normal | `Normal` | `Mean`, `StdDev` | all real numbers |
| Lognormal | `Lognormal` | `MeanLog`, `StdDevLog` | $x > 0$ |

`StdDev` and `StdDevLog` must be strictly positive. For the Normal functions, `Mean` and `StdDev` are optional (0 and 1). `MeanLog` and `StdDevLog` describe $\ln X$, not $X$ itself.

| Distribution | Density | Cumulative | Survival | Inv. | Inv. surv. | Moments |
|---|:---:|:---:|:---:|:---:|:---:|:---:|
| Standard Normal | yes | yes | yes | yes | yes | no |
| Normal | yes | yes | yes | yes | yes | no |
| Lognormal | yes | yes | yes | yes | yes | yes |

Additional functions:

| Function | Purpose |
|---|---|
| `K_STATS_NormalStandard_IntervalProbability` | Probability between two z-values |
| `K_STATS_NormalStandard_InverseCumulativeFast` | Fast quantile for VBA loops |
| `K_STATS_Normal_IntervalProbability` | Probability between two bounds |
| `K_STATS_Normal_ZScore` | Standardized value |
| `K_STATS_Lognormal_ParametersFromMeanStdDev` | Mean and std. dev. to parameters |

The interval functions require the upper bound to be at least the lower bound.

## 6.2 Test-statistic family

| Distribution | Name | Parameters | Support |
|---|---|---|---|
| Student t | `StudentT` | `DegreesFreedom` | all real numbers |
| Chi-square | `ChiSquare` | `DegreesFreedom` | $x \ge 0$ |
| F | `F` | `DegreesFreedom1`, `DegreesFreedom2` | $x \ge 0$ |

Degrees of freedom must be strictly positive and need not be integers. `DegreesFreedom1` is the numerator and `DegreesFreedom2` the denominator.

| Distribution | Density | Cumulative | Survival | Inv. | Inv. surv. | Moments |
|---|:---:|:---:|:---:|:---:|:---:|:---:|
| Student t | yes | yes | yes | yes | no | no |
| Chi-square | yes | yes | yes | yes | no | no |
| F | yes | yes | yes | yes | no | no |

Degrees of freedom above the following limits return `#NUM!`.

| Functions | Limit |
|---|---|
| Student t cumulative, survival, inverse | 1E8 |
| Chi-square cumulative, survival, inverse | 1E8 |
| F cumulative, survival, inverse | 1E10 each |
| Chi-square and F density | 1E20 |

The Student t density has no such limit.

## 6.3 Other continuous distributions

| Distribution | Name | Parameters | Support |
|---|---|---|---|
| Gamma | `Gamma` | `Shape`, `ScaleParam` | $x \ge 0$ |
| Beta | `Beta` | `Alpha`, `Beta` | $0 \le x \le 1$ |
| Exponential | `Exponential` | `Lambda` (rate) | $x \ge 0$ |
| Weibull | `Weibull` | `Shape`, `ScaleParam` | $x \ge 0$ |
| Uniform | `Uniform` | `LowerBound`, `UpperBound` | between the bounds |

All shape, scale and rate parameters must be strictly positive. `LowerBound` must be less than `UpperBound`.

| Distribution | Density | Cumulative | Survival | Inv. | Inv. surv. | Moments |
|---|:---:|:---:|:---:|:---:|:---:|:---:|
| Gamma | yes | yes | yes | yes | no | yes |
| Beta | yes | yes | yes | yes | no | yes |
| Exponential | yes | yes | yes | yes | no | no |
| Weibull | yes | yes | yes | yes | no | yes |
| Uniform | yes | yes | yes | yes | no | no |

Some densities are unbounded at an end point of the support. The library returns `#NUM!` there. The same holds for the Chi-square and F densities.

| Density | Point | `#NUM!` when |
|---|---|---|
| Gamma, Weibull | $x = 0$ | `Shape` < 1 |
| Chi-square | $x = 0$ | `DegreesFreedom` < 2 |
| F | $x = 0$ | `DegreesFreedom1` < 2 |
| Beta | $x = 0$ | `Alpha` < 1 |
| Beta | $x = 1$ | `Beta` < 1 |

The Gamma and Beta densities reject a shape above 1E20.

## 6.4 Discrete distributions

| Distribution | Name | First argument | Parameters |
|---|---|---|---|
| Binomial | `Binomial` | `NumberSuccesses` | `Trials`, `ProbSuccess` |
| Poisson | `Poisson` | `NumberEvents` | `Mean` |
| Geometric | `Geometric` | `NumberFailures` | `ProbSuccess` |
| Negative Binomial | `NegativeBinomial` | `NumberFailures` | `NumberSuccesses`, `ProbSuccess` |
| Hypergeometric | `Hypergeometric` | `SampleSuccesses` | see below |
| Discrete Uniform | `DiscreteUniform` | `X` | `LowerBound`, `UpperBound` |

The Hypergeometric parameters are `SampleSize`, `PopulationSuccesses`, `PopulationSize`.

| Distribution | Counts | Parameter rules |
|---|---|---|
| Binomial | Successes in `Trials` trials | `ProbSuccess` from 0 to 1 |
| Poisson | Events | `Mean` not negative |
| Geometric | Failures before the first success | `ProbSuccess` above 0, at most 1 |
| Negative Binomial | Failures before the required success | `NumberSuccesses` at least 1 |
| Hypergeometric | Successes in the sample | `PopulationSize` at least 1 |
| Discrete Uniform | An integer between the bounds | Lower not above upper |

Further rules:

- The Binomial inverse requires `ProbSuccess` strictly between 0 and 1.
- The Negative Binomial `ProbSuccess` must be above 0 and at most 1.
- Hypergeometric `PopulationSuccesses` and `SampleSize` must not exceed `PopulationSize`.
- Discrete Uniform bounds may be negative and are inclusive.

Every discrete distribution has the same eight functions:

| Distribution | PMF | LogPMF | Cumulative | Survival | Inv. | Moments |
|---|:---:|:---:|:---:|:---:|:---:|:---:|
| All six | yes | yes | yes | yes | yes | yes |

Section 7.6 explains the discrete conventions.

## 6.5 Supported size of discrete arguments

These are limits of the implementation, not of the mathematics. An argument above the limit returns `#NUM!`. $2^{53} - 1$ = 9,007,199,254,740,991 is the largest whole number a Double holds without gaps.

| Functions | Supported domain |
|---|---|
| Binomial PMF, LogPMF, moments | `Trials` up to $2^{53} - 1$ |
| Binomial cumulative, survival, inverse | `Trials` up to 10,000,000 |
| Poisson PMF, LogPMF | `NumberEvents`, `Mean` up to $2^{53} - 1$ |
| Poisson cumulative, survival | `NumberEvents` up to 20,000,000; `Mean` up to 10,000,000 |
| Poisson inverse | `Mean` up to 10,000,000; quantile up to 20,000,000 |
| Geometric | counts and quantiles up to $2^{53} - 1$ |
| Negative Binomial PMF, LogPMF | failures + successes up to $2^{53} - 1$ |
| Negative Binomial cumulative, survival, inverse | failures or quantile up to 20,000,000 |
| | `NumberSuccesses` up to 10,000,000 |
| Hypergeometric | `PopulationSize` up to 100,000,000 |
| Discrete Uniform | bounds within $\pm(2^{53} - 1)$; at most $2^{53} - 1$ integers |

The Hypergeometric cumulative and survival functions sum at most 200,000 terms.
