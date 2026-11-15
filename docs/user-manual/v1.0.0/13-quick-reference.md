# 13. Quick reference

This chapter was produced from the frozen inventory [`docs/PUBLIC_API.txt`](../../PUBLIC_API.txt). For the full declarations, see the Public API reference.

## 13.1 The naming rule

Every function name is `K_STATS_` + family + `_` + kind. In the tables below, the column headings are the family names and the first column is the kind, both written exactly as they appear in the function name. "yes" means the function exists.

Example: column `Gamma`, row `Survival` gives `K_STATS_Gamma_Survival`.

Arguments are the first argument, then the distribution parameters of chapter 6, then an optional `Status`.

| Kind | First argument |
|---|---|
| `Density`, `Cumulative`, `Survival` | `X` (`Z` for `NormalStandard`) |
| `PMF`, `LogPMF`, discrete `Cumulative`, `Survival` | the count (section 6.4) |
| `InverseCumulative`, `InverseSurvival` | `Probability` |
| `Mean`, `Variance`, `StdDev` | none |

## 13.2 Normal family

| Kind | `NormalStandard` | `Normal` | `Lognormal` |
|---|:---:|:---:|:---:|
| `Density` | yes | yes | yes |
| `Cumulative` | yes | yes | yes |
| `Survival` | yes | yes | yes |
| `InverseCumulative` | yes | yes | yes |
| `InverseSurvival` | yes | yes | yes |
| `Mean` | no | no | yes |
| `Variance` | no | no | yes |
| `StdDev` | no | no | yes |

## 13.3 Test-statistic family

| Kind | `StudentT` | `ChiSquare` | `F` |
|---|:---:|:---:|:---:|
| `Density` | yes | yes | yes |
| `Cumulative` | yes | yes | yes |
| `Survival` | yes | yes | yes |
| `InverseCumulative` | yes | yes | yes |

## 13.4 Other continuous distributions

| Kind | `Gamma` | `Beta` | `Exponential` | `Weibull` | `Uniform` |
|---|:---:|:---:|:---:|:---:|:---:|
| `Density` | yes | yes | yes | yes | yes |
| `Cumulative` | yes | yes | yes | yes | yes |
| `Survival` | yes | yes | yes | yes | yes |
| `InverseCumulative` | yes | yes | yes | yes | yes |
| `Mean` | yes | yes | no | yes | no |
| `Variance` | yes | yes | no | yes | no |
| `StdDev` | yes | yes | no | yes | no |

## 13.5 Discrete distributions

All six families, `Binomial`, `Poisson`, `Geometric`, `NegativeBinomial`, `Hypergeometric` and `DiscreteUniform`, have the same eight kinds.

| Kind | All six families |
|---|:---:|
| `PMF` | yes |
| `LogPMF` | yes |
| `Cumulative` | yes |
| `Survival` | yes |
| `InverseCumulative` | yes |
| `Mean` | yes |
| `Variance` | yes |
| `StdDev` | yes |

## 13.6 Functions outside the tables

Five functions have a kind that is not a row above.

| Function | Arguments |
|---|---|
| `K_STATS_Lognormal_ParametersFromMeanStdDev` | `Mean`, `StdDev` |
| `K_STATS_Normal_IntervalProbability` | `LowerBound`, `UpperBound`, `Mean`, `StdDev` |
| `K_STATS_Normal_ZScore` | `X`, `Mean`, `StdDev` |
| `K_STATS_NormalStandard_IntervalProbability` | `LowerZ`, `UpperZ` |
| `K_STATS_NormalStandard_InverseCumulativeFast` | `Probability` |

`K_STATS_NormalStandard_InverseCumulativeFast` has no `Status` argument, returns a Double and never returns an error value (section 5.8). The parameter conversion returns `MeanLog` and `StdDevLog` (section 5.5).

The tables account for all 112 functions: 23 in the Normal family, 12 in the test-statistic family, 29 for the other continuous distributions and 48 for the discrete distributions.

## 13.7 Rules to remember

| Topic | Rule | Section |
|---|---|---|
| Probabilities | Between 0 and 1, never percentages | 7.1 |
| `Cumulative` | $P(X \le x)$ | 7.2 |
| `Survival` | $P(X > x)$, strict inequality | 7.2 |
| Small upper tail | `Survival`, not 1 minus `Cumulative` | 7.3 |
| Small exceedance input | `InverseSurvival` (Normal family) | 7.4 |
| Discrete quantile | Smallest count reaching the probability | 7.5 |
| Counts | Truncated toward zero | 7.6 |
| Geometric, Negative Binomial | Count failures, not trials | 7.6 |
| Gamma, Weibull | Shape and scale | 7.7 |
| Exponential | Rate `Lambda` | 7.7 |
| Lognormal | Parameters of $\ln X$ | 7.7 |
| Invalid input | `#NUM!` | 9.2 |
| Unexpected failure | `#VALUE!` | 9.3 |
