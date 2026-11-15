# 8. Worked examples

Each example gives the formulas, explains the arguments and shows the expected results. The results were computed independently of the library, from the mathematical definitions, with high-precision arithmetic. The script [`checks/example_values.py`](checks/example_values.py) reproduces them and prints the identifier shown beside each result.

The examples show how to call the functions. They do not say that a distribution is the right model for your data.

## 8.1 A threshold and a tail from a Normal model

A processing time is Normal with mean 100 and standard deviation 15. Which time is exceeded in 1 percent of cases? How likely is a time above 130, or between 85 and 115?

```text
=K_STATS_Normal_InverseCumulative(0.99,100,15)
=K_STATS_Normal_Survival(130,100,15)
=K_STATS_Normal_IntervalProbability(85,115,100,15)
```

`100` is `Mean` and `15` is `StdDev`. The leading arguments are the probability, the value, and the lower and upper bound.

| Id | Result | Meaning |
|---|---|---|
| 8.1-a | 134.8952181 | The 99th percentile |
| 8.1-b | 0.02275013195 | $P(X > 130)$ |
| 8.1-c | 0.6826894921 | $P(85 \le X \le 115)$ |

## 8.2 A very small exceedance probability

How likely is a standard normal value above 8? A daily return is Normal with mean 0.0005 and standard deviation 0.012: which return is exceeded with probability one in a million?

```text
=K_STATS_NormalStandard_Survival(8)
=K_STATS_Normal_InverseSurvival(0.000001,0.0005,0.012)
```

`8` is the z-value. `0.000001` is an upper-tail probability, not a cumulative one.

| Id | Result | Meaning |
|---|---|---|
| 8.2-a | 6.220960574E-16 | $P(Z > 8)$ |
| 8.2-b | 0.05754109171 | Return exceeded once in a million |

The first result cannot be obtained as one minus the cumulative probability (section 7.3).

## 8.3 A two-sided t test

A test statistic of 2.5 has 12 degrees of freedom. What is the two-sided p-value, and what is the critical value at the 5 percent level?

```text
=2*K_STATS_StudentT_Survival(2.5,12)
=K_STATS_StudentT_InverseCumulative(0.975,12)
```

`2.5` is the statistic and `12` is `DegreesFreedom`. `0.975` is the cumulative probability: a two-sided 5 percent test leaves 2.5 percent in each tail.

| Id | Result | Meaning |
|---|---|---|
| 8.3-a | 0.02791539957 | Two-sided p-value |
| 8.3-b | 2.17881283 | Two-sided 5 percent critical value |

For a statistic that may be negative, use `=2*K_STATS_StudentT_Survival(ABS(B2),C2)`.

## 8.4 Chi-square and F tests

A chi-square statistic of 18.3 has 10 degrees of freedom. An F statistic of 4.2 has 3 and 20 degrees of freedom. What are the p-values, and what is the 5 percent critical value of F?

```text
=K_STATS_ChiSquare_Survival(18.3,10)
=K_STATS_F_Survival(4.2,3,20)
=K_STATS_F_InverseCumulative(0.95,3,20)
```

`3` is `DegreesFreedom1` (numerator) and `20` is `DegreesFreedom2` (denominator). `0.95` is the cumulative probability for an upper tail of 5 percent.

| Id | Result | Meaning |
|---|---|---|
| 8.4-a | 0.05010906141 | p-value of the chi-square statistic |
| 8.4-b | 0.01855741983 | p-value of the F statistic |
| 8.4-c | 3.098391212 | 5 percent critical value of F(3, 20) |

## 8.5 A Lognormal loss from its mean and standard deviation

A loss has arithmetic mean 100 and standard deviation 30 and is modelled as Lognormal. How likely is a loss above 200? What is the 99.5th percentile?

The Lognormal functions need the parameters of $\ln X$. Convert first, in cells `B1` and `B2`:

```text
=INDEX(K_STATS_Lognormal_ParametersFromMeanStdDev(100,30),1,1)
=INDEX(K_STATS_Lognormal_ParametersFromMeanStdDev(100,30),1,2)
```

Then use the two cells as `MeanLog` and `StdDevLog`:

```text
=K_STATS_Lognormal_Survival(200,B1,B2)
=K_STATS_Lognormal_InverseCumulative(0.995,B1,B2)
=K_STATS_Lognormal_Mean(B1,B2)
```

| Id | Result | Meaning |
|---|---|---|
| 8.5-a | 4.562081338 | `MeanLog` |
| 8.5-b | 0.2935603792 | `StdDevLog` |
| 8.5-c | 0.006071617778 | $P(X > 200)$ |
| 8.5-d | 204.0250469 | The 99.5th percentile |
| 8.5-e | 100 | Arithmetic mean, as a cross-check |

A common mistake is to pass 100 and 30 directly as `MeanLog` and `StdDevLog`. The last formula confirms that the conversion was done.

## 8.6 Gamma claim severity and a Weibull lifetime

A claim size is Gamma with shape 3 and scale 2. What is its mean, its 99th percentile, and the probability of a claim above 15? A component lifetime is Weibull with shape 1.5 and scale 2000 hours: how likely is it to survive 1000 hours?

```text
=K_STATS_Gamma_Mean(3,2)
=K_STATS_Gamma_InverseCumulative(0.99,3,2)
=K_STATS_Gamma_Survival(15,3,2)
=K_STATS_Weibull_Survival(1000,1.5,2000)
```

In each formula the last two arguments are `Shape` and `ScaleParam`. The scale is not a rate: a tool that asks for a Gamma rate would take 0.5 here.

| Id | Result | Meaning |
|---|---|---|
| 8.6-a | 6 | Mean claim, shape times scale |
| 8.6-b | 16.81189383 | The 99th percentile |
| 8.6-c | 0.02025671506 | $P(X > 15)$ |
| 8.6-d | 0.7021885013 | Probability of surviving 1000 hours |

## 8.7 Binomial default counts

Twenty independent loans each default with probability 0.35. How likely are exactly 7 defaults, and more than 14? Among 100 loans with default probability 0.02, which number of defaults is not exceeded with 99 percent probability?

```text
=K_STATS_Binomial_PMF(7,20,0.35)
=K_STATS_Binomial_Survival(14,20,0.35)
=K_STATS_Binomial_InverseCumulative(0.99,100,0.02)
=K_STATS_Binomial_LogPMF(7,20,0.35)
```

The first argument is `NumberSuccesses`, or `Probability` for the inverse. Then come `Trials` and `ProbSuccess`. A "success" is whatever is counted: here, a default.

| Id | Result | Meaning |
|---|---|---|
| 8.7-a | 0.1844011864 | $P(X = 7)$ |
| 8.7-b | 0.0003105749704 | $P(X > 14)$, that is 15 or more |
| 8.7-c | 6 | Smallest $k$ with $P(X \le k) \ge 0.99$ |
| 8.7-d | -1.690641534 | Natural log of result 8.7-a |

For "14 or more", pass 13 to the survival function.

## 8.8 Event counts and sampling

```text
=K_STATS_Poisson_Cumulative(12,8.5)
=K_STATS_Poisson_Survival(60,2)
=K_STATS_Geometric_Cumulative(4,0.2)
=K_STATS_NegativeBinomial_Cumulative(5,3,0.4)
=K_STATS_Hypergeometric_PMF(4,10,20,100)
```

- Poisson: `NumberEvents`, then `Mean`.
- Geometric: `NumberFailures`, then `ProbSuccess`.
- Negative Binomial: `NumberFailures`, `NumberSuccesses`, `ProbSuccess`.
- Hypergeometric: `SampleSuccesses`, `SampleSize`, `PopulationSuccesses`, `PopulationSize`. A sample of 10 is drawn without replacement from 100 items, of which 20 have an error.

| Id | Result | Meaning |
|---|---|---|
| 8.8-a | 0.9090829257 | At most 12 events when 8.5 are expected |
| 8.8-b | 6.352854125E-67 | More than 60 events when 2 are expected |
| 8.8-c | 0.67232 | At most 4 failures before the first success |
| 8.8-d | 0.68460544 | At most 5 failures before the 3rd success |
| 8.8-e | 0.08410730453 | Exactly 4 of the 10 sampled items have an error |

## 8.9 Reproducing the results

Run `python3 example_values.py` in the `checks` folder. It needs Python 3 with the `mpmath` package and prints one line per result in this chapter and in sections 4.3, 4.4, 5.7, 7.3, 7.4, 7.5 and 7.6.
