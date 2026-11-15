# 7. Conventions

## 7.1 Probabilities

Probabilities are numbers between 0 and 1. Enter 5 percent as `0.05`, not as `5`. A density is not a probability and can be greater than 1.

## 7.2 Lower tail and upper tail

| Kind | Tail | Definition |
|---|---|---|
| `Cumulative` | lower | $P(X \le x)$ |
| `Survival` | upper | $P(X > x)$ |

The survival definition uses a strict inequality. For a discrete distribution, the survival function at $k$ does not include the probability of $k$ itself.

| You want | Use |
|---|---|
| $P(X \le k)$ | `Cumulative` at $k$ |
| $P(X < k)$ | `Cumulative` at $k - 1$ |
| $P(X > k)$ | `Survival` at $k$ |
| $P(X \ge k)$ | `Survival` at $k - 1$ |

The two rows that use $k - 1$ need $k$ to be at least 1. A negative count returns `#NUM!` (section 7.8), so at the lower end of the support give the answer directly: $P(X < 0) = 0$ and $P(X \ge 0) = 1$. In a formula whose threshold can reach zero:

```text
=IF(A2<=0,1,K_STATS_Poisson_Survival(A2-1,8.5))
```

## 7.3 Why survival functions exist

Mathematically, the survival probability is one minus the cumulative probability. Numerically, that subtraction can destroy the result. A Double carries about 16 significant digits, so a cumulative probability extremely close to 1 is stored as exactly 1. Subtracting it from 1 gives 0, although the true upper tail is a small positive number.

```text
=K_STATS_NormalStandard_Survival(15)
```

The result is `3.670966199E-51`. The expression `=1-K_STATS_NormalStandard_Cumulative(15)` returns `0`. The loss is gradual: the closer the cumulative probability is to 1, the fewer correct digits survive the subtraction.

Every distribution in the library has a `Survival` function that computes the upper tail directly.

**Rule:** for a small upper-tail probability, call `Survival`. Do not subtract a `Cumulative` result from 1.

## 7.4 Why inverse survival functions exist

The same problem appears on the way from a probability to a quantile. To find the threshold exceeded with a small probability $q$, you could ask for the quantile at $1 - q$. But $1 - q$ is stored as exactly 1 when $q$ is very small, and the information is lost before the inverse is called. An `InverseSurvival` function takes $q$ itself.

```text
=K_STATS_NormalStandard_InverseSurvival(1E-18)
```

The result is `8.757290349`.

Inverse survival functions exist for the Normal family only: `K_STATS_NormalStandard_InverseSurvival`, `K_STATS_Normal_InverseSurvival` and `K_STATS_Lognormal_InverseSurvival`. For the other distributions, use `InverseCumulative` at $1 - q$ and accept the loss of accuracy when $q$ is very small.

## 7.5 Quantiles

Every inverse function requires a probability strictly between 0 and 1. The values 0 and 1 return `#NUM!`. The only exception is the fast helper of section 5.8.

For a continuous distribution, `InverseCumulative` returns the $x$ for which $P(X \le x)$ equals the probability.

A discrete distribution jumps from one cumulative probability to the next, so most probabilities are not matched exactly by any count. The convention is:

> `InverseCumulative` returns the smallest integer $k$ for which $P(X \le k)$ is greater than or equal to the given probability.

Example, Binomial with 10 trials and success probability 0.3:

| Formula | Result |
|---|---|
| `=K_STATS_Binomial_Cumulative(3,10,0.3)` | 0.6496107184 |
| `=K_STATS_Binomial_Cumulative(4,10,0.3)` | 0.8497316674 |
| `=K_STATS_Binomial_InverseCumulative(0.8,10,0.3)` | 4 |

The smallest count whose cumulative probability reaches 0.8 is 4. Applying `Cumulative` to a returned quantile therefore gives a probability at least as large as the one you started with, usually not equal to it.

## 7.6 Discrete parameterization

**Counts.** Count arguments are truncated toward zero: `7.9` is treated as `7`. A negative count returns `#NUM!`.

**Binomial.** `NumberSuccesses`, `Trials`, `ProbSuccess`.

**Poisson.** `NumberEvents`, `Mean`. `Mean` is the Poisson intensity $\lambda$.

**Geometric.** `NumberFailures`, `ProbSuccess`. The random variable counts failures before the first success and has support $0, 1, 2, \ldots$ Its mass is

$$P(X = k) = p\,(1 - p)^k$$

By this definition, the first success occurs on trial number $k + 1$. If your source counts trials up to and including the first success, pass that number minus 1.

**Negative Binomial.** `NumberFailures`, `NumberSuccesses`, `ProbSuccess`. The random variable counts failures before the `NumberSuccesses`-th success. By this definition, that success occurs on trial number `NumberFailures` + `NumberSuccesses`. If your source counts trials, pass that number minus `NumberSuccesses`.

**Hypergeometric.** `SampleSuccesses`, `SampleSize`, `PopulationSuccesses`, `PopulationSize`. Sampling is without replacement.

**Discrete Uniform.** `X`, `LowerBound`, `UpperBound`. The support contains every integer from the truncated lower bound through the truncated upper bound, inclusive. The PMF returns 0 when `X` is not a whole number or lies outside the support. The cumulative and survival functions accept any real `X` and behave as step functions: `=K_STATS_DiscreteUniform_Cumulative(2.8,-3,5)` returns `0.6666666667`, because six of the nine integers from -3 to 5 are at most 2.8.

**LogPMF.** A `LogPMF` function returns the natural logarithm of the probability mass. It stays finite and accurate where the `PMF` is too small for a Double and is returned as 0. When the outcome has probability exactly zero, `LogPMF` returns `#NUM!`.

## 7.7 Continuous parameterization

| Distribution | Convention | Note |
|---|---|---|
| Normal | Mean and standard deviation | `StdDev` > 0 |
| Lognormal | Mean and std. dev. of $\ln X$ | Not the arithmetic moments |
| Student t, Chi-square | Positive real degrees of freedom | Integers not required |
| F | Numerator and denominator d.f. | Both strictly positive |
| Gamma | Shape and **scale** | Scale, not rate |
| Beta | Shapes `Alpha` and `Beta` | Support is 0 to 1 |
| Exponential | **Rate** `Lambda` | Rate, not scale |
| Weibull | Shape and scale | Both strictly positive |
| Uniform | Lower and upper bound | Lower < upper |

A different parameterization is the most common reason for a mismatch with another tool.

## 7.8 Values outside the support

**Continuous distributions** return the natural limit and no error. Below the support, the density and the cumulative probability are 0 and the survival probability is 1. Above it, the cumulative probability is 1 and the survival probability is 0.

**Discrete distributions** are stricter: a negative count returns `#NUM!`. So does a Binomial count above `Trials`, a Hypergeometric `SampleSuccesses` above `SampleSize`, and a Hypergeometric PMF or LogPMF for a count the population makes impossible. The Discrete Uniform distribution is the exception (section 7.6).

A result too small to be stored as a Double is returned as 0, which is the correct limiting value. A result too large returns `#NUM!`.
