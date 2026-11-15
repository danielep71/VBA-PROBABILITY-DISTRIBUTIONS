# 12. Glossary

## 12.1 Terms

**Accuracy contract.** A tested combination of a function, a parameter region and a threshold (chapter 10).

**Bitness.** Whether Office is the 32-bit or the 64-bit edition.

**Characterized.** A recorded limitation whose behaviour is measured and documented but not removed.

**Cumulative distribution function (CDF).** The function that gives $P(X \le x)$: the `Cumulative` functions.

**Degrees of freedom.** The parameter of the Student t and Chi-square distributions and the two parameters of the F distribution.

**Density.** For a continuous distribution, the function whose area gives probabilities. A density is a height, not a probability.

**Double.** The 64-bit floating-point number type of Excel and VBA. It carries about 16 significant digits.

**Error value.** A special Excel result such as `#NUM!` or `#VALUE!`.

**Exceedance probability.** The probability that a variable is greater than a given value; an upper-tail or survival probability.

**Guarantee.** The accuracy a release promises for an accuracy contract (section 10.1).

**Inverse cumulative function.** Returns the quantile for a cumulative probability: the `InverseCumulative` functions.

**Inverse survival function.** Returns the value exceeded with a given probability: the `InverseSurvival` functions.

**LogPMF.** The natural logarithm of a probability mass.

**Measured target.** The accuracy the code achieves today for an accuracy contract, reported as evidence (section 10.1).

**Mitigated.** A recorded limitation that the library handles by refusing the affected inputs.

**Moments.** The mean, variance and standard deviation.

**Parameterization.** The choice of parameters that describe a distribution, for example scale versus rate.

**Percentile.** A quantile expressed in percent. The 99th percentile is the quantile for 0.99.

**Probability mass function (PMF).** For a discrete distribution, the function that gives $P(X = k)$.

**Production modules.** The six `.bas` files that make up the library.

**Quantile.** The value below which a given share of the distribution lies. For a discrete distribution: the smallest count whose cumulative probability reaches the given probability.

**Rate.** The parameter `Lambda` of the Exponential distribution; the reciprocal of its mean.

**Regression harness.** The optional VBA test module delivered with the library.

**Relative error.** The absolute error divided by the size of the true value.

**Scale.** A parameter that stretches a distribution. The Gamma mean is shape times scale.

**Shape.** A parameter that changes the form of a distribution.

**Standardized value (z-score).** A value minus the mean, divided by the standard deviation.

**Status.** The optional last argument of most functions. It receives a diagnostic message on failure.

**Support.** The set of values a distribution can take.

**Survival function.** The function that gives $P(X > x)$: the `Survival` functions.

**Truncation toward zero.** Dropping the fractional part of a number: `7.9` becomes `7`.

**Variant.** The VBA data type that can hold a number or an error value.

## 12.2 Notation

| Notation | Meaning |
|---|---|
| $X$, $T$, $Z$ | A random variable |
| $x$, $t$, $z$ | A value of that variable |
| $k$ | A count |
| $p$ | A probability; the success probability |
| $q$ | An upper-tail probability |
| $\lambda$ | The Poisson mean |
| $\ln$ | Natural logarithm |
| d.f. | Degrees of freedom |
