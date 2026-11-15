# 1. Purpose and scope

## 1.1 What the library is

VBA Probability Distributions is a set of six VBA modules that you import into a workbook. The workbook then has 112 functions whose names start with `K_STATS_`. You call them from a cell, like a built-in function, or from VBA.

The functions cover seventeen distributions: eleven continuous and six discrete. Depending on the distribution, the library provides:

- the density (continuous) or the probability mass and its logarithm (discrete);
- the cumulative probability $P(X \le x)$;
- the survival probability $P(X > x)$, computed directly;
- the inverse cumulative function (the quantile);
- the inverse survival function (Normal family only);
- the mean, variance and standard deviation;
- a z-score, interval probabilities and a Lognormal parameter conversion.

## 1.2 What distinguishes it from built-in functions

The library does not call Excel's statistical worksheet functions. The calculations are VBA code that travels inside the workbook and can be read in the VBA editor.

- **Direct upper tails.** A survival function does not compute `1 - CDF`, which loses a small result (section 7.3).
- **One error convention.** A function returns a number or an Excel error value. Invalid input is not silently corrected (chapter 9).
- **Stated limits.** The accuracy each function is held to, and the regions where it is weaker, are recorded (chapter 10).

## 1.3 Appropriate use

The library is suited to:

- moderate-volume Excel calculations;
- teaching and methodological demonstrations;
- model-validation workpapers;
- transparent control implementations;
- environments where external runtimes cannot be installed;
- reusable numerical building blocks inside VBA projects.

## 1.4 What it is not intended to replace

- high-performance compiled Monte Carlo engines;
- BLAS/LAPACK linear algebra;
- arbitrary-precision libraries;
- industrial optimization platforms;
- complete statistical ecosystems;
- independently certified regulatory or safety-critical software.

VBA is interpreted. A calculation repeated over millions of scenarios is much slower than in a compiled language.

## 1.5 Scope limits

- Univariate distributions only. Multivariate distributions and matrix calculations are out of scope.
- No random-number generation, no parameter estimation, no goodness-of-fit tests.
- Every function takes single numbers and returns one result. The exception is `K_STATS_Lognormal_ParametersFromMeanStdDev`, which returns two values (section 5.5).
- Results are Doubles and subject to the limits of that number type.

## 1.6 Your responsibility

The library is not an independently certified numerical package. Its accuracy statements apply only to the functions, parameter regions and thresholds that were tested (chapter 10).

You remain responsible for validating the library for your own application, in particular for regulated, financial, actuarial, engineering and safety-critical work. When a result must be reproducible, record the release tag or commit identifier you used.
