# Independent Code Review — VBA Probability Distributions

> **Repository:** [`danielep71/VBA-PROBABILITY-DISTRIBUTIONS`](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS)  
> **Branch reviewed:** `main`  
> **Commit reviewed:** [`40624eecaab9765f15f0e52233e93fbceeecb2f0`](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/commit/40624eecaab9765f15f0e52233e93fbceeecb2f0)  
> **Review date:** 2026-08-11  
> **Reviewer:** OpenAI GPT-5.6 Pro — extra-high reasoning configuration  
> **Suggested repository path:** `docs/INDEPENDENT_CODE_REVIEW_2026-08-11.md`

---

## 1. Executive assessment

### Overall repository score: **8.9 / 10**

### Production numerical library score: **8.7 / 10**

### Numerical-assurance and release-engineering score: **9.2 / 10**

The repository is an advanced native-VBA numerical probability library, not a thin worksheet-function wrapper. It provides a coherent scalar probability stack with:

- **112 worksheet-facing functions**;
- seventeen univariate distribution surfaces;
- shared floating-point safeguards;
- reusable beta, gamma, log-gamma, log-beta and combinatorial kernels;
- direct survival APIs;
- direct inverse-survival functions in the Normal family;
- safeguarded inverse solvers;
- explicit numerical-domain and worksheet-error contracts;
- a consolidated Excel/VBA regression harness;
- source-bound external accuracy evidence;
- a strict hosted accuracy gate;
- an independent holdout suite;
- detailed numerical provenance and limitations documentation.

The current contract registry contains **162 active contracts**. The principal observation grid contains **2,030 rows**, all tied to the current numerical source through a content manifest. The committed accuracy summary reports no failed or pending contracts, and the independent holdout reports **80 PASS, 0 FAIL**.

The strongest parts of the repository are:

- the one-way numerical architecture;
- direct-tail design;
- guarded arithmetic;
- cancellation-resistant large-shape density kernels;
- paired incomplete-beta arguments;
- direct incomplete-gamma `P` and `Q`;
- honest measured envelopes for Student t, Chi-square and F probability surfaces;
- a complete six-family discrete layer;
- strict rejection of incomplete benchmark evidence;
- source, grid, schema and contract provenance;
- a real desktop-Excel CI workflow.

A release-blocking correctness defect nevertheless remains in the accepted Gamma and Chi-square domains:

> **A positive standardized variate that underflows to zero is treated as the exact support boundary. For small shape or degrees of freedom, the true probability can be large and fully representable.**

For example:

```text
K_STATS_Gamma_Cumulative(1E-300, 1E-4, 1E100)
```

forms the mathematical standardized argument `1E-400`, which is below the smallest positive Double. The shared division helper returns success with a rounded result of zero. The public Gamma CDF then passes zero to the incomplete-gamma kernel, which returns the support-boundary value zero. The 100-digit reference is approximately:

```text
0.9120634760684959678789255
```

The corresponding survival function returns one instead of approximately:

```text
0.0879365239315040321210745
```

The Gamma density path reaches `Log(0)` inside the stable deviance helper and is expected to return `#VALUE!`, even though the true density is finite:

```text
9.1206347606849596788E+295
```

The same structural problem occurs in Chi-square functions when `X / 2` underflows. The defect is outside the current grid because the benchmark contains large-shape, large-degree and tail regimes, but no **positive standardized-ratio-underflow** regime.

Additional hardening is required for subnormal positive shape parameters, Gamma inverse seeding, root-README generation, holdout provenance, and contract headroom governance.

### Independent verdict

> **A numerically sophisticated and professionally governed VBA library, suitable for controlled use inside its measured contracts. It should not yet be described as production-ready over its complete accepted public domain because a valid positive Gamma/Chi-square input can silently receive a materially wrong boundary probability.**

---

# 2. Review scope and methodology

## 2.1 Exact source basis

The review was performed against the exact `main` revision identified above. Files were retrieved by repository path and commit SHA.

The reviewed scope included:

### Production VBA

- [`src/M_STATS_PROBDIST_CORE.bas`](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/blob/40624eecaab9765f15f0e52233e93fbceeecb2f0/src/M_STATS_PROBDIST_CORE.bas)
- [`src/M_STATS_PROBDIST_SPECIALFUNCS.bas`](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/blob/40624eecaab9765f15f0e52233e93fbceeecb2f0/src/M_STATS_PROBDIST_SPECIALFUNCS.bas)
- [`src/M_STATS_PROBDIST_NORMALFAMILY.bas`](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/blob/40624eecaab9765f15f0e52233e93fbceeecb2f0/src/M_STATS_PROBDIST_NORMALFAMILY.bas)
- [`src/M_STATS_PROBDIST_TFAMILY.bas`](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/blob/40624eecaab9765f15f0e52233e93fbceeecb2f0/src/M_STATS_PROBDIST_TFAMILY.bas)
- [`src/M_STATS_PROBDIST_CONTINUOUS.bas`](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/blob/40624eecaab9765f15f0e52233e93fbceeecb2f0/src/M_STATS_PROBDIST_CONTINUOUS.bas)
- [`src/M_STATS_PROBDIST_DISCRETE.bas`](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/blob/40624eecaab9765f15f0e52233e93fbceeecb2f0/src/M_STATS_PROBDIST_DISCRETE.bas)

### Tests and Excel automation

- [`tests/M_STATS_PROBDIST_TEST.bas`](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/blob/40624eecaab9765f15f0e52233e93fbceeecb2f0/tests/M_STATS_PROBDIST_TEST.bas)
- [`ci/Run-ExcelVbaTests.ps1`](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/blob/40624eecaab9765f15f0e52233e93fbceeecb2f0/ci/Run-ExcelVbaTests.ps1)
- [`.github/workflows/excel-vba-regression.yml`](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/blob/40624eecaab9765f15f0e52233e93fbceeecb2f0/.github/workflows/excel-vba-regression.yml)

### Numerical assurance

- [`.github/workflows/accuracy-gate.yml`](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/blob/40624eecaab9765f15f0e52233e93fbceeecb2f0/.github/workflows/accuracy-gate.yml)
- [`benchmark/accuracy_contracts.csv`](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/blob/40624eecaab9765f15f0e52233e93fbceeecb2f0/benchmark/accuracy_contracts.csv)
- [`benchmark/probability_accuracy_grid.csv`](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/blob/40624eecaab9765f15f0e52233e93fbceeecb2f0/benchmark/probability_accuracy_grid.csv)
- [`benchmark/accuracy_summary.md`](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/blob/40624eecaab9765f15f0e52233e93fbceeecb2f0/benchmark/accuracy_summary.md)
- [`benchmark/compute_errors.py`](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/blob/40624eecaab9765f15f0e52233e93fbceeecb2f0/benchmark/compute_errors.py)
- [`benchmark/_contract_eval.py`](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/blob/40624eecaab9765f15f0e52233e93fbceeecb2f0/benchmark/_contract_eval.py)
- [`benchmark/_manifest.py`](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/blob/40624eecaab9765f15f0e52233e93fbceeecb2f0/benchmark/_manifest.py)
- [`benchmark/observation_manifest.json`](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/blob/40624eecaab9765f15f0e52233e93fbceeecb2f0/benchmark/observation_manifest.json)
- [`benchmark/refresh_evidence.py`](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/blob/40624eecaab9765f15f0e52233e93fbceeecb2f0/benchmark/refresh_evidence.py)
- [`benchmark/holdout/analyze_holdout.py`](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/blob/40624eecaab9765f15f0e52233e93fbceeecb2f0/benchmark/holdout/analyze_holdout.py)
- [`benchmark/holdout/holdout_summary.md`](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/blob/40624eecaab9765f15f0e52233e93fbceeecb2f0/benchmark/holdout/holdout_summary.md)
- [`benchmark/numerical_limitations.csv`](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/blob/40624eecaab9765f15f0e52233e93fbceeecb2f0/benchmark/numerical_limitations.csv)
- [`benchmark/check_source_thresholds.py`](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/blob/40624eecaab9765f15f0e52233e93fbceeecb2f0/benchmark/check_source_thresholds.py)
- the focused large-shape, envelope, inverse, survival and seam studies.

### Repository documentation and policy

- [`README.md`](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/blob/40624eecaab9765f15f0e52233e93fbceeecb2f0/README.md)
- [`benchmark/README.md`](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/blob/40624eecaab9765f15f0e52233e93fbceeecb2f0/benchmark/README.md)
- [`docs/EXCEL_VBA_CI.md`](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/blob/40624eecaab9765f15f0e52233e93fbceeecb2f0/docs/EXCEL_VBA_CI.md)
- [`SECURITY.md`](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/blob/40624eecaab9765f15f0e52233e93fbceeecb2f0/SECURITY.md)
- contribution and repository-policy files.

## 2.2 Execution boundary

Desktop Excel was not available in the review environment. The reviewer therefore did **not**:

- import the `.bas` modules into the Visual Basic Editor;
- execute `Debug -> Compile VBAProject`;
- run `Test_STATS_PROBDIST_RunAll`;
- execute the self-hosted Excel workflow;
- regenerate the committed observations.

The review distinguishes among:

- **confirmed source behavior**, established directly from the committed control flow and IEEE-754 semantics;
- **committed numerical evidence**, produced by the repository's own Excel/Python pipeline;
- **targeted independent numerical analysis**, using 100-digit arithmetic;
- **operational state not independently evidenced**, such as live self-hosted-runner availability and branch-protection enforcement.

## 2.3 Independent calculations

The following were independently evaluated at 100 decimal digits:

- Gamma CDF, survival and density when `X / Scale` is below the smallest positive Double;
- Chi-square CDF, survival and density when `X / 2` underflows;
- Gamma, Beta and F densities at subnormal positive shape parameters.

The current-source result was inferred only where the code path is deterministic:

- `PROB_TryDivide` explicitly treats underflow to zero as success;
- incomplete-gamma kernels explicitly return boundary values at zero;
- the density helper's stated precondition requires a positive standardized argument;
- `Log(0)` in VBA is a runtime error routed to the public `#VALUE!` handler.

These findings should still be converted into permanent Excel/VBA regression cases.

---

# 3. Hard repository metrics

## 3.1 Production surface

| Family | Worksheet-facing functions |
|---|---:|
| Standard Normal | 7 |
| General Normal | 7 |
| Lognormal | 9 |
| Student t | 4 |
| Chi-square | 4 |
| F | 4 |
| Gamma | 7 |
| Beta | 7 |
| Exponential | 4 |
| Weibull | 7 |
| Continuous Uniform | 4 |
| Binomial | 8 |
| Poisson | 8 |
| Geometric | 8 |
| Negative Binomial | 8 |
| Hypergeometric | 8 |
| Discrete Uniform | 8 |
| **Total** | **112** |

## 3.2 Code scale

The current codebase contains approximately:

```text
23,400 physical lines of production VBA
 6,400 physical lines in the consolidated VBA test module
29,800 production-plus-test physical lines
```

These figures include the repository's extensive comments and blank lines.

The largest modules are approximately:

```text
M_STATS_PROBDIST_DISCRETE.bas   8,000+ lines
M_STATS_PROBDIST_TEST.bas       6,400 lines
M_STATS_PROBDIST_CONTINUOUS.bas 5,000 lines
```

## 3.3 Test and assurance scale

| Artifact | Current state |
|---|---:|
| Registered VBA test sections | **97** |
| Assertions reported in the root README | **835** |
| Active accuracy contracts | **162** |
| Main-grid observations | **2,030** |
| Main-grid FAIL | **0** |
| Main-grid PENDING | **0** |
| Independent holdout contracts | **80** |
| Holdout PASS / FAIL | **80 / 0** |
| Source-bound `.bas` modules | **26** |
| Registered numerical limitations | **6** |
| GitHub Actions workflows | **2** |
| Open GitHub issues | **1** |

The assertion count is repository-reported; it was not independently executed during this review.

## 3.4 Current provenance

The observation manifest records:

```text
Generated UTC:    2026-08-11T17:04:16Z
Source commit:    18cf1f1
Excel version:    16.0
Excel build:      20131
Office bitness:   64-bit
Bound .bas files: 26
```

The content hashes bind:

- production source;
- tests;
- benchmark export modules;
- the main observation grid;
- the contract registry;
- the grid schema.

---

# 4. Scoring methodology

A score of 10 requires:

- correct behavior throughout the accepted public domain;
- no known silent wrong-result path;
- explicit error and convergence contracts;
- deterministic regression coverage;
- source-bound independent accuracy evidence;
- current-source execution in CI;
- complete generated-document freshness checks;
- reproducible performance evidence;
- maintainable module boundaries.

## Weighted scorecard

| Area | Weight | Score | Weighted contribution |
|---|---:|---:|---:|
| Functional correctness | 18% | **8.1** | 1.458 |
| Numerical robustness | 17% | **8.6** | 1.462 |
| Architecture and modularity | 10% | **9.7** | 0.970 |
| Public API design | 8% | **9.4** | 0.752 |
| Error handling and diagnostics | 7% | **8.6** | 0.602 |
| Regression testing | 10% | **9.3** | 0.930 |
| External accuracy assurance | 11% | **9.4** | 1.034 |
| CI and release engineering | 8% | **9.2** | 0.736 |
| Documentation and governance | 5% | **8.3** | 0.415 |
| Maintainability and repository hygiene | 5% | **8.7** | 0.435 |
| Performance engineering | 1% | **8.0** | 0.080 |
| **Total** | **100%** |  | **8.874 / 10** |

Rounded overall score:

```text
8.9 / 10
```

## Score interpretation

| Score | Interpretation |
|---:|---|
| 9.5-10.0 | Exceptional and independently release-certified |
| 9.0-9.4 | Advanced professional numerical library with limited gaps |
| 8.0-8.9 | Strong implementation requiring material hardening |
| 7.0-7.9 | Good foundation with significant correctness or assurance work |
| Below 7.0 | Major design, correctness or governance deficiencies |

---

# 5. Component scores

| Component | Score | Assessment |
|---|---:|---|
| `M_STATS_PROBDIST_CORE` | **9.4** | Strong shared primitives and exact-limit overflow reporting; division underflow needs richer classification |
| `M_STATS_PROBDIST_SPECIALFUNCS` | **8.7** | Sophisticated kernels and strong large-shape repairs; subnormal-shape paths remain unsafe |
| `M_STATS_PROBDIST_NORMALFAMILY` | **9.5** | Excellent direct tails, inverse survival, stable intervals and exact-limit standardization recovery |
| `M_STATS_PROBDIST_TFAMILY` | **8.5** | Strong envelopes and stable densities; Chi-square loses positive arguments when `X/2` underflows |
| `M_STATS_PROBDIST_CONTINUOUS` | **8.4** | Broad and mature; Gamma CDF/SF/density have a material positive-ratio-underflow defect |
| `M_STATS_PROBDIST_DISCRETE` | **9.5** | Complete six-family implementation with explicit domains, direct tails and holdout-backed contracts |
| `M_STATS_PROBDIST_TEST` | **9.2** | Exceptional VBA harness; missing ratio-underflow, subnormal-shape and tiny-Gamma-inverse regressions |
| External benchmark framework | **9.4** | Strict, source-bound and regime-aware; accepted-domain coverage is not exhaustive |
| Excel/PowerShell CI | **9.1** | Real Excel execution, complete module import, counters, failure log and trust boundary |
| Hosted accuracy CI | **9.3** | Evaluator tests, provenance checks, summary freshness and holdout rerun are automated |
| Documentation | **8.3** | Extensive and premium; root metrics and some source claims are not fully generated or synchronized |

---

# 6. Architectural review

## 6.1 Layering

The dependency architecture is clear:

```text
M_STATS_PROBDIST_CORE
        |
        v
M_STATS_PROBDIST_SPECIALFUNCS
        |
        +--> M_STATS_PROBDIST_NORMALFAMILY
        +--> M_STATS_PROBDIST_TFAMILY
        +--> M_STATS_PROBDIST_CONTINUOUS
        +--> M_STATS_PROBDIST_DISCRETE
                    |
                    v
          M_STATS_PROBDIST_TEST
```

The separation reflects numerical responsibility:

### Core

Owns:

- constants;
- finiteness and magnitude predicates;
- guarded arithmetic;
- `Log1p`, `Expm1`, `LogExpm1`;
- the raw inverse-normal seed;
- diagnostics.

### Special Functions

Owns:

- log-gamma and stable differences;
- log-beta;
- Stirling error;
- log-combination;
- Loader deviance;
- stable Gamma/Beta log densities;
- incomplete beta and gamma;
- inverse beta and gamma.

### Family modules

Own:

- parameterization;
- support behavior;
- public validation;
- tail orientation;
- worksheet errors;
- diagnostic text.

### Assurance layers

Own:

- deterministic VBA regression;
- Excel CI;
- high-precision references;
- contracts;
- holdouts;
- provenance;
- generated summaries.

## 6.2 Visibility

`Option Private Module` is used for Core and Special Functions. Project-scoped `PROB_*` helpers remain reusable by sibling modules without polluting the worksheet Function Wizard.

## 6.3 Architectural verdict

**9.7 / 10**

No architectural rewrite is needed. The principal concern is future file size, not numerical ownership.

---

# 7. Production-code review

## 7.1 `M_STATS_PROBDIST_CORE` — **9.4 / 10**

### Strengths

#### True finiteness versus project magnitude

The code distinguishes:

```text
PROB_IsFinite
PROB_IsWithinSupportedMagnitude
PROB_IsPositiveFinite
PROB_IsPositiveWithinSupportedMagnitude
```

This is the correct conceptual model. Representability, project policy, convergence and measured accuracy are different contracts.

#### Guarded arithmetic

The shared Try layer covers:

- exponential;
- addition;
- multiplication;
- division;
- affine reconstruction;
- standardization.

#### Signed standardization overflow

`PROB_TryStandardize` now reports an overflow direction, allowing the Normal family to return exact CDF/SF/density limits rather than treating every unrepresentable standardized value as `#NUM!`.

#### Stable elementary primitives

`PROB_Log1p`, `PROB_Expm1` and `PROB_LogExpm1` are used as first-class shared primitives.

### Material limitation

`PROB_TryDivide` intentionally treats arithmetic underflow to zero as success:

```text
finite result, including underflow to zero -> TRUE
```

That policy is valid for many calculations, but it discards whether:

- the mathematical quotient is exactly zero; or
- a positive quotient rounded to zero.

The distinction matters for Gamma and Chi-square, where a tiny positive standardized argument can produce a material CDF when the shape is also small.

### Recommended improvement

Introduce a richer positive-division contract, for example:

```text
FINITE_NONZERO
UNDERFLOW_TO_POSITIVE_ZERO
OVERFLOW
INVALID
```

or return an additional `Underflowed As Boolean`.

The ordinary `PROB_TryDivide` contract can remain unchanged for existing callers.

---

## 7.2 `M_STATS_PROBDIST_SPECIALFUNCS` — **8.7 / 10**

### Strengths

#### Iterative contracts

Every series, continued fraction and inverse solver returns Boolean success/failure. Partial sums are not returned as converged results.

#### Paired incomplete-beta arguments

The kernel accepts both `X` and `Y`, preserving the small complement.

#### Cancellation-resistant large-shape paths

The module now contains:

```text
PROB_TryDeviancePart
PROB_TryGammaLogPdf
PROB_TryBetaLogPdf
```

These remove the large-shape cancellation previously present in Gamma, Chi-square, Beta and F density/prefactor formulas.

#### Stable incomplete-gamma prefactor

The incomplete-gamma series and continued fraction reuse the stable Gamma log-density kernel.

#### Measured regime dispatch

`PROB_LogBeta` uses its stable delta route only where its preconditions hold.

### Finding: subnormal shape overflow

`PROB_StirlingError` handles `n < 0.5` through:

```text
delta(n+1) + (n+0.5) * Log((n+1)/n) - 1
```

For positive subnormal `n` below approximately `1 / DoubleMax`, `(n+1)/n` overflows before `Log` is evaluated.

The public validators accept every strictly positive shape below the upper magnitude guard. Consequently, accepted calls can reach an unhandled runtime overflow.

Representative finite reference values:

```text
K_STATS_Gamma_Density(1, 1E-310, 1)
    true density ≈ 3.678794411714423E-311

K_STATS_Gamma_Survival(1, 1E-310, 1)
    true survival ≈ 2.193839343955203E-311

K_STATS_Beta_Density(0.5, 1E-310, 1E-310)
    true density ≈ 2.0E-310

K_STATS_F_Density(1, 2E-310, 2E-310)
    true density ≈ 5.0E-311
```

The current paths are expected to return `#VALUE!`, not the finite values or a deliberate `#NUM!`.

The Gamma lower series has a second issue:

```text
SumValue = 1 / A
```

which overflows for similarly tiny positive `A`.

### Recommended correction

1. Replace:

```text
Log((n+1)/n)
```

with:

```text
PROB_Log1p(n) - Log(n)
```

or another ratio-free equivalent.

2. Rewrite the Gamma-series initialization in a scaled or logarithmic form for tiny `A`, or enforce an explicit lower kernel domain with a deliberate `#NUM!`.

3. Add subnormal-shape contracts and VBA regressions.

---

## 7.3 `M_STATS_PROBDIST_NORMALFAMILY` — **9.5 / 10**

### Strengths

- complete Standard Normal, Normal and Lognormal public surface;
- direct survival;
- direct inverse survival;
- stable same-tail interval probability;
- Acklam seed with refinement;
- deep-tail contracts;
- log-domain density reconstruction;
- stable Lognormal variance and standard deviation;
- stable moment-to-parameter conversion;
- exact-limit recovery on standardization overflow.

### Documentation inconsistency

One header section still says standardization overflow is caught and returned as `#NUM!`, while the later current section correctly states that CDF, SF and density return exact limits. The function-level behavior is clear, but the module-level prose should contain one authoritative description.

A second note says both arithmetic inputs to `K_STATS_Lognormal_ParametersFromMeanStdDev` must be strictly positive. The function correctly accepts `StdDev = 0` and returns a degenerate parameter pair.

These are documentation defects, not numerical defects.

---

## 7.4 `M_STATS_PROBDIST_TFAMILY` — **8.5 / 10**

### Strengths

#### Student t

- stable half-step log-gamma difference;
- direct survival;
- central expansion;
- exact low-degree branches;
- small-degree beta inversion;
- measured large-degree density behavior.

#### Chi-square

- direct incomplete-gamma `P` and `Q`;
- measured probability envelope;
- stable large-shape density helper;
- guarded inverse rescaling.

#### F

- stable logistic pair;
- explicit complement preservation;
- stable Beta log-density transformation;
- measured probability envelope;
- direct survival;
- representability-aware inverse behavior.

### Finding: `X / 2` underflow

Chi-square CDF and survival pass:

```text
0.5 * X
```

to the incomplete-gamma kernel. Chi-square density passes the same rounded quantity to the stable Gamma log-density helper.

For the smallest positive subnormal Double:

```text
X = 4.9406564584124654E-324
```

the multiplication by `0.5` rounds to zero.

With:

```text
DegreesFreedom = 0.0002
```

the mathematical CDF is approximately:

```text
0.92824867943255041433481
```

but the current CDF path returns the zero boundary. The current survival path returns one rather than approximately:

```text
0.07175132056744958566519004
```

For:

```text
DegreesFreedom = 1.99
```

the true density at the same positive `X` is finite:

```text
20.6892082696048338
```

but the density helper receives a zero standardized argument and is expected to reach the unexpected-runtime-error path.

### Recommended correction

Form and preserve:

```text
LogHalfX = Log(X) - Log(2)
```

before the product can underflow.

The incomplete-gamma and stable density APIs need a log-argument fallback when the rounded standardized value is zero but the original input is positive.

---

## 7.5 `M_STATS_PROBDIST_CONTINUOUS` — **8.4 / 10**

### Strengths

- complete Gamma, Beta, Exponential, Weibull and Uniform API;
- direct Gamma and Beta survival;
- paired Beta arguments;
- stable large-shape Gamma/Beta density kernels;
- density shape envelope;
- stable Exponential and Weibull left tails;
- stable Weibull large-shape moments;
- full-finite-range continuous Uniform formulas;
- explicit support-edge and error contracts.

### Release-blocking Gamma defect

The public Gamma CDF and survival compute:

```text
StandardX = X / ScaleParam
```

through a helper that regards underflow to zero as successful.

They distinguish only:

- successful division; and
- overflow/failure.

They do not distinguish a positive quotient that rounded to zero.

For:

```text
X          = 1E-300
Shape      = 1E-4
ScaleParam = 1E100
```

the mathematical standardized argument is:

```text
1E-400
```

The current path passes zero into the incomplete-gamma kernel.

Independent references:

| Function | Current source path | 100-digit reference |
|---|---:|---:|
| Gamma CDF | `0` | `0.9120634760684959678789255` |
| Gamma survival | `1` | `0.0879365239315040321210745` |
| Gamma density | expected `#VALUE!` | `9.1206347606849596788E295` |

The input is valid under every documented public validation rule.

### Recommended numerical branch

Compute first:

```text
LogStandardX = Log(X) - Log(ScaleParam)
```

If the rounded quotient is zero while `X > 0`, use a log-domain small-argument path.

For a standardized argument below the smallest positive Double, higher-order powers of the argument are themselves unrepresentable. The leading incomplete-gamma term is sufficient in Double:

```text
LogP = Shape * LogStandardX - LogGamma(Shape + 1)
P    = Exp(LogP)
Q    = -Expm1(LogP)
```

The density can be assembled as:

```text
LogDensity =
    (Shape - 1) * LogStandardX
    - Log(ScaleParam)
    - LogGamma(Shape)
```

because `Exp(-StandardX)` rounds to one in that regime.

The exact branch and crossover should be verified in a dedicated benchmark study.

---

## 7.6 `M_STATS_PROBDIST_DISCRETE` — **9.5 / 10**

### Strengths

The module provides six complete families:

- Binomial;
- Poisson;
- Geometric;
- Negative Binomial;
- Hypergeometric;
- Discrete Uniform.

Every family exposes:

- PMF;
- LogPMF;
- CDF;
- survival;
- inverse CDF;
- mean;
- variance;
- standard deviation.

### Numerical quality

- Loader-style Binomial and Poisson masses;
- LogPMFs for underflowed probabilities;
- direct incomplete-beta and incomplete-gamma tails;
- tail-oriented integer inverses;
- explicit exact-integer policy;
- explicit kernel-backed limits;
- finite-population validation;
- bounded Hypergeometric ratio summation;
- signed Discrete Uniform support;
- real-threshold step behavior;
- direct Discrete Uniform survival;
- corrected discrete quantile jumps.

### Assurance quality

All six families have:

- VBA regression coverage;
- main-grid contracts;
- independent holdout evidence.

### Maintainability concern

At more than 8,000 lines, the module is approaching a practical reviewability boundary. A future split should preserve one numerical source of truth and the public `K_STATS_*` API.

---

# 8. Public API and error-contract review

## 8.1 Naming

The public naming convention is consistent:

```text
K_STATS_<Distribution>_<Operation>
```

## 8.2 Return model

Most worksheet-facing functions return `Variant`:

| Outcome | Return |
|---|---|
| Valid calculation | `Double` |
| Invalid mathematical domain | `#NUM!` |
| Unsupported measured domain | `#NUM!` |
| Predictable overflow | `#NUM!` |
| Density pole | `#NUM!` |
| Non-convergence | `#NUM!` |
| Unexpected runtime failure | `#VALUE!` |
| Valid probability underflow | `0` |

## 8.3 Diagnostics

The optional `ByRef Status` channel is consistently used. No numerical UDF raises a modal `MsgBox`.

## 8.4 Current contract gap

The subnormal-shape and positive-ratio-underflow paths convert predictable floating-point conditions into:

- silently wrong boundary values; or
- unexpected `#VALUE!`.

That is inconsistent with the library's otherwise strong public contract.

---

# 9. Regression-test review

## 9.1 Strengths

The consolidated harness has **97 registered test sections**:

```text
Core and Special Functions     8
Normal and Lognormal          19
Student t, Chi-square and F   18
Continuous distributions     27
Discrete distributions       25
```

Coverage includes:

- exact constants;
- known values;
- direct tails;
- complement identities;
- inverse round-trips;
- support boundaries;
- full-range Uniform cases;
- large-shape and tiny-shape historical regressions;
- exact worksheet error codes;
- diagnostic Status behavior;
- complete discrete-family registration.

## 9.2 Missing critical cases

Add permanent VBA tests for:

```text
Gamma_Cumulative(1E-300, 1E-4, 1E100)
Gamma_Survival(1E-300, 1E-4, 1E100)
Gamma_Density(1E-300, 1E-4, 1E100)

ChiSquare_Cumulative(min_subnormal, 2E-4)
ChiSquare_Survival(min_subnormal, 2E-4)
ChiSquare_Density(min_subnormal, 1.99)

Gamma_Density(1, 1E-310, 1)
Gamma_Survival(1, 1E-310, 1)
Beta_Density(0.5, 1E-310, 1E-310)
F_Density(1, 2E-310, 2E-310)
```

The expected behavior should be established from independent references and the chosen supported-domain policy.

## 9.3 Minimum assertion/suite inventory

The CI bridge rejects zero and inconsistent assertion counts. A stronger gate would also enforce:

- an expected suite inventory;
- a minimum assertion count;
- no unregistered private `Test_*` procedure.

---

# 10. External accuracy and provenance review

## 10.1 Strengths

The assurance layer is a major repository asset.

### Contract schema

Each contract records:

- function;
- regime;
- measure;
- metric;
- threshold;
- domain;
- provenance;
- status;
- evidence;
- notes.

### Strict evidence completeness

Active rows must be:

- present;
- non-`ERROR` unless an error is explicitly expected;
- parseable;
- matched to a valid reference;
- actually consumed by the evaluator.

### Shared evaluator semantics

The main gate and holdout use the same contract-evaluation primitives.

### Exact main-grid provenance

The manifest binds:

- source;
- tests;
- exporters;
- the exact grid content;
- the exact registry;
- schema;
- Excel environment.

### Current result

The committed summary reports:

```text
FAIL:                  0
KNOWN LIMITATION:      0
CHARACTERIZATION ONLY: 0
PENDING:               0
```

The separate limitations register contains six measured numerical boundaries.

## 10.2 Coverage versus accepted domain

A contract proves its documented regime. It does not automatically certify the complete range accepted by a public validator.

The current Gamma/Chi-square ratio-underflow defect demonstrates this distinction.

A new regime should be added:

```text
positive_ratio_underflow
```

with:

- Gamma density;
- Gamma CDF;
- Gamma survival;
- Chi-square density;
- Chi-square CDF;
- Chi-square survival.

## 10.3 Holdout provenance gap

The hosted workflow reruns the holdout analyzer against the current registry and checks that `holdout_summary.md` is fresh.

However, the holdout observation CSV is not bound to production source through the same manifest mechanism as the main grid.

The workflow comments explicitly acknowledge that the holdout rerun and summary diff do not bind the observations themselves to source.

### Recommendation

Extend the manifest with authoritative evidence artifacts:

```json
"evidence_artifacts": {
  "probability_accuracy_grid.csv": "sha256:...",
  "holdout/holdout_grid.csv": "sha256:...",
  "holdout_older/holdout_grid.csv": "sha256:...",
  "density_large_shape/density_large_shape_grid.csv": "sha256:...",
  "cdf_large_shape/cdf_large_shape_grid.csv": "sha256:..."
}
```

Generated summaries can be regenerated and diffed rather than manifest-bound.

## 10.4 Contract headroom governance

The sole open GitHub issue identifies 32 contracts with at least 50x main-grid margin.

Some are legitimate:

- exact or near-exact formulas;
- conservative defaults;
- contracts frozen from holdout rather than main-grid worst.

Others require tracing to determine whether the threshold is:

- intentionally conservative;
- inherited;
- stale;
- or too weak to detect a meaningful regression.

The issue is correctly framed: thresholds should not be tightened mechanically.

---

# 11. CI and release-engineering review

## 11.1 Excel regression workflow

The self-hosted Windows/Excel workflow:

- imports all production modules;
- imports the test module;
- injects a CI-only bridge;
- runs all five suites;
- returns machine-readable counters;
- extracts failed assertion details;
- rejects zero or inconsistent counters;
- uploads the result artifact;
- excludes untrusted fork execution;
- releases COM objects.

This is strong engineering for a runtime unavailable on GitHub-hosted runners.

## 11.2 Hosted accuracy workflow

The hosted workflow now runs:

- contract evaluator tests;
- manifest tests;
- strict main-grid evaluation;
- committed summary regeneration and diff;
- degradation tests;
- benchmark README regeneration and diff;
- source-threshold checks;
- holdout-analyzer tests;
- current-registry holdout evaluation and summary diff.

This closes several common assurance-process gaps.

## 11.3 Local evidence script mismatch

`refresh_evidence.py` states that `--check` verifies the same sequence as the hosted workflow.

In `--check` mode it skips the regeneration list, while the verify list does not explicitly:

- regenerate/diff `accuracy_summary.md`;
- regenerate/diff the benchmark README contract table.

The hosted workflow does both.

Therefore:

```text
python refresh_evidence.py --check
```

can report green while generated committed documentation is stale.

### Recommendation

Either:

- include generated-artifact freshness checks in `VERIFY`; or
- narrow the script's claim and provide a distinct `--ci-check` mode.

## 11.4 Operational evidence

The available GitHub status and workflow-run interfaces exposed no runs for the reviewed SHA.

This review therefore cannot independently confirm:

- a green Excel run on the exact head;
- a green hosted gate on the exact head;
- required branch-protection checks;
- current self-hosted-runner availability.

The workflow definitions are strong; the live operational state was simply not visible through the available interface.

---

# 12. Documentation review

## 12.1 Strengths

Documentation includes:

- detailed module headers;
- procedure contracts;
- algorithm provenance;
- public parameterization;
- supported domains;
- error policies;
- a polished README;
- benchmark methodology;
- limitations;
- CI operations;
- security reporting.

## 12.2 Root README drift

The current root README reports:

```text
161 active accuracy contracts
1,905 observation rows
0 known silent-wrong paths
```

The authoritative files currently contain:

```text
162 active contracts
2,030 observation rows
```

The root README also states:

> Every number above is regenerated by CI and fails the build if it drifts.

The hosted workflow regenerates and checks the **benchmark README**, not the root README.

The current drift demonstrates that the root assurance panel is not generated or enforced as claimed.

### Recommendation

Place the root assurance metrics inside generated markers and add a deterministic renderer:

```text
README metrics <- source/API inventory + contracts + grid + latest test artifact
```

Then run and diff it in CI.

Until the P1 finding is resolved, remove or qualify:

```text
0 known silent-wrong paths
```

## 12.3 Threshold checker remains incomplete

`check_source_thresholds.py` was expanded substantially, but the source still contains measured statements such as:

```text
LogGamma relative error below 6.1E-14
LogGammaHalfDiff at or below 2E-14
LogChoose at or below 3.2E-16
```

The checker does not currently catch every wording variant, including `below` and `at or below`.

### Recommendation

Do not attempt to enumerate every English formulation indefinitely.

Prefer one of:

1. generate measured-accuracy sections from the registry; or
2. prohibit scientific-notation accuracy numbers in production source comments except whitelisted constants and counterexample descriptions.

---

# 13. Security and platform assessment

No high-severity security vulnerability was identified in the production numerical source.

Positive controls include:

- no external DLL;
- no network access from production UDFs;
- no shell execution from production modules;
- no `WorksheetFunction` dependency;
- no modal UI;
- read-only hosted workflow permissions;
- isolated Excel workbook;
- fork trust boundary;
- explicit COM cleanup;
- a clear private vulnerability-reporting policy.

The self-hosted runner necessarily enables VBA project automation and should remain dedicated, patched and access-controlled.

---

# 14. Findings summary

| ID | Severity | Area | Finding |
|---|---|---|---|
| ICR-P1-01 | **P1** | Numerical correctness | Positive Gamma/Chi-square standardized arguments can underflow to zero and be returned as exact support-boundary probabilities |
| ICR-P2-01 | **P2** | Numerical contract | Subnormal positive shapes can overflow inside `PROB_StirlingError` and Gamma-series initialization |
| ICR-P2-02 | **P2** | Inverse error handling | Gamma inverse computes an unguarded Wilson-Hilferty seed before its small-shape fallback |
| ICR-P2-03 | **P2** | Documentation governance | Root README assurance metrics are stale and are not regenerated by CI as claimed |
| ICR-P2-04 | **P2** | Evidence provenance | Holdout observations are rerun but not source-bound like the main grid |
| ICR-P2-05 | **P2** | Contract governance | 32 contracts have excessive main-grid headroom and require evidence-based audit |
| ICR-P2-06 | **P2** | Documentation governance | Source-threshold checker misses existing measured-claim phrasings |
| ICR-P3-01 | **P3** | Local tooling | `refresh_evidence.py --check` does not fully reproduce generated-artifact freshness checks performed in CI |
| ICR-P3-02 | **P3** | Maintainability | Discrete and Test modules are approaching a practical reviewability boundary |
| ICR-P3-03 | **P3** | Static assurance | No complete source/API/test-registration/export-dispatch static gate was identified |
| ICR-P3-04 | **P3** | Performance | No reproducible Excel/VBA timing baseline is committed |
| ICR-P3-05 | **P3** | Operations | Current workflow and branch-protection status was not independently visible |

---

# 15. Detailed findings

## ICR-P1-01 — Positive standardized-ratio underflow becomes a false boundary

### Severity

**P1 — release blocking**

### Affected functions

```text
K_STATS_Gamma_Density
K_STATS_Gamma_Cumulative
K_STATS_Gamma_Survival
K_STATS_ChiSquare_Density
K_STATS_ChiSquare_Cumulative
K_STATS_ChiSquare_Survival
```

### Root cause

The public wrapper forms a positive standardized argument in binary64.

If the mathematical quotient is smaller than the minimum positive Double, the result rounds to zero.

The shared division contract regards that as a successful finite underflow. The downstream distribution kernel cannot distinguish it from a true support-boundary zero.

### Confirmed Gamma example

```text
X          = 1E-300
Shape      = 1E-4
ScaleParam = 1E100
X/Scale    = 1E-400 mathematically
X/Scale    = 0 in binary64
```

| Quantity | Current source path | 100-digit reference |
|---|---:|---:|
| CDF | `0` | `0.9120634760684959678789255` |
| Survival | `1` | `0.0879365239315040321210745` |
| Density | expected `#VALUE!` | `9.1206347606849596788E295` |

### Confirmed Chi-square example

```text
X = 4.9406564584124654E-324
df = 0.0002
X/2 rounds to 0
```

| Quantity | Current source path | 100-digit reference |
|---|---:|---:|
| CDF | `0` | `0.92824867943255041433481` |
| Survival | `1` | `0.07175132056744958566519004` |

For:

```text
X  = minimum positive subnormal
df = 1.99
```

the true density is approximately:

```text
20.6892082696048338
```

but the stable density helper receives a zero standardized argument.

### Why this matters

The error is:

- inside the documented accepted domain;
- silent for CDF and survival;
- material rather than sub-ulp;
- absent from the active contracts;
- inconsistent with the public numerical contract.

### Remediation

1. Preserve the logarithm of the positive standardized argument.
2. Distinguish exact zero from positive-underflow zero.
3. Add a log-argument fallback to Gamma probability and density kernels.
4. Add main-grid, holdout and VBA regression coverage.
5. Freeze a dedicated `positive_ratio_underflow` contract.

---

## ICR-P2-01 — Subnormal positive shape paths overflow unexpectedly

### Severity

**P2 — public error-contract and accepted-domain defect**

### Root cause 1: Stirling recurrence

For `n < 0.5`, the current recurrence forms:

```text
(n + 1) / n
```

before applying `Log`.

For positive `n` below approximately `1 / DoubleMax`, that ratio overflows.

### Root cause 2: Gamma series

The lower incomplete-gamma series initializes:

```text
1 / A
```

which also overflows for sufficiently small accepted positive `A`.

### Representative finite values

| Call | High-precision result |
|---|---:|
| Gamma density `(1, 1E-310, 1)` | `3.678794411714423E-311` |
| Gamma survival `(1, 1E-310, 1)` | `2.193839343955203E-311` |
| Beta density `(0.5, 1E-310, 1E-310)` | `2.0E-310` |
| F density `(1, 2E-310, 2E-310)` | `5.0E-311` |

The current paths are expected to reach unexpected runtime handling rather than returning those values or a deliberate supported-domain error.

### Remediation

- rewrite the recurrence without forming the ratio;
- scale or logarithmically reformulate the Gamma-series initialization;
- or establish and enforce a measured positive lower shape boundary.

---

## ICR-P2-02 — Gamma inverse seed can overflow before the fallback

### Severity

**P2 — predictable failure classified as unexpected runtime error**

### Current sequence

`PROB_TryGammaInvP` computes the Wilson-Hilferty seed:

```text
T = 1 - 1/(9A) + Z/Sqrt(9A)
X = A*T*T*T
```

before checking whether `A < 1` and replacing the seed.

For `A` below roughly `3E-156`, the cubic expression can overflow before the fallback condition is reached.

### Consequence

A valid positive shape can reach the outer `#VALUE!` error handler even though the issue is predictable and should be:

- handled by a small-shape seed;
- classified as a representability boundary; or
- returned deliberately as `#NUM!`.

### Remediation

Select the small-shape seed **before** evaluating Wilson-Hilferty, or guard every multiplication.

---

## ICR-P2-03 — Root README assurance metrics are not authoritative

### Severity

**P2 — public trust and documentation-governance defect**

The root README is currently behind the authoritative data:

| Metric | Root README | Authoritative file |
|---|---:|---:|
| Active contracts | 161 | 162 |
| Observation rows | 1,905 | 2,030 |

Its claim that every number is regenerated by CI is not supported by the workflow definition.

### Remediation

Generate and diff the root metrics panel in CI.

---

## ICR-P2-04 — Holdout observations are not source-bound

### Severity

**P2 — evidence-chain gap**

The holdout analyzer and summary are current-registry checked, which is strong.

However, the holdout observation inputs do not carry the same cryptographic source binding as the main grid.

### Remediation

Add holdout and other threshold-freezing input grids to the evidence manifest.

---

## ICR-P2-05 — Excessive contract margins require audit

### Severity

**P2 — regression sensitivity and governance**

The open issue correctly avoids automatic threshold tightening.

Closure should classify each high-margin contract as:

```text
deliberate headroom
holdout-derived
exact-formula floor
stale inherited value
too weak to detect a material regression
```

Every change should be evidence-led.

---

## ICR-P2-06 — Threshold-prose checker is incomplete

### Severity

**P2 — generated-document consistency**

The checker is much stronger than a simple single regex, but current source comments still contain exact measured thresholds in forms it does not reject.

### Remediation

Generate accuracy prose or prohibit non-whitelisted scientific-notation claims in source comments.

---

# 16. Prioritized remediation plan

## Release Gate 1 — Preserve positive ratio underflow

1. Add an underflow flag to positive division.
2. Preserve `Log(X/Scale)` and `Log(X/2)`.
3. Implement log-argument Gamma CDF/SF/density fallbacks.
4. Add VBA regressions.
5. Add main-grid and holdout regimes.
6. Re-export and rebind evidence.

## Release Gate 2 — Harden subnormal shape behavior

1. Rewrite the Stirling recurrence.
2. Scale the Gamma-series initial term.
3. guard the Gamma inverse seed.
4. define representability behavior.
5. add exact error-code tests.

## Release Gate 3 — Generate public metrics

1. render root README assurance counts;
2. compare in CI;
3. remove the unsupported “zero known silent-wrong paths” statement until closure.

## Release Gate 4 — Complete evidence provenance

Bind every authoritative holdout or threshold-freezing input artifact.

## Release Gate 5 — Complete static assurance

Add a parser/checker for:

- module names;
- duplicate procedures;
- `Option Explicit`;
- broken continuations;
- public API inventory;
- test registration;
- exporter dispatch;
- documentation drift.

## Release Gate 6 — Maintainability and performance

- plan a compatibility-preserving module split;
- add nonblocking performance benchmarks;
- publish environment-recorded trends.

---

# 17. Release-readiness assessment

## Suitable now within measured regimes

The library is suitable for:

- teaching;
- numerical demonstrations;
- controlled model-validation comparisons;
- quantitative prototypes;
- direct-tail work;
- discrete probability calculations inside the documented limits;
- governed internal use with a pinned commit and archived evidence.

## Not suitable for an unconditional full-domain claim

Do not claim correct behavior throughout all accepted positive Gamma/Chi-square arguments until the positive-ratio-underflow defect is closed.

## High-stakes use

For banking, actuarial, engineering or other governed use:

- pin the exact commit;
- archive the manifest and summary;
- archive Excel version/build;
- validate the actual parameter domain;
- review `numerical_limitations.csv`;
- exclude standardized-ratio-underflow cases;
- require both Excel and hosted gates.

---

# 18. Final verdict

The repository combines qualities rarely found together in pure VBA:

- transparent algorithms;
- reusable special functions;
- direct tails;
- safeguarded inverses;
- stable large-shape formulas;
- complete discrete coverage;
- explicit error semantics;
- source-bound accuracy evidence;
- independent holdouts;
- real Excel CI;
- extensive documentation.

The current engineering foundation is strong.

The main remaining blocker is a narrow but serious floating-point classification error: a positive standardized argument can disappear to zero before the probability kernel sees it. When the shape is small, the lost argument does **not** imply a negligible probability.

> **Final score: 8.9 / 10**  
> **Classification: advanced professional numerical library with one release-blocking accepted-domain defect and several targeted assurance and governance improvements remaining.**

---

# Appendix A — Current representative assurance evidence

| Area | Current evidence |
|---|---|
| Active contracts | 162 |
| Main observations | 2,030 |
| Main-grid verdict | 0 FAIL, 0 PENDING |
| Holdout | 80 PASS, 0 FAIL |
| Bound VBA modules | 26 |
| Numerical limitations | 6 |
| Excel environment | 16.0 build 20131, 64-bit |
| VBA test sections | 97 |
| Worksheet-facing UDFs | 112 |

---

# Appendix B — Suggested GitHub issues

1. `Preserve positive Gamma standardized arguments that underflow to zero`
2. `Add log-argument fallback for Chi-square X/2 underflow`
3. `Make StirlingError recurrence safe for subnormal positive shapes`
4. `Guard the Gamma inverse Wilson-Hilferty seed before small-shape fallback`
5. `Add positive-ratio-underflow contracts and holdout cases`
6. `Generate root README assurance metrics from authoritative artifacts`
7. `Bind holdout input grids in observation provenance`
8. `Complete the excessive-contract-margin audit`
9. `Replace source threshold prose with generated contract references`
10. `Make refresh_evidence.py --check reproduce all CI freshness checks`
11. `Add static exported-VBA API and test-registration validation`
12. `Publish a reproducible Excel/VBA performance baseline`

---

# Appendix C — Evidence confidence

| Conclusion | Confidence |
|---|---|
| Architecture and API inventory | High |
| Contract and workflow design | High |
| Main-grid and holdout committed verdicts | High as committed evidence |
| Gamma ratio-underflow CDF/SF defect | High |
| Gamma density error-path classification | High from source; desktop Excel confirmation still required |
| Chi-square half-X underflow defect | High |
| Subnormal-shape runtime-overflow path | High from source |
| Current desktop-Excel regression result | Not independently executed |
| Current branch-protection enforcement | Not independently verified |
| Performance assessment | Medium; algorithmic review only |
