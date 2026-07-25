# Independent Code Review — VBA Probability Distributions

> **Repository:** [`danielep71/VBA-PROBABILITY-DISTRIBUTIONS`](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS)  
> **Branch reviewed:** `main`  
> **Commit reviewed:** [`aba81dce4ddd2231957605fb9c02916c3bff9b22`](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/commit/aba81dce4ddd2231957605fb9c02916c3bff9b22)  
> **Review date:** 2026-07-25  
> **Reviewer:** OpenAI GPT-5.6 Pro — extra-high reasoning configuration  
> **Suggested repository path:** `docs/INDEPENDENT_CODE_REVIEW_2026-07-25.md`

---

## 1. Executive assessment

### Overall repository score: **8.9 / 10**

### Production numerical library score: **8.8 / 10**

### Numerical-assurance and release-engineering score: **9.1 / 10**

This repository is an advanced, native-VBA numerical probability library. It has a coherent layered architecture, **112 worksheet-facing functions**, reusable beta/gamma/combinatorial kernels, direct survival APIs, guarded arithmetic, safeguarded inverse solvers, explicit worksheet-error semantics, a large consolidated VBA regression harness, a source-bound external accuracy framework, independent holdout evidence, and Excel-driven CI.

The strongest engineering characteristics are:

- clean separation among elementary numerics, reusable special functions, distribution families, tests, benchmark tooling, and CI;
- direct right-tail calculations rather than indiscriminate subtraction from one;
- direct inverse-survival functions in the Normal family;
- paired complementary arguments in the incomplete-beta layer;
- separate lower and upper incomplete-gamma kernels;
- explicit handling of valid underflow, predictable overflow, unsupported domains, and unexpected runtime failure;
- Loader-style discrete mass calculations and LogPMFs;
- explicit exact-integer and kernel-backed domains for discrete functions;
- a measured F accuracy envelope rather than reliance on a local convergence indicator alone;
- 141 active accuracy contracts, all currently passing;
- 1,589 committed observations tied to exact source hashes, exact grid bytes, the exact contract registry, a recorded Excel environment, and a schema manifest;
- 80 independently held-out contract checks, all currently passing.

The repository is nevertheless **not fully production-ready across its complete accepted parameter domain**. A material density defect remains:

> The Gamma, Chi-square, Beta, and F density formulas evaluate very large logarithmic terms separately and then subtract them to obtain a modest log-density. For sufficiently large accepted shape or degree parameters, this causes catastrophic cancellation. The functions can silently return values wrong by many orders of magnitude, or return `#NUM!` for a mathematically finite density.

A direct IEEE-754 binary64 reproduction of the current formulas and current Lanczos constants produced the following representative results:

| Function and accepted input | Current-formula result | 100-digit reference | Effect |
|---|---:|---:|---|
| Gamma density, `Shape=1E16`, `X=1E16`, `Scale=1` | `1.0` | `3.989422804E-9` | relative error about `2.51E8` |
| Chi-square density, `df=2E16`, `X=2E16` | `1.603810891E-28` | `1.994711402E-9` | essentially complete loss |
| Beta density, `a=b=1E16`, `x=0.5` | `1.664279892E-81` | `1.128379167E8` | wrong by about 89 orders of magnitude |
| F density, `d1=d2=1E16`, `x=1` | `1.586013452E15` | `1.994711402E7` | relative error about `7.95E7` |
| F density, `d1=d2=1E20`, `x=1` | `#NUM!` from computed overflow | `1.994711402E9` | false overflow; true result is finite |

The issue is not that the code uses logarithms. The problem is that the current log formulas still combine terms of order `shape * log(shape)` whose leading parts cancel. A relative-error contract on `LogGamma(shape)` does not control the **absolute** error left after this cancellation.

### Independent verdict

> **A highly credible and unusually mature VBA numerical project, suitable for controlled use inside its measured accuracy domains. One release-blocking large-shape density defect and several assurance/maintainability improvements remain before the full accepted domain can be called production-ready.**

---

# 2. Review scope and evidence boundary

## 2.1 Files reviewed

The review examined the exact files committed at the stated SHA, including:

### Production VBA

- [`src/M_STATS_PROBDIST_CORE.bas`](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/blob/aba81dce4ddd2231957605fb9c02916c3bff9b22/src/M_STATS_PROBDIST_CORE.bas)
- [`src/M_STATS_PROBDIST_SPECIALFUNCS.bas`](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/blob/aba81dce4ddd2231957605fb9c02916c3bff9b22/src/M_STATS_PROBDIST_SPECIALFUNCS.bas)
- [`src/M_STATS_PROBDIST_NORMALFAMILY.bas`](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/blob/aba81dce4ddd2231957605fb9c02916c3bff9b22/src/M_STATS_PROBDIST_NORMALFAMILY.bas)
- [`src/M_STATS_PROBDIST_TFAMILY.bas`](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/blob/aba81dce4ddd2231957605fb9c02916c3bff9b22/src/M_STATS_PROBDIST_TFAMILY.bas)
- [`src/M_STATS_PROBDIST_CONTINUOUS.bas`](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/blob/aba81dce4ddd2231957605fb9c02916c3bff9b22/src/M_STATS_PROBDIST_CONTINUOUS.bas)
- [`src/M_STATS_PROBDIST_DISCRETE.bas`](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/blob/aba81dce4ddd2231957605fb9c02916c3bff9b22/src/M_STATS_PROBDIST_DISCRETE.bas)

### Tests and Excel automation

- [`tests/M_STATS_PROBDIST_TEST.bas`](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/blob/aba81dce4ddd2231957605fb9c02916c3bff9b22/tests/M_STATS_PROBDIST_TEST.bas)
- [`ci/Run-ExcelVbaTests.ps1`](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/blob/aba81dce4ddd2231957605fb9c02916c3bff9b22/ci/Run-ExcelVbaTests.ps1)
- [`.github/workflows/excel-vba-regression.yml`](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/blob/aba81dce4ddd2231957605fb9c02916c3bff9b22/.github/workflows/excel-vba-regression.yml)

### Accuracy and provenance framework

- [`.github/workflows/accuracy-gate.yml`](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/blob/aba81dce4ddd2231957605fb9c02916c3bff9b22/.github/workflows/accuracy-gate.yml)
- [`benchmark/accuracy_contracts.csv`](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/blob/aba81dce4ddd2231957605fb9c02916c3bff9b22/benchmark/accuracy_contracts.csv)
- [`benchmark/probability_accuracy_grid.csv`](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/blob/aba81dce4ddd2231957605fb9c02916c3bff9b22/benchmark/probability_accuracy_grid.csv)
- [`benchmark/accuracy_summary.md`](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/blob/aba81dce4ddd2231957605fb9c02916c3bff9b22/benchmark/accuracy_summary.md)
- [`benchmark/compute_errors.py`](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/blob/aba81dce4ddd2231957605fb9c02916c3bff9b22/benchmark/compute_errors.py)
- [`benchmark/_contract_eval.py`](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/blob/aba81dce4ddd2231957605fb9c02916c3bff9b22/benchmark/_contract_eval.py)
- [`benchmark/_manifest.py`](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/blob/aba81dce4ddd2231957605fb9c02916c3bff9b22/benchmark/_manifest.py)
- [`benchmark/observation_manifest.json`](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/blob/aba81dce4ddd2231957605fb9c02916c3bff9b22/benchmark/observation_manifest.json)
- [`benchmark/test_contract_eval.py`](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/blob/aba81dce4ddd2231957605fb9c02916c3bff9b22/benchmark/test_contract_eval.py)
- [`benchmark/test_manifest.py`](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/blob/aba81dce4ddd2231957605fb9c02916c3bff9b22/benchmark/test_manifest.py)
- [`benchmark/test_gate_degradation.py`](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/blob/aba81dce4ddd2231957605fb9c02916c3bff9b22/benchmark/test_gate_degradation.py)
- [`benchmark/holdout/analyze_holdout.py`](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/blob/aba81dce4ddd2231957605fb9c02916c3bff9b22/benchmark/holdout/analyze_holdout.py)
- [`benchmark/holdout/test_analyze_holdout.py`](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/blob/aba81dce4ddd2231957605fb9c02916c3bff9b22/benchmark/holdout/test_analyze_holdout.py)
- [`benchmark/holdout/holdout_summary.md`](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/blob/aba81dce4ddd2231957605fb9c02916c3bff9b22/benchmark/holdout/holdout_summary.md)
- [`benchmark/numerical_limitations.csv`](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/blob/aba81dce4ddd2231957605fb9c02916c3bff9b22/benchmark/numerical_limitations.csv)
- [`benchmark/render_contract_table.py`](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/blob/aba81dce4ddd2231957605fb9c02916c3bff9b22/benchmark/render_contract_table.py)
- [`benchmark/check_source_thresholds.py`](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/blob/aba81dce4ddd2231957605fb9c02916c3bff9b22/benchmark/check_source_thresholds.py)

### Repository documentation

- [`README.md`](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/blob/aba81dce4ddd2231957605fb9c02916c3bff9b22/README.md)
- [`benchmark/README.md`](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/blob/aba81dce4ddd2231957605fb9c02916c3bff9b22/benchmark/README.md)
- [`docs/EXCEL_VBA_CI.md`](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/blob/aba81dce4ddd2231957605fb9c02916c3bff9b22/docs/EXCEL_VBA_CI.md)

## 2.2 Execution boundary

Desktop Excel was not available in the review environment. The reviewer therefore did **not**:

- import the `.bas` files into the Visual Basic Editor;
- run `Debug -> Compile VBAProject`;
- execute `Test_STATS_PROBDIST_RunAll`;
- execute the self-hosted Excel workflow.

The review distinguishes among:

- source conclusions established directly from the committed formulas;
- assurance conclusions established directly from scripts, workflows, contracts, and manifests;
- committed numerical evidence generated by the repository;
- targeted independent binary64 reproductions of current formulas;
- operational matters that could not be independently confirmed, such as live self-hosted-runner availability and branch-protection configuration.

## 2.3 Independent numerical method used for the large-shape finding

The density investigation reproduced:

- the current `PROB_LogGamma` Lanczos coefficients;
- the current operation order;
- the current Gamma, Chi-square, Beta, and F log-density formulas;
- IEEE-754 binary64 arithmetic.

References were calculated with 100 decimal digits.

Python's `math` functions are not guaranteed to be bit-for-bit identical to VBA's implementation. That caveat is immaterial to the conclusion here: the observed discrepancies range from factors of tens of millions to roughly 89 orders of magnitude. A desktop-Excel regression should still be added before release.

---

# 3. Hard repository metrics

## 3.1 Source size

| Component | Physical lines |
|---|---:|
| `M_STATS_PROBDIST_CORE.bas` | **915** |
| `M_STATS_PROBDIST_SPECIALFUNCS.bas` | **1,634** |
| `M_STATS_PROBDIST_NORMALFAMILY.bas` | **3,667** |
| `M_STATS_PROBDIST_TFAMILY.bas` | **3,362** |
| `M_STATS_PROBDIST_CONTINUOUS.bas` | **4,956** |
| `M_STATS_PROBDIST_DISCRETE.bas` | **8,160** |
| **Production VBA total** | **22,694** |
| `M_STATS_PROBDIST_TEST.bas` | **6,223** |
| **Production plus primary tests** | **28,917** |
| Main `README.md` | **1,517** |
| Main benchmark grid | **1,589 observations** plus header |

These are physical line counts, including the project's extensive comments and blank lines.

## 3.2 Worksheet-facing API

| Family | Public functions |
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
| **Total worksheet-facing UDFs** | **112** |

## 3.3 Regression structure

The consolidated test driver registers **97 test sections**:

| Suite | Registered sections |
|---|---:|
| Core and Special Functions | 8 |
| Normal and Lognormal | 19 |
| Student t, Chi-square, and F | 18 |
| Other continuous distributions | 27 |
| Discrete distributions | 25 |
| **Total** | **97** |

This count refers to registered test-section procedures, not the number of individual assertions executed at runtime.

## 3.4 Numerical-assurance evidence

| Artifact | Current state |
|---|---:|
| Active accuracy contracts | **141** |
| Main-grid observations | **1,589** |
| Generated main-grid verdicts | **141 PASS** |
| Main-grid FAIL / PENDING | **0 / 0** |
| Independent holdout contracts | **80** |
| Independent holdout result | **80 PASS, 0 FAIL** |
| Bound `.bas` files in provenance manifest | **19** |
| Numerical-limitations entries | **4** |
| GitHub Actions workflows | **2** |

The provenance manifest binds:

- the exact normalized content of every relevant `.bas` module;
- the exact observation-grid content;
- the exact contract registry;
- the grid schema;
- Excel version `2606`;
- Excel build `16.0.20131.20152`;
- 64-bit Office.

---

# 4. Scoring methodology

A score of 10 requires more than good formulas. It requires:

- correct behavior throughout the accepted public domain;
- no known silent wrong-result path;
- explicit numerical contracts;
- complete deterministic regression coverage;
- independent accuracy evidence;
- source/evidence provenance;
- current-source execution in CI;
- generated-document consistency;
- reproducible performance evidence.

## Weighted scorecard

| Area | Weight | Score | Weighted contribution |
|---|---:|---:|---:|
| Functional correctness | 18% | **8.3** | 1.494 |
| Numerical robustness | 17% | **8.8** | 1.496 |
| Architecture and modularity | 10% | **9.6** | 0.960 |
| Public API design | 8% | **9.4** | 0.752 |
| Error handling and diagnostics | 7% | **9.3** | 0.651 |
| Regression testing | 10% | **9.3** | 0.930 |
| External accuracy assurance | 11% | **9.2** | 1.012 |
| CI and release engineering | 8% | **8.7** | 0.696 |
| Documentation | 5% | **8.4** | 0.420 |
| Maintainability and repository hygiene | 5% | **8.5** | 0.425 |
| Performance engineering | 1% | **8.2** | 0.082 |
| **Total** | **100%** |  | **8.918 / 10** |

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
| Below 7.0 | Major design, correctness, or governance deficiencies |

---

# 5. Component scores

| Component | Score | Assessment |
|---|---:|---|
| `M_STATS_PROBDIST_CORE` | **9.4** | Strong shared primitives, explicit contracts, and clean project-private scope |
| `M_STATS_PROBDIST_SPECIALFUNCS` | **9.1** | Sophisticated, reusable kernels and strong inverse design; callers still need cancellation-aware composite formulas |
| `M_STATS_PROBDIST_NORMALFAMILY` | **9.5** | Excellent tail orientation, inverse-tail API, interval logic, and guarded arithmetic |
| `M_STATS_PROBDIST_TFAMILY` | **8.6** | Strong CDF/SF/inverse design and measured F envelope; Chi-square and F densities fail at large accepted parameters |
| `M_STATS_PROBDIST_CONTINUOUS` | **8.5** | Broad and disciplined; Gamma and Beta densities have a release-blocking large-shape cancellation defect |
| `M_STATS_PROBDIST_DISCRETE` | **9.5** | Complete six-family layer with direct tails, LogPMFs, explicit domains, and external contracts |
| `M_STATS_PROBDIST_TEST` | **9.3** | Large, dependency-ordered, consolidated harness; large-shape density regressions are absent |
| External benchmark framework | **9.2** | Regime-aware, source-bound, complete, and independently held out; coverage remains uneven in some large-parameter regimes |
| Excel/PowerShell CI | **9.0** | Imports all production modules, runs every suite, captures counters and detailed failures |
| Hosted accuracy CI | **8.4** | Strict main gate and documentation checks are present, but several critical assurance unit tests are not executed |
| Documentation | **8.4** | Rich and technically useful; duplicated measured thresholds remain and one current source claim conflicts with the registry |

---

# 6. Architectural review

## 6.1 Layering

The production architecture is coherent:

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

The separation is substantive:

- **Core** owns elementary floating-point policy.
- **Special Functions** owns reusable distribution-independent kernels.
- **Family modules** own parameterization, support behavior, tail orientation, public errors, and diagnostics.
- **Tests** own one consolidated verdict.
- **Benchmark tooling** owns high-precision references, contracts, holdouts, and provenance.
- **CI tooling** owns automated execution and release gating.

## 6.2 Scope boundaries

`Option Private Module` is used appropriately in Core and Special Functions. Project-internal `PROB_*` names remain reusable by sibling modules but hidden from worksheet users.

This is an excellent compromise between:

- duplicating private helpers in every module; and
- exposing implementation internals as worksheet functions.

## 6.3 Dependency direction

The code generally maintains one-way dependency flow. Shared special functions are not duplicated inside public families. This materially improves numerical consistency and maintainability.

## 6.4 Architectural recommendation

Random variates and array/range wrappers should remain separate future modules:

```text
M_STATS_RNG_CORE
M_STATS_PROBDIST_RANDOM
M_STATS_PROBDIST_ARRAY
```

They should reuse scalar kernels rather than enlarging the current family modules or duplicating formulas.

Matrix, vector, decomposition, and multivariate work belongs in the separate `K_MAT_*` project boundary.

---

# 7. Production-code review

## 7.1 `M_STATS_PROBDIST_CORE` — **9.4 / 10**

### Strengths

#### One source of truth

Constants, predicates, arithmetic guards, stable elementary functions, the raw inverse-normal seed, and diagnostics are centralized.

#### Finiteness is separate from project policy

The code distinguishes:

```text
PROB_IsFinite
PROB_IsWithinSupportedMagnitude
```

A finite Double is not automatically inside a validated algorithmic domain.

#### Guarded arithmetic

The module provides:

- `PROB_TryExp`;
- `PROB_TryAdd`;
- `PROB_TryMultiply`;
- `PROB_TryDivide`;
- `PROB_TryAffineTransform`;
- `PROB_TryStandardize`.

Predictable arithmetic failure becomes a Boolean contract rather than an unexpected public runtime error.

#### Explicit underflow policy

Negative exponential underflow is treated as a valid zero. Positive overflow remains a failure.

#### Stable elementary helpers

`PROB_Log1p` and `PROB_Expm1` materially improve small-increment and small-left-tail calculations.

### Improvement opportunities

#### Standardization overflow loses sign

`PROB_TryStandardize` returns only success/failure. For CDF and survival functions, the sign of a mathematical overflow provides an exact limit:

| Function | Positive standardized overflow | Negative standardized overflow |
|---|---:|---:|
| Normal density | 0 | 0 |
| Normal CDF | 1 | 0 |
| Normal survival | 0 | 1 |
| Z-score | `#NUM!` | `#NUM!` |

A richer result enum could return exact probability limits rather than `#NUM!` for every standardization overflow.

#### Conservative General-Normal magnitude guard

The Normal module documents that guarded standardization makes the `1E100` argument restriction redundant for overflow protection. Relaxation should be evidence-led, but the current guard excludes finite, mathematically resolvable cases.

---

## 7.2 `M_STATS_PROBDIST_SPECIALFUNCS` — **9.1 / 10**

### Strengths

#### Real Try-contracts

Iterative routines return Boolean success and do not return unconverged partial results.

#### Paired incomplete-beta arguments

The incomplete-beta API receives both complementary arguments. This supports stable direct tails and prevents callers from reconstructing a tiny complement by subtraction.

#### Separate incomplete-gamma tails

Lower `P` and upper `Q` are first-class kernels.

#### Safeguarded inverses

Beta and Gamma inverses use:

- analytical/asymptotic seeds;
- Newton refinement;
- brackets;
- bisection fallback;
- explicit iteration caps;
- forward representability checks.

#### Stable LogBeta dispatch

`PROB_LogBeta` now:

- uses half-integer shortcuts where applicable;
- uses `PROB_LogGammaDelta` only inside its documented `LargeArg >= 1` regime;
- uses the direct identity for both-small shapes;
- has tiny/tiny regression and external contract coverage.

#### Discrete support kernels

`PROB_StirlingError` and `PROB_LogChoose` provide strong foundations for large-count discrete mass calculations.

### Important design limitation

`PROB_LogGamma` has a **relative** error contract. When its value is of size `shape * log(shape)`, even a small relative error can correspond to a large absolute error.

That is acceptable for `LogGamma` as a standalone function. It becomes unsafe when a caller subtracts several large log terms to produce a modest log-density.

The correct fix belongs in stable composite density kernels rather than in pretending that a generic relative LogGamma contract controls every downstream cancellation.

### Recommended reusable additions

```text
PROB_BD0
    Stable Loader-style deviance term

PROB_TryGammaLogDensity
    Cancellation-resistant Gamma log density

PROB_TryBetaLogDensity
    Cancellation-resistant Beta log density
```

Chi-square should reuse the Gamma helper. F should reuse the Beta helper plus a stable transformation Jacobian.

---

## 7.3 `M_STATS_PROBDIST_NORMALFAMILY` — **9.5 / 10**

### Strengths

- complete Standard Normal, Normal, and Lognormal API;
- direct survival and inverse survival;
- stable same-tail interval probability;
- Acklam inverse seed with guarded refinement;
- direct deep-tail paths;
- log-domain densities;
- stable Lognormal moments;
- parameter conversion using `Log1p`;
- clear `#NUM!` versus `#VALUE!` behavior;
- no modal UI.

### Notable quality

The module explicitly distinguishes:

- true finite arguments;
- conservative general-Normal magnitude policy;
- direct standard-Normal tails;
- valid exponential underflow;
- predictable reconstruction overflow.

The current deep-tail contracts and split-boundary holdouts are a strong assurance feature.

### Improvement opportunities

The general-Normal public layer could eventually map signed standardization overflow to exact CDF/SF limits, as described in the Core review.

---

## 7.4 `M_STATS_PROBDIST_TFAMILY` — **8.6 / 10**

### Strengths

#### Student t

- stable log-density normalization through `PROB_LogGammaHalfDiff`;
- direct survival;
- exact low-degree branches;
- central expansion;
- safeguarded inverse;
- small-degree beta-inversion branch.

#### Chi-square CDF/SF/inverse

- direct incomplete-gamma `P` and `Q`;
- no partial sum on non-convergence;
- guarded inverse rescaling.

#### F CDF/SF/inverse

- stable log-ratio pair;
- explicit complementary beta arguments;
- measured `df <= 1E5` public envelope;
- correctly handled expected `#NUM!` evidence outside that envelope;
- representability-aware inverse behavior.

### Material defect: Chi-square density

The current Chi-square density evaluates:

```vb
LogDensity = _
    (HalfDF - 1#) * Log(X) - _
    0.5 * X - _
    HalfDF * Log(2#) - _
    PROB_LogGamma(HalfDF)
```

For large `df` near the distribution's mode, the leading terms are enormous and cancel.

At:

```text
df = 2E16
x  = 2E16
```

the current binary64 formula gives:

```text
log density = -64
density     = 1.603810891E-28
```

The 100-digit reference is:

```text
log density = -20.0327664577
density     = 1.994711402E-9
```

The input passes the current `<1E100` degree validation.

### Material defect: F density

The current F density uses a stable ratio orientation but still combines:

```text
large degree-weighted logarithms
minus LogBeta(df1/2, df2/2)
```

At:

```text
df1 = df2 = 1E16
x   = 1
```

the current binary64 formula gives a density about `7.95E7` times too large.

At:

```text
df1 = df2 = 1E20
x   = 1
```

the current log-density is about `483328`, so the public function reports density overflow. The true finite density is about `1.994711402E9`.

The existing measured F envelope intentionally excludes `F_Density`, so it does not protect this path.

---

## 7.5 `M_STATS_PROBDIST_CONTINUOUS` — **8.5 / 10**

### Strengths

- broad and consistent public API;
- direct Gamma/Beta/Exponential/Weibull survival;
- guarded Gamma, Exponential, and Weibull arithmetic;
- stable Exponential and Weibull left tails;
- large-shape Weibull moment branch;
- full-finite-range continuous Uniform design;
- exact support-edge behavior;
- tiny/tiny and unbalanced Beta contracts.

### Material defect: Gamma density

The current formula is:

```vb
LogDensity = _
    (Shape - 1#) * LogRatio - _
    StandardX - _
    Log(ScaleParam) - _
    PROB_LogGamma(Shape)
```

This is algebraically correct but not cancellation-safe near `X / Scale = Shape`.

At:

```text
Shape = 1E16
X     = 1E16
Scale = 1
```

the current binary64 formula rounds the log-density to zero and returns density `1`. The correct density is about `3.989422804E-9`.

### Material defect: Beta density

The current formula is:

```vb
LogDensity = _
    (Alpha - 1#) * Log(X) + _
    (Beta - 1#) * Log1p(-X) - _
    PROB_LogBeta(Alpha, Beta)
```

At:

```text
Alpha = Beta = 1E16
X     = 0.5
```

the current formula returns approximately `1.664E-81`; the correct density is approximately `1.128E8`.

The current public shape validator accepts these values because both are below `1E100`.

### Why existing contracts do not catch this

The current density grids are intentionally small:

- Gamma density: 3 main-grid points;
- F density: 6 main-grid points;
- balanced Beta density: 3 main-grid points.

Those contracts are valid for their stated grids and regimes. They are not evidence for the full accepted `<1E100` shape domain.

---

## 7.6 `M_STATS_PROBDIST_DISCRETE` — **9.5 / 10**

### Strengths

The discrete layer is complete and coherent:

- Binomial;
- Poisson;
- Geometric;
- Negative Binomial;
- Hypergeometric;
- Discrete Uniform.

Each family exposes:

- PMF;
- LogPMF;
- CDF;
- direct survival;
- inverse CDF;
- mean;
- variance;
- standard deviation.

### Numerical strengths

#### Binomial and Poisson

- Loader-style mass calculations;
- LogPMFs;
- direct incomplete-beta/incomplete-gamma tails;
- bounded integer inverse searches.

#### Geometric

- stable closed forms using `Log1p`/`Expm1`;
- exact support behavior.

#### Negative Binomial

- stable combinatorial structure;
- direct incomplete-beta tails;
- explicit count domains.

#### Hypergeometric

- stable log-mass composition;
- near-tail successive-ratio summation;
- feasibility checks;
- explicit population and summation ceilings.

#### Discrete Uniform

- signed support;
- correct real-threshold step CDF/SF behavior;
- direct right tail;
- adjacent-step inverse correction;
- factored moment formulas;
- exact `2^53 - 1` count policy.

### Assurance quality

All six discrete families have:

- consolidated VBA regression coverage;
- main-grid accuracy contracts;
- independent holdout coverage.

Several discrete contracts have limited holdout margin, which warrants more evidence but not automatic threshold widening.

---

# 8. Regression-test review

## 8.1 Strengths

The 6,223-line test module is exceptional for a VBA numerical project.

It provides:

- one authoritative counter set;
- one final verdict;
- dependency-ordered suites;
- exact constants;
- known values;
- support boundaries;
- CDF/SF complements;
- inverse round-trips;
- direct-tail cases;
- moment identities;
- full-range Uniform cases;
- discrete inverse minimality;
- exact worksheet-error-code checks;
- named regression tests;
- machine-readable CI integration.

The tiny/tiny LogBeta regression is especially well constructed because it exercises:

- the kernel;
- argument symmetry;
- Beta density/CDF/SF;
- F density on both branches;
- F CDF/SF.

## 8.2 Missing critical regression

No permanent VBA regression was identified for the large-shape modal-density cancellation described above.

Add at least:

```text
Gamma_Density(1E16, 1E16, 1)
ChiSquare_Density(2E16, 2E16)
Beta_Density(0.5, 1E16, 1E16)
F_Density(1, 1E16, 1E16)
F_Density(1, 1E20, 1E20)
```

The final corrected values should be checked with an appropriate log-density or relative contract. The current formulas will fail these cases dramatically.

## 8.3 Test tolerance role

The VBA test suite correctly uses broader deterministic tolerances than some external contracts.

The distinction should remain explicit:

- **VBA suite:** behavioral and regression gate;
- **external benchmark:** high-precision measured-accuracy gate.

---

# 9. External benchmark review

## 9.1 Major strengths

### Regime-aware contracts

The schema separates:

- function;
- regime;
- measure;
- metric;
- threshold;
- provenance;
- status;
- evidence.

### Function-appropriate metrics

The framework supports:

- ordinary output error;
- quantile error;
- absolute log error;
- forward tail-probability residual;
- exact discrete inverse error.

### Strict row completeness

The shared evaluator blocks:

- missing observations;
- unexpected `ERROR`;
- malformed `hi;lo`;
- malformed references;
- missing residual arguments;
- expected-envelope rows that fail to return `#NUM!`;
- internal measurement skips.

### Single-sourced evaluation semantics

The main gate and holdout import the same parsing, metric, residual, and disposition logic.

### Exact provenance

The manifest binds:

- every relevant `.bas` file;
- exact grid bytes;
- exact contracts;
- schema;
- Excel environment.

### Holdout evidence

Eighty contracts are rerun on independent holdout data, all passing.

## 9.2 Coverage limitation exposed by the density finding

The benchmark is strong, but coverage density is uneven.

A green contract means:

> every listed point in the stated regime met the threshold

It does not mean:

> every value accepted by the public validator is accurate.

The large-shape density defect is a direct example. Current density contracts contain too few large-shape modal points to expose cancellation.

## 9.3 Thin holdout margins

Several contracts pass with less than two-times holdout headroom:

| Contract | Holdout margin |
|---|---:|
| `F_Density.all.output_rel` | 1.3x |
| `F_InverseCumulative.tiny_unbalanced_representable.tail_rel` | 1.4x |
| `Geometric_PMF.all.output_rel` | 1.7x |
| `Hypergeometric_PMF.all.output_rel` | 1.7x |
| `Hypergeometric_Survival.all.output_rel` | 1.3x |
| `Poisson_LogPMF.all.output_rel` | 1.5x |

These are valid PASS results. The correct next step is to expand independent points, not to widen thresholds automatically.

---

# 10. CI and release-engineering review

## 10.1 Excel regression workflow — strong

The PowerShell/COM runner:

- creates an isolated workbook;
- imports all six production modules;
- imports the consolidated test module;
- injects a CI-only bridge;
- runs all five suites;
- records Excel version/build;
- returns machine-readable counters;
- retrieves detailed failed assertions;
- rejects zero or inconsistent counts;
- releases COM objects.

The workflow also excludes untrusted fork pull requests from the privileged self-hosted Excel runner.

## 10.2 Hosted accuracy workflow — strong core, incomplete self-testing

The hosted workflow currently runs:

1. strict `compute_errors.py`;
2. reference-helper degradation test;
3. generated benchmark-README table check;
4. source-threshold duplication check.

It does **not** currently run:

- `benchmark/test_contract_eval.py`;
- `benchmark/test_manifest.py`;
- `benchmark/holdout/test_analyze_holdout.py`;
- `benchmark/holdout/analyze_holdout.py`.

These files protect critical guarantees:

- metric semantics;
- invalid evidence blocking;
- observation/source/contract provenance;
- expected-error classification;
- holdout consistency.

A future change can therefore break the assurance infrastructure without a CI failure, provided the main gate still happens to pass on the current data.

## 10.3 Generated summary freshness

The strict gate writes its output to:

```text
${RUNNER_TEMP}/accuracy_summary.md
```

The workflow does not compare that fresh output with the committed:

```text
benchmark/accuracy_summary.md
```

The committed summary can therefore become stale while CI remains green.

The same applies to `holdout_summary.md`, because the holdout analyzer is not rerun in the workflow.

## 10.4 Recommended hosted workflow

Add steps equivalent to:

```bash
python test_contract_eval.py
python test_manifest.py
python test_gate_degradation.py

cd holdout
python test_analyze_holdout.py
python analyze_holdout.py
git diff --exit-code holdout_summary.md
cd ..

python compute_errors.py --out accuracy_summary.generated.md
diff -u accuracy_summary.md accuracy_summary.generated.md

python render_contract_table.py --write
git diff --exit-code README.md

python check_source_thresholds.py
```

A single aggregate `verify_repository.py` entry point would make local and CI execution identical.

## 10.5 Operational evidence

The committed workflow definitions are strong. The available GitHub status interface exposed no status records for the reviewed SHA, so this review could not independently verify:

- that both workflows ran on the exact commit;
- that the self-hosted runner was online;
- that both checks are required by branch protection.

This is an operational verification gap, not evidence that the workflows are defective.

---

# 11. Documentation review

## 11.1 Strengths

Documentation is a major asset:

- detailed module headers;
- procedure contracts;
- algorithm provenance;
- parameterization;
- public error policy;
- supported domains;
- README visuals and API catalogue;
- benchmark methodology;
- numerical limitations;
- CI operating instructions.

## 11.2 Source-threshold checker is too narrow

`check_source_thresholds.py` claims to prevent measured thresholds from being duplicated in source, but its regex catches only the exact form:

```text
<= <scientific number> relative|absolute
```

It does not catch phrases such as:

- `below 6.1E-14`;
- `at or below 2E-14`;
- `~5E-15`;
- `density ... to 5E-15`;
- comma-separated threshold tables.

## 11.3 Current source drift

The Continuous module still states:

```text
Balanced Beta density and survival to 5E-15
```

The authoritative balanced Beta density contract is:

```text
1E-14
```

The same header duplicates several other contract numbers and omits the newer `tiny_unbalanced` regime.

The checker reports clean because those phrases do not match its narrow pattern.

### Recommended documentation policy

Source modules should state:

> Authoritative, regime-specific measured thresholds live in `benchmark/accuracy_contracts.csv`.

Exact measured numbers should be generated into benchmark documentation, not manually duplicated in source prose.

---

# 12. Maintainability review

## 12.1 Strengths

- consistent naming;
- disciplined comments;
- explicit preconditions;
- one-way dependencies;
- no WorksheetFunction dependency;
- no external DLL;
- no modal UI;
- generated contract tables;
- machine-readable evidence.

## 12.2 Module size

Two modules are approaching a practical reviewability boundary:

```text
M_STATS_PROBDIST_DISCRETE.bas   8,160 lines
M_STATS_PROBDIST_TEST.bas       6,223 lines
```

No split is needed merely for aesthetics. A split should preserve:

- public names;
- one source of numerical truth;
- dependency direction;
- CI import order.

Before adding random variates or array APIs, consider a measured boundary such as:

```text
M_STATS_PROBDIST_DISCRETE_COUNT
    Binomial
    Poisson
    Geometric
    Negative Binomial

M_STATS_PROBDIST_DISCRETE_FINITE
    Hypergeometric
    Discrete Uniform
```

The test module could mirror suite modules while retaining one public consolidated runner.

## 12.3 Static exported-VBA checks

A cheap hosted check could verify:

- unique module names;
- unique procedures;
- valid `Attribute VB_Name`;
- `Option Explicit`;
- no broken continuation;
- public API inventory;
- all test sections registered;
- every contracted function dispatchable by the exporter;
- generated documentation freshness.

---

# 13. Security and platform assessment

No high-severity production security issue was identified.

Positive controls include:

- pure VBA production implementation;
- no network access from numerical UDFs;
- no external DLL;
- no shell execution from production modules;
- no modal UI;
- read-only workflow permissions;
- isolated temporary workbook;
- same-repository trust check for the self-hosted runner;
- explicit COM cleanup.

The self-hosted runner necessarily enables programmatic access to the VBA project model and lowers macro security for its isolated Excel instance. It should remain dedicated, patched, access-controlled, and unavailable to untrusted fork code.

---

# 14. Findings summary

| ID | Severity | Area | Finding |
|---|---|---|---|
| CR-P1-01 | **P1** | Numerical correctness | Gamma, Chi-square, Beta, and F densities catastrophically cancel for large accepted parameters |
| CR-P2-01 | **P2** | CI assurance | Hosted accuracy CI does not run several critical evaluator, manifest, and holdout tests |
| CR-P2-02 | **P2** | Generated evidence | Fresh accuracy and holdout summaries are not diff-checked against committed artifacts |
| CR-P2-03 | **P2** | Documentation governance | Source-threshold checker misses current duplicated and conflicting measured claims |
| CR-P2-04 | **P2** | Benchmark coverage | Large-shape modal-density coverage is absent; several holdout contracts have thin margins |
| CR-P2-05 | **P2** | Maintainability | Discrete and Test modules are reaching a reviewability boundary |
| CR-P3-01 | **P3** | API completeness | Signed standardization overflow is not mapped to exact CDF/SF limits |
| CR-P3-02 | **P3** | Static assurance | No comprehensive exported-VBA/API/registration static gate was identified |
| CR-P3-03 | **P3** | Performance | No reproducible timing baseline is committed |
| CR-P3-04 | **P3** | Operations | Current branch-protection and live workflow status were not independently visible |

---

# 15. Detailed remediation for CR-P1-01

## 15.1 Root cause

The current density implementations rely on algebraically correct expressions such as:

```text
(shape - 1) log(y) - y - log Gamma(shape)
```

or:

```text
(a - 1) log(x) + (b - 1) log(1-x) - log Beta(a,b)
```

At large shape, each term is enormous. The true log-density near the mode is only order `log(shape)`. The leading terms cancel, exposing absolute error from separately rounded components.

A log-domain implementation prevents power overflow, but it does **not** automatically prevent cancellation among logarithmic terms.

## 15.2 Preferred Gamma/Chi-square branch

Let:

```text
a = Shape
y = X / Scale
delta(a) = StirlingError(a)
D(a,y) = a log(a/y) + y - a
```

Then:

```text
log f =
    -D(a,y)
    + 0.5 log(a)
    - log(y)
    - 0.5 log(2*pi)
    - delta(a)
    - log(Scale)
```

`D(a,y)` must be evaluated with a Loader-style stable deviance routine when `a` and `y` are close.

Chi-square should reuse the same Gamma helper with:

```text
Shape = df / 2
Scale = 2
```

## 15.3 Preferred Beta/F branch

Let:

```text
n = a + b
y = 1 - x

D =
    a log(a / (n x))
  + b log(b / (n y))
```

A stable Stirling/deviance form is:

```text
log f =
    -D
    + 0.5 [log(a) + log(b) - log(n)]
    - log(x)
    - log(y)
    - 0.5 log(2*pi)
    - delta(a)
    - delta(b)
    + delta(n)
```

Again, the deviance terms require `Log1p`/Loader treatment near the mode.

F density can reuse a stable Beta log-density after transforming:

```text
u = r / (1 + r)
v = 1 / (1 + r)
r = x * df1 / df2
```

The Jacobian can be formed as:

```text
log(du/dx) = log(u) + log(v) - log(x)
```

without forming a large raw ratio.

## 15.4 Interim release-safe option

Until the stable branches are implemented:

1. create a dedicated density-envelope study;
2. identify a conservative measured upper shape/df boundary;
3. reject larger density parameters with `#NUM!`;
4. document the boundary as operational, not mathematical.

Do not choose a cutoff from theory alone. The measured error is not monotonic at every decade because it depends on binary cancellation.

## 15.5 Required benchmark expansion

Add main-grid and holdout points around:

```text
1E2
1E4
1E6
1E8
1E10
1E12
1E16
1E20
```

For each family:

### Gamma

```text
X / Scale near Shape
X / Scale moderately below Shape
X / Scale moderately above Shape
```

### Chi-square

```text
X near df
```

### Beta

```text
a = b, x = 0.5
unbalanced large shapes, x near a/(a+b)
```

### F

```text
df1 = df2, x = 1
unbalanced large degrees, x near df2/df1
both log-ratio branches
```

Primary metrics should include:

- absolute log-density error;
- relative density error where representable;
- exact error-code behavior.

---

# 16. Prioritized remediation roadmap

## Release Gate 1 — Fix large-shape densities

1. Add stable shared deviance helpers.
2. Implement stable Gamma/Beta log-density kernels.
3. Reuse them from Chi-square and F.
4. Add VBA regressions.
5. Add main-grid and independent-holdout regimes.
6. Freeze honest thresholds.
7. Re-export observations and regenerate the provenance manifest.

## Release Gate 2 — Make hosted assurance self-testing complete

Add CI execution of:

```text
test_contract_eval.py
test_manifest.py
holdout/test_analyze_holdout.py
holdout/analyze_holdout.py
```

Fail on changed generated holdout artifacts.

## Release Gate 3 — Enforce generated-summary freshness

- generate the main accuracy summary in CI;
- diff it against the committed summary;
- generate and diff benchmark README tables;
- include generated summaries as workflow artifacts.

## Release Gate 4 — Remove duplicated threshold prose

- make `accuracy_contracts.csv` the only threshold authority;
- replace exact source numbers with links to the registry;
- broaden or replace `check_source_thresholds.py`;
- add a regression case for the current Beta-header phrasing.

## Release Gate 5 — Static repository checks

Add a hosted parser for:

- exported VBA module metadata;
- duplicate names;
- public API inventory;
- test registration;
- exporter dispatch;
- generated-document consistency.

## Release Gate 6 — Performance baseline

Publish a nonblocking benchmark with:

- Excel version/build;
- Office bitness;
- CPU;
- warm-up;
- repetition count;
- median/p95;
- regression-suite wall time.

---

# 17. Release-readiness assessment

## Suitable now, within measured domains

- teaching and numerical demonstrations;
- model-validation comparisons;
- controlled quantitative prototypes;
- internal workbook components constrained to the committed benchmark regimes;
- direct-tail calculations;
- discrete probability work within explicit count limits;
- governed use with an exact pinned commit and archived evidence.

## Not yet suitable for an unconditional full-domain claim

Do not claim accurate density evaluation throughout the currently accepted `<1E100` shape/degree domain until CR-P1-01 is corrected or the public density domain is restricted.

## Regulated or high-stakes use

For banking, actuarial, engineering, or other governed use:

- pin the full commit SHA;
- archive the manifest and accuracy summary;
- archive Excel version/build and workflow run IDs;
- validate the actual parameter domain used by the application;
- treat `numerical_limitations.csv` as part of the model specification;
- prohibit large-shape density use until the new branch is validated.

---

# 18. Final verdict

The project has the architecture and assurance foundations of a serious numerical library:

- transparent algorithms;
- reusable special functions;
- direct tails;
- safeguarded inverses;
- explicit error semantics;
- complete discrete coverage;
- consolidated tests;
- source-bound accuracy evidence;
- independent holdouts;
- real Excel CI.

The principal remaining weakness is not structural. It is a specific but severe numerical pattern: **large-shape density cancellation** in four public families. The current validation accepts those inputs, while the formulas can return plausible but dramatically incorrect numbers.

Once that defect is corrected and the hosted assurance workflow is made fully self-testing, the repository would merit a materially higher score.

> **Final score: 8.9 / 10**  
> **Classification: advanced professional numerical library with one release-blocking density defect and several targeted assurance/maintainability improvements remaining.**

---

# Appendix A — Representative current accuracy evidence

| Contract | Main-grid worst | Threshold | Holdout margin where present |
|---|---:|---:|---:|
| `PROB_LogBeta.tiny_unbalanced` | `9.3E-15` absolute | `5E-14` | 4.4x |
| Beta density, tiny unbalanced | `5.65E-15` relative | `5E-14` | 4.5x |
| Beta CDF, tiny unbalanced | `4.05E-14` relative | `5E-13` | 2.8x |
| Beta inverse tiny representable | `4.73E-15` tail residual | `1E-14` | 1.8x |
| F density, tiny unbalanced | `5.56E-15` relative | `5E-14` | 3.5x |
| F inverse tiny representable | `5.57E-14` tail residual | `1E-13` | 1.4x |
| Binomial PMF | `6.62E-13` relative | `1E-12` | 37.3x |
| Poisson PMF | current grid PASS | `1E-14` | 2.0x |
| Hypergeometric PMF | `4.83E-15` relative | `2E-13` | 1.7x |
| Discrete Uniform PMF | `7.92E-17` relative | `5E-15` | 67.3x |
| Normal inverse survival, deep tail | current grid PASS | `5E-16` | 4.0x |
| Lognormal inverse survival, split boundary | `3.84E-10` relative | `1E-9` | 2.0x |
| Gamma density | `1.61E-14` relative on 3 points | `2E-14` | not shown in current holdout summary |
| Weibull standard deviation | `4.12E-15` relative | `5E-15` | not shown in current holdout summary |

These figures are valid committed evidence for the listed points. They are not universal proofs over every accepted parameter.

---

# Appendix B — Independent large-shape density reproduction

## Reproduced current formulas

### Gamma

```text
(shape - 1) log(x/scale)
- x/scale
- log(scale)
- LogGamma(shape)
```

### Chi-square

```text
(df/2 - 1) log(x)
- x/2
- (df/2) log(2)
- LogGamma(df/2)
```

### Beta

```text
(a - 1) log(x)
+ (b - 1) log(1-x)
- LogBeta(a,b)
```

### F

The current positive/negative log-ratio forms were reproduced, including current `LogBeta(df1/2,df2/2)`.

## Reference environment

- IEEE-754 binary64 reproduction of current operation order;
- current repository constants;
- 100-decimal-digit references;
- no desktop-Excel execution.

## Representative log-density differences

| Case | Current log density | Reference log density | Difference |
|---|---:|---:|---:|
| Gamma `a=1E12`, mode | `-14.734375` | `-14.7344490912` | `+7.41E-5` |
| Gamma `a=1E16`, mode | `0` | `-19.3396192772` | `+19.3396` |
| Chi-square `df=2E16`, `x=df` | `-64` | `-20.0327664577` | `-43.9672` |
| Beta `a=b=1E12`, `x=.5` | `13.9409179688` | `13.9362927956` | `+4.63E-3` |
| Beta `a=b=1E16`, `x=.5` | `-186` | `18.5414629816` | `-204.541` |
| F `d1=d2=1E12`, `x=1` | `12.20703125` | `12.2034248442` | `+3.61E-3` |
| F `d1=d2=1E16`, `x=1` | `35` | `16.8085950302` | `+18.1914` |
| F `d1=d2=1E20`, `x=1` | `483328` | `21.4137652162` | `+483306.6` |

The defect appears before the most extreme examples. A dedicated envelope study should determine the precise dispatch point for a stable branch.

---

# Appendix C — Suggested GitHub issues

1. `Implement cancellation-resistant large-shape Gamma and Chi-square densities`
2. `Implement cancellation-resistant large-shape Beta and F densities`
3. `Add large-shape modal-density benchmark and holdout regimes`
4. `Run contract-evaluator and manifest unit tests in Accuracy Gate`
5. `Recompute and diff committed accuracy_summary.md in CI`
6. `Recompute and diff holdout_summary.md in CI`
7. `Replace narrow source-threshold regex with generated contract references`
8. `Remove conflicting Beta density threshold from Continuous module header`
9. `Add exported-VBA static API and test-registration checks`
10. `Publish reproducible Excel/VBA performance baselines`
11. `Evaluate signed standardization-overflow limiting behavior`
12. `Plan Discrete and Test module boundaries before random/array expansion`

---

# Appendix D — Evidence confidence

| Conclusion | Confidence |
|---|---|
| Repository metrics and API counts | High |
| Architecture and dependency assessment | High |
| Main-grid and holdout verdicts | High as committed, source-bound evidence |
| Large-shape density defect | High; formula-level binary64 reproduction shows enormous discrepancies |
| Exact desktop-Excel values for the new density cases | Medium-high; must be confirmed in VBA |
| Test-harness design | High |
| Hosted CI definition | High |
| Live self-hosted-runner availability | Not independently established |
| Branch-protection enforcement | Not independently established |
| Performance characteristics | Medium; algorithmic review only, no timing run |
