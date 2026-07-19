# Independent Code Review — VBA Probability Distributions

> **Repository:** `danielep71/VBA-PROBABILITY-DISTRIBUTIONS`  
> **Branch:** `main`  
> **Commit reviewed:** [`dae05c770d9ef1cc1e3e4ae70175112a5ce3d969`](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/commit/dae05c770d9ef1cc1e3e4ae70175112a5ce3d969)  
> **Review date:** 19 July 2026  
> **Reviewer:** OpenAI GPT-5.6 Thinking  
> **Recommended repository path:** `docs/INDEPENDENT_CODE_REVIEW_2026-07-19.md`

---

## 1. Executive assessment

### Overall score: **9.0 / 10**

### Production numerical library: **9.0 / 10**

### Numerical-assurance framework: **9.0 / 10**

This repository is a serious numerical-computing project, not a collection of worksheet wrappers. It has a coherent layered architecture, native special-function kernels, direct tail APIs, guarded arithmetic, explicit worksheet-error contracts, a substantial deterministic VBA test suite, and a high-precision external benchmark system with regime-specific contracts.

The strongest engineering characteristics are:

- a clean separation between elementary numerics, special functions, distribution families, tests, and external accuracy evidence;
- direct survival and inverse-survival calculations where subtraction from one would destroy tail information;
- explicit distinction between valid underflow, predictable numerical failure, and unexpected runtime failure;
- Boolean Try-contracts for iterative kernels, with no return of unconverged partial sums;
- stable `Log1p`, `Expm1`, log-density, paired-tail, and safeguarded inverse formulations;
- a measured correction to unbalanced `LogBeta` using a stable Lanczos log-gamma difference;
- 73 explicit accuracy contracts, including separate balanced and unbalanced Beta regimes and separate quantile and forward-tail residual criteria for inverse functions;
- committed VBA observations reconstructed through a two-part `hi;lo` representation and analyzed with `Decimal`;
- a fresh 134-point holdout used to validate and freeze the new regime-specific contracts;
- unusually complete source documentation and numerical provenance.

The repository nevertheless has three material issues that should block an unconditional release-quality claim:

1. **Lognormal variance and standard deviation can return zero or `#NUM!` even when the mathematically correct result is finite and representable.** The formulas still evaluate large and small exponential factors separately instead of reconstructing the final moment in one logarithmic expression.

2. **F functions accept parameters outside the empirically validated incomplete-beta range and can return inaccurate numeric values rather than a clean non-convergence error.** Committed evidence reports errors up to approximately `4E-7`, while several source comments state that such cases return a clean failure rather than a wrong answer.

3. **The Python release gate can exit successfully when an active tail-residual contract cannot be evaluated.** A failed import of the high-precision incomplete-beta helper downgrades active contracts to non-blocking “CHARACTERIZATION ONLY.”

These issues do not invalidate the architecture or the broad benchmark evidence. They do, however, matter for a library positioned as a governed numerical component.

### Independent verdict

> **Advanced, credible, and unusually well engineered for native VBA. Suitable for controlled quantitative, teaching, and model-validation use within its measured domains. Three release-gate issues should be corrected before presenting the current commit as fully hardened across all documented parameter ranges.**

---

## 2. Review scope and evidence boundary

The review examined the exact files committed at the stated SHA through GitHub, including:

- [`src/M_STATS_PROBDIST_CORE.bas`](../src/M_STATS_PROBDIST_CORE.bas)
- [`src/M_STATS_PROBDIST_SPECIALFUNCS.bas`](../src/M_STATS_PROBDIST_SPECIALFUNCS.bas)
- [`src/M_STATS_PROBDIST_NORMALFAMILY.bas`](../src/M_STATS_PROBDIST_NORMALFAMILY.bas)
- [`src/M_STATS_PROBDIST_TFAMILY.bas`](../src/M_STATS_PROBDIST_TFAMILY.bas)
- [`src/M_STATS_PROBDIST_CONTINUOUS.bas`](../src/M_STATS_PROBDIST_CONTINUOUS.bas)
- [`tests/M_STATS_PROBDIST_TEST.bas`](../tests/M_STATS_PROBDIST_TEST.bas)
- [`ci/Run-ExcelVbaTests.ps1`](../ci/Run-ExcelVbaTests.ps1)
- [`.github/workflows/excel-vba-regression.yml`](../.github/workflows/excel-vba-regression.yml)
- [`benchmark/accuracy_contracts.csv`](../benchmark/accuracy_contracts.csv)
- [`benchmark/compute_errors.py`](../benchmark/compute_errors.py)
- [`benchmark/probability_accuracy_grid.csv`](../benchmark/probability_accuracy_grid.csv)
- [`benchmark/accuracy_summary.md`](../benchmark/accuracy_summary.md)
- [`benchmark/delta_seam_study/`](../benchmark/delta_seam_study/)
- [`benchmark/beta_f_unbalanced/`](../benchmark/beta_f_unbalanced/)
- [`benchmark/holdout/`](../benchmark/holdout/)
- [`benchmark/numerical_limitations.csv`](../benchmark/numerical_limitations.csv)
- [`README.md`](../README.md)
- repository policy and configuration files.

### Execution boundary

This is a static source and committed-evidence audit. The reviewer did **not** execute desktop Excel or the VBA regression suite. The committed grids contain populated VBA observations, and the generated accuracy summary reports all active contracts passing, but the current commit exposes no attached GitHub status checks or workflow-run evidence through the repository API.

Accordingly:

- implementation findings are based on the committed source;
- accuracy findings are based on the committed observations, contracts, and analyzers;
- no claim is made that the current commit was independently compiled or executed by the reviewer;
- CI design is assessed separately from CI operational evidence.

---

## 3. Repository metrics

### 3.1 Source size

| Component | Lines |
|---|---:|
| `M_STATS_PROBDIST_CORE.bas` | 875 |
| `M_STATS_PROBDIST_SPECIALFUNCS.bas` | 1,593 |
| `M_STATS_PROBDIST_NORMALFAMILY.bas` | 3,653 |
| `M_STATS_PROBDIST_TFAMILY.bas` | 3,298 |
| `M_STATS_PROBDIST_CONTINUOUS.bas` | 4,939 |
| **Production VBA total** | **14,358** |
| `M_STATS_PROBDIST_TEST.bas` | 4,881 |
| **Production plus primary tests** | **19,239** |
| `ci/Run-ExcelVbaTests.ps1` | 259 |
| GitHub Actions workflow | 64 |
| Main README | 959 |

### 3.2 Public numerical surface

The module headers expose:

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
| **Total worksheet-facing UDFs** | **64** |

The internal layers additionally expose project-scoped constants, predicates, guarded arithmetic, log-gamma/log-beta functions, regularized incomplete-beta and incomplete-gamma functions, and inverse kernels.

### 3.3 Numerical-assurance evidence

| Artifact | Size / coverage |
|---|---:|
| Regime-aware accuracy contracts | 73 |
| Main accuracy-grid observations | 685 |
| Generated contract verdicts | 73 PASS |
| Delta seam study | 396 data rows |
| Unbalanced Beta/F forward study | 64 data rows |
| Beta/F inverse study | 52 data rows |
| Independent holdout | 134 fresh points |
| Recorded numerical limitations | 1 |

The generated summary reports:

```text
FAIL: 0
KNOWN LIMITATION in the contract table: 0
CHARACTERIZATION ONLY: 0
PENDING: 0
```

A separate numerical-limitations register records the extreme incomplete-beta/F domain rather than treating it as a contract that the implementation cannot meet.

---

## 4. Scoring methodology

The score is a weighted engineering assessment. A score of 10 requires not merely good algorithms, but internally consistent contracts, reproducible evidence, active release gates, and no known wrong-result path inside the documented public domain.

| Area | Weight | Score | Weighted contribution |
|---|---:|---:|---:|
| Functional correctness | 18% | 8.7 | 1.566 |
| Numerical robustness | 17% | 9.0 | 1.530 |
| Architecture and modularity | 11% | 9.6 | 1.056 |
| Public API design | 8% | 9.3 | 0.744 |
| Error handling and diagnostics | 7% | 9.0 | 0.630 |
| Regression testing | 11% | 9.4 | 1.034 |
| Accuracy benchmarking | 10% | 9.0 | 0.900 |
| CI and release engineering | 7% | 8.3 | 0.581 |
| Documentation | 5% | 8.5 | 0.425 |
| Maintainability and repository hygiene | 5% | 8.8 | 0.440 |
| Performance engineering | 1% | 8.7 | 0.087 |
| **Total** | **100%** |  | **8.993 / 10** |

Rounded overall score:

```text
9.0 / 10
```

### Score interpretation

| Score | Interpretation |
|---:|---|
| 9.5–10.0 | Exceptional, independently executable and release-certified |
| 9.0–9.4 | Advanced professional numerical library with limited material gaps |
| 8.0–8.9 | Strong implementation requiring further hardening |
| 7.0–7.9 | Good foundation with substantial correctness or governance work |
| Below 7.0 | Significant design, correctness, or assurance deficiencies |

---

## 5. Component scores

| Component | Score | Assessment |
|---|---:|---|
| `M_STATS_PROBDIST_CORE` | **9.3** | Excellent shared primitive layer; conservative domain naming and signed limiting behavior can improve. |
| `M_STATS_PROBDIST_SPECIALFUNCS` | **9.1** | Sophisticated kernels and a strong LogBeta correction; incomplete-beta range and one precondition mismatch remain. |
| `M_STATS_PROBDIST_NORMALFAMILY` | **8.8** | Very strong tail design and densities; lognormal moment reconstruction contains a confirmed finite-result defect. |
| `M_STATS_PROBDIST_TFAMILY` | **9.0** | Stable Student t, Chi-square, and F transformations; extreme F domain is not safely enforced. |
| `M_STATS_PROBDIST_CONTINUOUS` | **9.1** | Broad, consistent public API with guarded arithmetic and full-range Uniform formulas. |
| `M_STATS_PROBDIST_TEST` | **9.4** | Exceptional deterministic regression harness for VBA; several newly identified cases are not yet covered. |
| External accuracy framework | **9.0** | Regime-aware, Decimal-based and holdout-validated; release-gate fallback and oracle independence need tightening. |
| Excel/PowerShell CI | **8.3** | Strong design, but no current-run evidence and no hosted benchmark/static gate. |
| Documentation | **8.5** | Rich and transparent, but recent numerical changes have left several contradictions. |

---

## 6. Architecture review

### 6.1 Layering

The production architecture is clean:

```text
M_STATS_PROBDIST_CORE
        ↓
M_STATS_PROBDIST_SPECIALFUNCS
        ↓
M_STATS_PROBDIST_NORMALFAMILY
M_STATS_PROBDIST_TFAMILY
M_STATS_PROBDIST_CONTINUOUS
        ↓
M_STATS_PROBDIST_TEST
```

The separation is meaningful rather than cosmetic:

- **Core** owns finite-value predicates, magnitude policy, stable elementary functions, guarded arithmetic, the raw inverse-normal seed, and diagnostics.
- **Special Functions** owns reusable gamma/beta kernels and their inverses.
- **Distribution modules** own public validation, parameterization, support behavior, tail orientation, worksheet errors, and diagnostics.
- **Tests** own the assertion framework, suite order, counters, regression registry, and consolidated verdict.
- **Benchmark tooling** owns high-precision references and measured accuracy contracts independently of worksheet-facing source comments.

This is the correct architecture for a growing numerical library.

### 6.2 Visibility boundaries

`Option Private Module` is used for the Core and Special Functions layers. This preserves project-level reuse while hiding low-level names from worksheet formula autocomplete.

That is a strong choice. It avoids forcing every internal primitive into `Private` duplication while keeping the public worksheet surface intentional.

### 6.3 Dependency direction

The dependency direction is consistently downward. No distribution module appears to own a duplicate incomplete-beta or incomplete-gamma implementation. Shared numerical behavior is centralized.

This materially lowers the risk of family-specific divergence.

### 6.4 Architecture score rationale

**9.6 / 10**

The remaining deduction is primarily maintainability:

- Lanczos coefficients are duplicated inside `PROB_LogGamma` and `PROB_LogGammaDelta`;
- several source headers no longer describe the actual current internal surface;
- reference-code helpers are duplicated across benchmark folders.

---

# 7. Production-code review

## 7.1 `M_STATS_PROBDIST_CORE`

### Strengths

#### True finiteness is separated from policy

The distinction between:

```vb
PROB_IsFinite
```

and:

```vb
PROB_IsWithinSupportedMagnitude
```

is conceptually correct.

A finite IEEE-754 value is not automatically inside the project's validated or representational contract. Conversely, the `1E100` policy is documented as a coarse guard rather than a mathematical boundary.

#### Guarded arithmetic is reusable

The module provides:

- `PROB_TryExp`
- `PROB_TryAdd`
- `PROB_TryMultiply`
- `PROB_TryDivide`
- `PROB_TryAffineTransform`
- `PROB_TryStandardize`

This prevents predictable overflow from being misclassified as an unexpected VBA runtime failure.

#### Underflow policy is explicit

`PROB_TryExp` treats negative exponential underflow as a valid zero and positive overflow as failure. This is the correct default for densities and tail probabilities.

#### `Log1p` and `Expm1` are implemented deliberately

The compensated formulations materially improve:

- small Exponential and Weibull CDFs;
- small parameter transformations;
- tail and interval calculations;
- stable log-gamma differences.

#### Diagnostic side effects are controlled

`PROB_SetStatus` always writes the ByRef diagnostic but gates `Application.StatusBar` behind a project constant. This is appropriate for worksheet UDFs, where object-model side effects are unreliable and expensive.

### Improvements

#### Signed standardization overflow is discarded

`PROB_TryStandardize` returns only Boolean success/failure. It does not preserve whether the mathematical standardized value tends to positive or negative infinity.

For CDF and survival functions, the sign gives an exact limiting value:

| Function | Positive overflow | Negative overflow |
|---|---:|---:|
| Normal density | 0 | 0 |
| Normal CDF | 1 | 0 |
| Normal survival | 0 | 1 |
| Z-score | `#NUM!` | `#NUM!` |

A richer outcome type could improve mathematical completeness.

#### Naming still says “supported magnitude”

The constant has been correctly reframed as a parameter-magnitude guard, but predicates retain names such as:

```vb
PROB_IsWithinSupportedMagnitude
```

This is not a correctness issue, but it perpetuates ambiguity between:

- representational guard;
- convergence range;
- validated accuracy range.

### Core verdict

A high-quality primitive layer with explicit floating-point policy. No major structural change is required.

---

## 7.2 `M_STATS_PROBDIST_SPECIALFUNCS`

### Strengths

#### Iterative routines use a real Try-contract

The incomplete-beta and incomplete-gamma routines return Boolean success and leave output non-contractual on failure. They do not silently return partial sums after hitting iteration caps.

This is one of the most important design decisions in the repository.

#### Paired incomplete-beta arguments are accepted

`PROB_TryBetaRegularized` receives both `X` and `Y`, allowing callers to construct complementary quantities without cancellation.

This is particularly valuable for:

- Student t near the center;
- F ratios at extreme scales;
- inverse beta roots and their complements;
- direct tail orientation.

#### Lower and upper incomplete-gamma functions are separate

`P(a,x)` and `Q(a,x)` are exposed separately, so upper tails do not depend on subtracting a nearly-one lower tail from one.

#### Inverse kernels are safeguarded

The inverse beta and gamma functions combine:

- analytical or asymptotic seeds;
- Newton steps;
- brackets;
- bisection fallback;
- explicit iteration caps;
- convergence checks.

#### Stable `PROB_LogGammaHalfDiff`

The half-step gamma ratio avoids cancellation in Student t normalization and `Beta(z, 1/2)`.

#### Stable `PROB_LogGammaDelta`

The new kernel computes:

```text
LogGamma(z+s) - LogGamma(z)
```

without forming two large log-gamma values.

Its implementation uses:

```text
s Log(T)
+ (z+s-1/2) Log1p(s/T)
- s
+ Log1p((A(z+s)-A(z))/A(z))
```

and constructs the Lanczos-series difference directly.

This is a strong, elegant correction to unbalanced `LogBeta`.

#### Regime-aware `PROB_LogBeta`

The function now dispatches among:

- half-integer shortcuts;
- stable log-gamma delta for unbalanced arguments;
- direct identity for balanced arguments.

The external contract correctly judges LogBeta primarily by **absolute error**, because downstream calculations exponentiate `-LogBeta`.

### Improvements

#### Public-surface documentation omits `PROB_LogGammaDelta`

The module header lists:

- `PROB_LogGamma`
- `PROB_LogGammaHalfDiff`
- `PROB_LogBeta`

but not the new project-scoped delta kernel.

#### The crossover is still described as provisional

The production constant comment states:

```vb
' provisional; confirm from the VBA seam study
```

and the LogBeta procedure says the exact constant remains to be confirmed.

That no longer matches the repository evidence:

- the seam study is committed;
- the main contract is validated and frozen;
- the holdout stresses both sides of the `0.1` seam.

The source comments should state the actual validation evidence rather than a future action.

#### `PROB_LogGammaDelta` precondition does not match every caller

The documented precondition is:

```text
LargeArg >= 1
```

However `PROB_LogBeta` can call it whenever:

```text
SmallArg / LargeArg < 0.1
```

including cases where both arguments are below one, for example:

```text
A = 0.20
B = 0.01
```

Then:

```text
LargeArg = 0.20
Increment = 0.01
```

which violates the declared precondition.

The formula may still be accurate in this region, but the committed seam and holdout grids primarily validate `LargeArg >= 1`. This is a contract gap.

Recommended resolution:

1. add a dedicated both-subunit unbalanced grid;
2. validate the real VBA implementation;
3. either broaden the precondition to `LargeArg > 0` or keep such cases on a reflection-aware/direct route.

#### Lanczos constants are duplicated

`PROB_LogGamma` and `PROB_LogGammaDelta` each contain their own copy of:

```text
g, P0, ..., P8
```

The delta comment explicitly says they must match.

Move them to module-level constants or a shared private evaluation helper. Numerical constants that must remain identical should have one source of truth.

### Special-functions verdict

The core algorithms are sophisticated and the LogBeta repair is a genuine improvement. The remaining issues are contract alignment and maintenance, not a need for wholesale numerical redesign.

---

## 7.3 `M_STATS_PROBDIST_NORMALFAMILY`

### Strengths

#### Complete tail-aware API

The module exposes:

- density;
- CDF;
- survival;
- inverse CDF;
- inverse survival;
- interval probability;
- Z-score;
- fast inverse helper;
- Lognormal moments and parameter conversion.

Direct inverse-survival functions are especially valuable because `1-q` loses tiny upper-tail probabilities.

#### Tail-specific Standard Normal algorithms

The CDF/survival design avoids reconstructing a tiny tail through subtraction. The inverse uses an Acklam seed and a refinement step, with care around saturated tails.

#### Stable interval probabilities

Same-tail interval probabilities are treated as their own numerical problem rather than always subtracting two CDFs.

#### Normal and Lognormal densities use the log domain

The general Normal density is reconstructed from:

```text
-0.5 z² - log(σ) - 0.5 log(2π)
```

The Lognormal density uses:

```text
-0.5 z² - log(x) - log(σlog) - 0.5 log(2π)
```

This prevents independent numerator and denominator underflow.

#### Error behavior is consistent

Public functions:

- return `Variant`;
- map predictable domain/overflow failure to `#NUM!`;
- map unexpected runtime failure to `#VALUE!`;
- populate optional diagnostics;
- raise no modal UI.

### Material defect: Lognormal moments split cancelling exponentials

The Lognormal variance code evaluates:

```vb
Factor = PROB_Expm1(VarianceLog)
ExpShift = Exp(2 * MeanLog + VarianceLog)
Variance = Factor * ExpShift
```

The standard deviation similarly evaluates:

```vb
Factor = Sqr(PROB_Expm1(VarianceLog))
ExpShift = Exp(MeanLog + 0.5 * VarianceLog)
StdDev = Factor * ExpShift
```

These are algebraically correct but not generally floating-point safe. The factors can overflow or underflow independently even when the final product is finite.

#### Demonstrable finite-result failure

Let:

```text
MeanLog = -750
StdDevLog² = 700
```

The true variance is:

```text
(expm1(700)) × exp(-800)
≈ exp(-100)
≈ 3.720075976020836E-44
```

This is representable.

The current code evaluates `ExpShift = Exp(-800)`, which underflows to zero, and returns variance zero.

A second case:

```text
MeanLog = -800
StdDevLog² = 800
```

has:

```text
Variance ≈ 1
StdDev ≈ 1
```

but the current code first tries to evaluate `Exp(800)` as a guard for `Expm1(800)`. That overflows and returns `#NUM!`.

#### Correct architecture

Compute the final moments in one log expression.

Define a stable positive helper:

```text
LogExpm1(v) = log(exp(v)-1)
```

using:

```text
small/moderate v:
    log(PROB_Expm1(v))

large v:
    v + PROB_Log1p(-exp(-v))
```

Then:

```text
logVariance =
    LogExpm1(v) + 2*MeanLog + v

logStdDev =
    0.5*LogExpm1(v) + MeanLog + 0.5*v
```

and call `PROB_TryExp` exactly once on the final logarithm.

This approach:

- preserves finite results produced by cancellation;
- returns zero only when the final moment genuinely underflows;
- returns `#NUM!` only when the final moment genuinely overflows;
- removes the need for separate product overflow tests.

#### Required regressions

Add at least:

```text
Variance(-750, sqrt(700)) ≈ exp(-100)
Variance(-800, sqrt(800)) ≈ 1
StdDev(-800, sqrt(800)) ≈ 1
```

plus nearby values on both sides of the overflow/underflow boundaries.

### Normal-family verdict

The module is otherwise excellent, but this is a real production correctness defect in a public moment API and is the largest deduction in the review.

---

## 7.4 `M_STATS_PROBDIST_TFAMILY`

### Strengths

#### Student t density

Uses a logarithmic form and `PROB_LogGammaHalfDiff`, avoiding cancellation at large degrees of freedom.

#### Student t CDF and survival

Uses:

- exact low-degree formulas;
- incomplete-beta transformations;
- paired complementary arguments;
- direct tail orientation;
- local central-mass handling.

#### Student t inverse

Uses:

- exact branches;
- beta inversion where appropriate;
- Cornish-Fisher seed;
- safeguarded Newton iteration;
- bisection;
- no artificial quantile ceiling.

#### Chi-square family

Uses direct regularized incomplete-gamma `P` and `Q`, with guarded quantile scaling.

#### F family

Builds the beta logistic pair through log ratios rather than forming:

```text
x * df1 / df2
```

directly. Quantiles are reconstructed in the logarithmic domain from both beta-root components.

### Material domain-contract problem: extreme F parameters

The source documentation states that incomplete-beta parameters beyond the approximate convergence range are:

- accepted;
- attempted;
- expected to exhaust the iteration budget;
- returned as clean non-convergence errors;
- never returned as wrong partial answers.

The committed numerical-limitations register instead reports:

```text
at least one beta shape parameter above approximately 1E7
errors up to approximately 4E-7
```

This means the continued fraction can satisfy its local increment criterion while the returned value is outside the published F accuracy contract.

The problem is not the LogBeta normalization. It is a distinct incomplete-beta convergence/accuracy limitation.

### Why this matters

A clean `#NUM!` is auditable. A plausible numeric result with an unmarked error of `4E-7` is much harder to detect.

The public F API currently validates against the coarse `1E100` parameter guard, not the empirically validated incomplete-beta range.

### Recommended policy

The repository should choose and state one of two policies.

#### Preferred for a governed numerical library

Use a strict validated-domain policy for public F UDFs:

- reject parameters outside a conservatively measured safe envelope;
- return `#NUM!`;
- explain that the boundary is operational, not mathematical;
- offer a lower-level best-effort kernel only to advanced VBA callers if needed.

#### Alternative

Retain best-effort evaluation, but explicitly state:

- values outside the validated range may be inaccurate without non-convergence;
- successful return does not imply contract-level accuracy;
- the optional `Status` channel should report that the input is outside the validated range.

The current wording—“returns a clean non-convergence error, never a wrong answer”—is contradicted by the committed evidence and should be removed immediately.

### T-family verdict

Excellent numerical transformations inside the validated range. Domain enforcement and documentation need to match the measured behavior.

---

## 7.5 `M_STATS_PROBDIST_CONTINUOUS`

### Strengths

#### Gamma family

- density evaluated in log space;
- support behavior handled explicitly;
- CDF and survival use direct incomplete-gamma functions;
- ratio overflow saturates to mathematical limits;
- moments use guarded multiplication;
- standard deviation avoids square-rooting an overflowing variance.

#### Beta family

- public validation is consistent;
- density uses LogBeta;
- CDF and survival use paired incomplete-beta arguments;
- inverse uses the shared safeguarded beta inverse;
- balanced and unbalanced behavior is measured separately.

#### Exponential family

Uses `Expm1` and guarded products/inverses, preserving very small left-tail probabilities.

#### Weibull family

- logarithmic/guarded reconstruction;
- direct survival;
- stable left tail;
- large-shape moment treatment;
- explicit tiny-shape failure classification.

#### Uniform family

The implementation is deliberately designed for the full finite Double range, including opposite-signed bounds whose naive width would overflow.

This is an unusually thoughtful feature for a VBA library.

### Improvements

#### Public source comments do not expose regime-specific Beta accuracy

The benchmark contracts correctly distinguish balanced and unbalanced Beta behavior, but the public distribution module does not give users a concise local statement of those regimes.

A short module-level accuracy section should link to the contract table and explain that:

- balanced Beta retains the tight contract;
- unbalanced density/CDF/survival/inverse use separate measured thresholds;
- LogBeta catastrophic cancellation has been removed;
- residual accuracy depends on the downstream incomplete-beta calculation.

### Continuous-family verdict

A broad and well-engineered module. No major defect was found beyond inherited special-function limitations and documentation alignment.

---

# 8. Regression-test review

## 8.1 Strengths

The 4,881-line consolidated test module is exceptional for VBA.

It includes:

- one authoritative suite order;
- shared counters;
- shared assertion helpers;
- known values;
- exact support cases;
- complement identities;
- symmetry identities;
- inverse round-trips;
- deep tails;
- full-range Uniform cases;
- guarded-overflow regressions;
- exact `#NUM!` versus `#VALUE!` checks;
- named historical regressions;
- a CI-readable failure buffer.

Assertion failures carry actual, expected, error, and tolerance text in the assertion label. The CI runner therefore receives more than a bare test name.

## 8.2 Tolerance policy

The primary VBA suite uses tolerances such as:

```text
1E-10 absolute
1E-10 relative
1E-9 tail
1E-6 loose
```

These are intentionally broader than several external benchmark contracts.

That is acceptable if the roles remain distinct:

- VBA suite: fast deterministic regression and public-contract smoke test;
- external benchmark: measured high-precision accuracy gate.

The README and contribution guide should keep this distinction explicit.

## 8.3 Missing regressions

Add permanent tests for:

1. Lognormal variance cancellation:
   ```text
   MeanLog=-750, VarianceLog=700
   ```
2. Lognormal variance/stddev with individually overflowing factors but finite final moments:
   ```text
   MeanLog=-800, VarianceLog=800
   ```
3. `PROB_LogGammaDelta` where:
   ```text
   0 < LargeArg < 1
   ```
4. F inputs just below, near, and above the operationally validated incomplete-beta range.
5. Strict release-gate behavior when the external reference helper is unavailable.

## 8.4 Test score rationale

**9.4 / 10**

The test architecture is excellent. The deduction reflects untested newly identified boundaries and the absence of independent current-run evidence.

---

# 9. External accuracy and benchmark review

## 9.1 Strengths

### Regime-aware contracts

The contract schema is well designed:

```csv
contract_id,function,regime,measure,metric,threshold,domain,provenance,status,evidence,notes
```

It avoids fake function names while allowing one real function to carry multiple numerical contracts.

### Measures are function-appropriate

The framework distinguishes:

- output relative/absolute error;
- LogBeta absolute error;
- inverse quantile error;
- inverse forward-tail residual.

That is substantially better than applying one relative-error metric to every numerical object.

### Two-part observed values

VBA writes:

```text
hi;lo
```

and Python reconstructs with `Decimal`, avoiding an artificial one-literal 15-digit floor.

### Main benchmark coverage

The main CSV contains 685 observed rows and feeds 73 contracts.

### Holdout validation

A 134-point fresh holdout includes:

- new non-integer unbalanced Beta shapes;
- fresh F degree combinations;
- new central and tail probabilities;
- near-seam LogBeta ratios on both sides of `0.1`;
- between-decade ratios.

The regime-specific contracts were frozen only after holdout confirmation.

### Honest numerical-limitations register

The extreme incomplete-beta/F issue is recorded separately rather than hidden inside a passing contract.

## 9.2 Release-gate defect: active contracts can become non-blocking characterization

`compute_errors.py` imports the high-precision incomplete-beta helper inside:

```python
try:
    ...
except Exception:
    _HAVE_IBETA = False
```

For an active `tail_probability_residual` contract, if the import fails, the analyzer emits:

```text
CHARACTERIZATION ONLY
```

The strict gate does not include `n_char` in its blocking count.

Therefore:

> deleting, corrupting, or making `_ibeta.py` incompatible can cause active frozen tail-residual contracts to stop being evaluated while the strict release gate still exits successfully.

### Required fix

- catch only expected import exceptions;
- retain and print the exact import failure;
- treat an unevaluable **active** contract as `PENDING` or `ERROR`;
- return non-zero;
- reserve `CHARACTERIZATION ONLY` for contracts explicitly marked that way in the contract file;
- optionally assert that every active contract has at least one matching observation and a registered evaluator.

Suggested logic:

```python
if status == "active" and evaluator_unavailable:
    n_pending += 1
    verdict = "PENDING — evaluator unavailable"
```

The strict gate should never pass with an active contract unevaluated.

## 9.3 Holdout analyzer is no longer rerunnable after freezing

`benchmark/holdout/analyze_holdout.py` filters contracts with:

```python
if c["provenance"] != "measured provisional":
    continue
```

All relevant contracts are now:

```text
validated and frozen
```

A current rerun therefore tests zero contracts and falls into the generic failure message.

This undermines reproducibility of the evidence used to justify freezing.

### Required fix

Select contracts by:

- evidence set;
- explicit contract IDs;
- regime/function presence in the holdout;
- or both provisional and frozen provenance.

Also commit a generated:

```text
benchmark/holdout/holdout_summary.md
```

showing every tested contract, holdout worst error, threshold, margin, and verdict.

## 9.4 Reference-oracle independence

The extreme Beta/F studies use a high-precision `_ibeta.py` continued-fraction implementation that is structurally similar to the production algorithm.

Using 50-digit arithmetic greatly reduces floating-point error and is useful. It is not, however, fully independent of algorithmic arrangement.

For frozen high-stakes contracts, selected points should also be cross-checked against an independent oracle, such as:

- SciPy where its parameter range remains reliable;
- R `pbeta` / `qbeta`;
- Boost.Math;
- numerical quadrature with high precision;
- a second asymptotic or saddlepoint formulation.

The benchmark README states that references were cross-checked against SciPy for every function. That assertion should be tied to a generated cross-check artifact, especially for the new extreme and inverse grids.

## 9.5 Benchmark score rationale

**9.0 / 10**

The design is advanced and evidence-driven. The release-gate fallback and holdout reproducibility bug prevent a higher score.

---

# 10. CI and release-engineering review

## 10.1 Excel/VBA workflow strengths

The GitHub Actions workflow:

- targets a self-hosted Windows x64 Excel runner;
- skips untrusted fork pull requests;
- uses read-only repository permissions;
- has a timeout;
- checks out exact source;
- uploads a test artifact.

The PowerShell runner:

- creates an isolated macro-enabled workbook;
- imports production modules in dependency order;
- imports the test module;
- injects a CI-only bridge without widening production visibility;
- runs the complete suite through Excel COM;
- validates total/pass/fail consistency;
- retrieves detailed failed-assertion text;
- records Excel version and build;
- returns non-zero on failure;
- performs explicit COM cleanup.

This is strong engineering.

## 10.2 Operational evidence gap

The reviewed commit exposes:

- no combined status checks;
- no associated workflow-run evidence through the repository API.

The workflow may be correctly configured but unavailable because the self-hosted runner is offline or not yet registered. A workflow file is not equivalent to an enforced release gate.

## 10.3 Benchmark changes are not CI-gated

The Excel workflow path filter covers:

```text
src/**
tests/**
ci/**
workflow file
```

It does not cover:

```text
benchmark/**
```

There is also no standard hosted Python workflow for:

- `compute_errors.py`;
- contract-schema validation;
- grid completeness;
- generated README table drift;
- Python syntax;
- holdout analyzer;
- forbidden generated files;
- numerical-limitations schema.

### Recommended CI split

#### Hosted Python assurance job

Run on ordinary GitHub-hosted Windows or Ubuntu:

```text
python -m py_compile benchmark/**/*.py
python benchmark/compute_errors.py
python benchmark/render_contract_table.py --check
python benchmark/holdout/analyze_holdout.py
verify no active contract is pending
verify no generated table drift
verify no __pycache__ or *.pyc committed
```

#### Self-hosted Excel job

Keep the existing VBA regression workflow for source/test changes.

#### Branch protection

Require both checks on `main`.

## 10.4 CI score rationale

**8.3 / 10**

The design is strong, but operational evidence and benchmark automation are incomplete.

---

# 11. Documentation review

## 11.1 Strengths

Documentation is a major project asset.

The repository includes:

- a polished README;
- module and procedure contracts;
- numerical provenance;
- worksheet equivalents;
- public error policy;
- supported-domain explanation;
- benchmark methodology;
- accuracy contracts;
- holdout design;
- numerical-limitations register;
- contribution, security, and conduct policies;
- Wiki links.

The source comments are far more complete than typical VBA numerical code.

## 11.2 Current contradictions

### Benchmark README retains obsolete floor language

The benchmark README still says:

- Standard Normal survival and Lognormal density have `5E-15` claims;
- misses are reported as “below harness precision” rather than failure.

The current contract table correctly uses:

```text
NormalStandard_Survival: 2E-14
Lognormal_Density: 3E-14
```

and the generic floor has been removed.

### Test-statistic accuracy table is stale

The T-family source states:

```text
F quantile <= 5.9E-13
```

The validated regime-aware contract is:

```text
F inverse quantile <= 2E-10
F inverse tail residual <= 2E-10
```

The source should not publish a tighter global value than the frozen contract.

### Special-functions comments are stale

- module public surface omits `PROB_LogGammaDelta`;
- crossover remains described as provisional;
- exact validation is described as future work;
- module update date predates the new kernel and benchmark program.

### Main README understates current automation

The validation boundary says:

```text
regression execution is manual inside Excel/VBA
```

The roadmap says:

```text
investigation of Windows/Excel regression execution in CI
```

Yet a self-hosted Excel COM workflow and runner already exist.

The accurate wording is:

> An Excel COM workflow is implemented; current runner availability and required-check enforcement must be verified.

### Extreme F limitation is not prominent enough

The benchmark README records the limitation, but the public README and T-family source still imply clean non-convergence outside the approximate kernel range.

A user should not have to inspect a benchmark CSV to discover a wrong-result risk.

## 11.3 Documentation score rationale

**8.5 / 10**

The depth and quality are excellent. The deduction is for recent-contract drift and one materially misleading domain statement.

---

# 12. Maintainability and repository hygiene

## 12.1 Strengths

- consistent naming;
- `Option Explicit`;
- disciplined procedure structure;
- shared validation helpers;
- no hidden worksheet-function dependency;
- no modal UI from numerical UDFs;
- one consolidated test harness;
- one machine-readable accuracy-contract source.

## 12.2 Issues

### Committed Python bytecode

The repository contains:

```text
benchmark/beta_f_unbalanced/__pycache__/_ibeta.cpython-311.pyc
benchmark/beta_f_unbalanced/__pycache__/_ibeta.cpython-313.pyc
```

Binary interpreter caches should not be versioned.

### `.gitignore` lacks Python cache rules

Add:

```gitignore
__pycache__/
*.py[cod]
.pytest_cache/
.mypy_cache/
```

### Duplicate `_ibeta.py`

High-precision incomplete-beta helper code exists in more than one benchmark folder.

Move it to a shared location such as:

```text
benchmark/reference/incomplete_beta.py
```

and import it consistently.

### Duplicated Lanczos coefficients

Move `g` and `P0..P8` to one module-level source.

### Stale `UPDATED` metadata

Several materially changed procedures still report `2026-07-11`.

If update dates are part of the house style, they need automated consistency checks or should be removed. Manually maintained dates drift quickly.

---

# 13. Security and platform considerations

No high-severity security defect was identified in the numerical source.

Positive controls include:

- pure VBA with no external DLL;
- no shell execution from production UDFs;
- no network access;
- no modal UI in numerical functions;
- read-only workflow permissions;
- fork-PR exclusion on the self-hosted Excel runner;
- isolated temporary workbook;
- COM cleanup.

The self-hosted runner necessarily enables programmatic access to the VBA project model and lowers automation security for its isolated Excel instance. That runner should remain dedicated, patched, access-controlled, and unavailable to untrusted pull-request code.

---

# 14. Findings summary

| ID | Severity | Area | Finding |
|---|---|---|---|
| P1-01 | P1 | Functional correctness | Lognormal variance/stddev can return zero or `#NUM!` when the correct final moment is finite |
| P1-02 | P1 | Public numerical domain | Extreme F parameters can return inaccurate numeric values despite source claims of clean non-convergence |
| P1-03 | P1 | Release gate | Active inverse tail-residual contracts can become non-blocking characterization when the reference helper fails to import |
| P2-01 | P2 | Special-function contract | `PROB_LogGammaDelta` documents `LargeArg >= 1`, but `PROB_LogBeta` can call it below one |
| P2-02 | P2 | Benchmark reproducibility | Holdout analyzer selects zero contracts after provenance is changed to validated/frozen |
| P2-03 | P2 | CI | Current commit has no visible status/run evidence; benchmark changes have no hosted CI gate |
| P2-04 | P2 | Documentation | Accuracy tables, crossover comments, automation status, and update dates have drifted |
| P2-05 | P2 | Benchmark methodology | Extreme-reference helper mirrors the production continued-fraction arrangement and needs independent cross-check evidence |
| P2-06 | P2 | Maintainability | Lanczos constants and high-precision incomplete-beta helpers have duplicate sources |
| P3-01 | P3 | Repository hygiene | Python bytecode caches are committed and not ignored |
| P3-02 | P3 | API completeness | Signed standardization overflow is discarded rather than mapped to exact CDF/survival limits |
| P3-03 | P3 | Domain policy | General Normal retains a conservative `1E100` restriction that its guarded arithmetic no longer requires |
| P3-04 | P3 | Performance | No reproducible timing baseline is committed |

---

# 15. Detailed remediation plan

## Priority 1 — Correct Lognormal moment reconstruction

### Files

```text
src/M_STATS_PROBDIST_NORMALFAMILY.bas
src/M_STATS_PROBDIST_CORE.bas or a private Normal-family helper
tests/M_STATS_PROBDIST_TEST.bas
benchmark/generate_reference_values.py
benchmark/probability_accuracy_grid.csv
```

### Deliverables

1. stable `LogExpm1Positive`;
2. one-log reconstruction for variance;
3. one-log reconstruction for standard deviation;
4. cancellation regressions;
5. external benchmark points spanning finite cancellation, true underflow, and true overflow.

### Exit criterion

All finite examples return the correct finite result, and current contracts remain green.

---

## Priority 2 — Resolve extreme F domain semantics

### Files

```text
src/M_STATS_PROBDIST_SPECIALFUNCS.bas
src/M_STATS_PROBDIST_TFAMILY.bas
README.md
benchmark/numerical_limitations.csv
tests/M_STATS_PROBDIST_TEST.bas
```

### Deliverables

1. remove the false guarantee that parameters outside the approximate range necessarily fail cleanly;
2. select strict-versus-best-effort public policy;
3. expose the operational domain clearly;
4. add boundary tests;
5. define whether a returned result outside the validated domain is contractual.

### Exit criterion

No public documentation suggests that a successful return outside the validated range necessarily meets the accuracy contract.

---

## Priority 3 — Make every active accuracy contract blocking

### File

```text
benchmark/compute_errors.py
```

### Deliverables

1. narrow exception handling;
2. preserve import error detail;
3. mark unavailable active evaluators as blocking;
4. fail on active `CHARACTERIZATION ONLY`;
5. assert full evaluator coverage.

### Exit criterion

Deleting `_ibeta.py` or breaking its import causes non-zero exit.

---

## Priority 4 — Restore holdout reproducibility

### Files

```text
benchmark/holdout/analyze_holdout.py
benchmark/holdout/README.md
benchmark/holdout/holdout_summary.md
```

### Deliverables

1. evaluate validated/frozen contracts;
2. generate a committed Markdown summary;
3. list threshold margins and worst locations;
4. return non-zero on failure or missing observations.

---

## Priority 5 — Align source, contracts, and public documentation

Update:

- Special Functions public-surface list;
- LogBeta crossover status;
- T-family F inverse thresholds;
- Normal/Lognormal benchmark tables;
- README automation status;
- extreme-F public boundary;
- stale update dates.

Automate table generation wherever possible.

---

## Priority 6 — Add hosted benchmark CI

Create a standard GitHub-hosted workflow for:

- Python syntax;
- contract schema;
- accuracy gate;
- generated docs drift;
- holdout verification;
- numerical-limitations schema;
- repository hygiene.

Keep Excel execution on the self-hosted runner.

---

## Priority 7 — Remove duplicate numerical constants and generated files

- centralize Lanczos coefficients;
- centralize `_ibeta.py`;
- delete `__pycache__`;
- update `.gitignore`;
- add a CI rule that rejects tracked cache artifacts.

---

# 16. Release-readiness assessment

## Suitable now

- teaching and numerical demonstrations;
- controlled VBA model components;
- model-validation comparisons;
- quantitative prototypes;
- direct-tail and inverse-tail calculations within measured domains;
- internal libraries where the documented limitation is accepted and independently tested.

## Conditions before a stronger release claim

The following should be closed:

```text
P1-01 Lognormal moment finite-result defect
P1-02 Extreme F public-domain semantics
P1-03 Accuracy-gate soft pass
```

The Excel regression and hosted accuracy checks should then be required branch-protection statuses.

## Regulated or high-stakes use

For banking, actuarial, engineering, or other governed contexts:

- record the exact commit SHA;
- independently validate the required parameter domain;
- archive benchmark outputs and Excel build information;
- enforce change control;
- do not rely on behavior outside the frozen contract domains;
- treat `numerical_limitations.csv` as part of the model specification.

---

# 17. Notable accuracy evidence

The committed summary provides the following representative worst cases:

| Contract | Worst measured error | Threshold |
|---|---:|---:|
| Balanced Beta density | `2.78E-15` relative | `5E-15` |
| Unbalanced Beta density | `1.23E-12` relative | `4E-12` |
| Balanced Beta CDF | `5.43E-15` relative | `2E-14` |
| Unbalanced Beta CDF | `5.30E-11` relative | `1E-10` |
| Unbalanced Beta survival | `9.16E-11` relative | `2E-10` |
| Unbalanced Beta inverse quantile | `4.72E-11` relative | `1E-10` |
| Unbalanced Beta inverse tail residual | `3.94E-10` relative | `1E-9` |
| Validated F CDF | `2.59E-14` relative | `1.1E-10` |
| Validated F inverse quantile | `7.77E-11` relative | `2E-10` |
| Validated F inverse tail residual | `6.91E-11` relative | `2E-10` |
| `PROB_LogBeta` | `5.93E-14` absolute | `2E-13` |
| LogGamma | `3.77E-15` relative | `6.1E-14` |
| LogGamma half difference | `1.53E-14` relative | `2E-14` |
| Standard Normal survival | `1.52E-14` relative | `2E-14` |
| Lognormal density | `2.36E-14` relative | `3E-14` |
| Student t density | `1.93E-14` relative | `2E-14` |
| Student t inverse | `9.96E-14` relative | `3E-12` |
| Gamma CDF | `1.71E-14` relative | `2E-14` |
| Gamma inverse | `9.54E-15` relative | `2E-14` |
| Stirling error | `2.52E-17` absolute | `3E-17` |
| Weibull standard deviation | `4.12E-15` relative | `5E-15` |

These figures support the repository's numerical credibility while also showing why contracts must be function- and regime-specific.

---

# 18. Final verdict

The project combines several qualities rarely found together in a pure-VBA numerical library:

- transparent algorithms;
- strong floating-point awareness;
- direct tail APIs;
- reusable special functions;
- explicit error semantics;
- deep source documentation;
- deterministic regression testing;
- committed high-precision evidence;
- regime-aware contracts;
- holdout validation.

The stable Lanczos log-gamma difference and the regime-specific Beta/F benchmark program are particularly strong examples of evidence-led numerical engineering.

The current commit is not flawless. The Lognormal moment defect is a concrete wrong-result path; extreme F inputs can return inaccurate numbers outside the validated range; and the benchmark gate can soft-pass when a required derived evaluator is unavailable.

Those are focused, remediable issues rather than architectural failures.

> **Final score: 9.0 / 10.**  
> **Classification: advanced professional numerical library with three material release-gate corrections remaining.**

---

## Appendix A — Review checklist

### Production numerics

- [x] finiteness policy reviewed
- [x] parameter-magnitude policy reviewed
- [x] guarded arithmetic reviewed
- [x] underflow/overflow policy reviewed
- [x] direct tail orientation reviewed
- [x] inverse solvers reviewed
- [x] special-function convergence policy reviewed
- [x] LogBeta regime dispatch reviewed
- [x] moment formulas reviewed
- [x] support-edge behavior reviewed

### Assurance

- [x] primary VBA test harness reviewed
- [x] CI runner reviewed
- [x] workflow trust boundary reviewed
- [x] accuracy-contract schema reviewed
- [x] main benchmark summary reviewed
- [x] unbalanced Beta/F evidence reviewed
- [x] inverse residual measurement reviewed
- [x] holdout design reviewed
- [x] limitations register reviewed

### Repository quality

- [x] README reviewed
- [x] source headers reviewed
- [x] generated-document process reviewed
- [x] `.gitignore` reviewed
- [x] committed generated artifacts reviewed
- [x] public-domain statements compared with measured evidence

---

## Appendix B — Evidence confidence

| Conclusion | Confidence |
|---|---|
| Architecture and modularity | High |
| Public API inventory | High |
| Lognormal moment defect | High — follows directly from committed formulas and representable counterexamples |
| LogBeta stable-delta implementation | High |
| Main accuracy-summary results | High as committed evidence; not independently re-executed |
| Holdout threshold freeze | Medium-high; grid is populated, analyzer currently needs reproducibility repair |
| Extreme F limitation | High — explicitly recorded in committed numerical-limitations evidence |
| Current CI operational status | Low-to-medium — workflow reviewed, no current run/status evidence available |
| Performance characteristics | Medium — algorithmic review only; no timing execution |

---

## Appendix C — Suggested issue titles

1. `Fix finite Lognormal variance/stddev lost through split exponential reconstruction`
2. `Align extreme-F public domain with incomplete-beta validated accuracy range`
3. `Make unavailable tail-residual evaluator block the strict accuracy gate`
4. `Validate PROB_LogGammaDelta for unbalanced arguments below one`
5. `Make frozen-contract holdout analyzer reproducible`
6. `Add hosted Python accuracy and documentation CI`
7. `Synchronize source accuracy comments with regime-aware contracts`
8. `Centralize Lanczos coefficients and benchmark incomplete-beta reference helper`
9. `Remove tracked Python cache files and update .gitignore`
10. `Publish current Excel CI run evidence and require branch checks`
