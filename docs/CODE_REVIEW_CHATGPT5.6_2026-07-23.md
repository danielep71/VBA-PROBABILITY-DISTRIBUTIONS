# Independent Code Review — VBA Probability Distributions

> **Repository:** [`danielep71/VBA-PROBABILITY-DISTRIBUTIONS`](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS)  
> **Branch reviewed:** `main`  
> **Commit reviewed:** [`e0ee3cb308e2d1956055262870649afb3ff61e9c`](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/commit/e0ee3cb308e2d1956055262870649afb3ff61e9c)  
> **Review date:** 23 July 2026  
> **Reviewer:** OpenAI ChatGPT 5.6 Thinking — extra-high reasoning configuration  
> **Suggested repository path:** `docs/INDEPENDENT_CODE_REVIEW_2026-07-23.md`

---

## 1. Executive assessment

### Overall repository score: **8.9 / 10**

### Production numerical library score: **9.1 / 10**

### Numerical assurance and release-engineering score: **8.4 / 10**

This repository is an advanced native-VBA numerical library with a disciplined layered architecture, **112 worksheet-facing probability functions**, reusable special-function kernels, direct survival APIs, direct inverse-survival functions in the Normal family, guarded arithmetic, explicit numerical failure contracts, a large consolidated regression harness, Excel-driven CI, and a substantial high-precision benchmark framework.

The implementation demonstrates unusually strong floating-point awareness for a VBA codebase:

- direct upper-tail calculations rather than indiscriminate `1 - CDF`;
- direct inverse-survival functions where `1 - q` would lose the input probability;
- paired incomplete-beta arguments;
- direct lower and upper incomplete-gamma kernels;
- `Log1p`, `Expm1`, log-density, and log-mass paths;
- guarded addition, multiplication, division, standardization, and affine reconstruction;
- Loader-style Binomial and Poisson mass calculations;
- explicit exact-integer and kernel-backed domains for discrete distributions;
- stable formulas for extreme continuous-Uniform bounds;
- measured numerical envelopes where evidence shows that local convergence is not sufficient;
- consistent `CVErr(xlErrNum)` versus `CVErr(xlErrValue)` classification.

The current assurance layer is also materially stronger than a conventional test-only project. The contract registry contains **132 active, validated-and-frozen contracts**, the main grid contains **1,471 observations**, and the generated summary reports no failed or pending contracts. The complete discrete surface, including Discrete Uniform, is represented in the external contract system.

Three findings nevertheless block an unconditional production-ready conclusion for the complete advertised domain and release process:

1. **`PROB_LogBeta` dispatches to `PROB_LogGammaDelta` outside that kernel's documented precondition.** When both Beta shapes are extremely small and strongly unbalanced, the current accepted-input path can produce silently inaccurate results. A fresh binary64 reproduction found an absolute `LogBeta` error of approximately `8.98E-3` for a valid input pair, leading to approximately **0.90% relative error** in the Beta density at `x = 0.5`.

2. **The strict accuracy analyzer can still PASS an active contract containing `ERROR` observations or malformed references.** Blank ordinary rows are blocked, but the parser maps `ERROR` to an omitted value, the ordinary error loop skips it, and tail-residual contracts bypass the common completeness check.

3. **The tight accuracy evidence is not cryptographically or procedurally bound to the current VBA source revision.** The hosted accuracy workflow checks committed observations, is not triggered by `src/**`, and the observation grid carries no source SHA or module hashes.

### Independent verdict

> **A professionally designed and numerically serious VBA library with strong architecture, broad coverage, and unusually mature validation assets. One confirmed silent numerical defect in the accepted Beta domain and two material release-gate weaknesses prevent an unconditional production-ready rating for the repository as a whole.**

The core design is strong enough that the required remediation is targeted rather than architectural. The highest-value actions are to correct the `LogBeta` dispatch, make every active benchmark row mandatory and parseable, and bind benchmark evidence to the exact source that generated it.

---

# 2. Review scope and methodology

## 2.1 Source basis

The review was performed against the exact `main` revision identified above. Repository files were retrieved by exact path and commit reference, not inferred from filenames, copied snippets, or an unpinned branch snapshot.

The review covered:

- all six production VBA modules;
- the consolidated VBA regression module;
- the PowerShell/COM Excel runner;
- both GitHub Actions workflows;
- the benchmark reference generator;
- the VBA accuracy exporter;
- the accuracy analyzer and strict release gate;
- the machine-readable accuracy-contract registry;
- the committed main observation grid;
- the generated accuracy summary;
- the numerical-limitations registry;
- focused benchmark-study organization;
- the repository README and benchmark documentation;
- release and documentation-generation mechanics.

## 2.2 Execution boundary

Desktop Excel was not available in the review environment. The reviewer therefore did **not**:

- import the modules into the Visual Basic Editor;
- execute `Debug -> Compile VBAProject`;
- run `Test_STATS_PROBDIST_RunAll`;
- independently execute the self-hosted Excel workflow.

The review distinguishes among:

- **confirmed source defects**, established directly from the implementation;
- **confirmed assurance defects**, established directly from scripts and workflows;
- **committed numerical evidence**, produced by the repository's benchmark pipeline;
- **targeted independent numerical analysis**, reproducing current formulas in IEEE-754 binary64 and comparing them with 100-digit arithmetic;
- **operational state not independently evidenced**, such as the availability and branch-protection status of the self-hosted runner.

The GitHub combined-status interface exposed no status records for the reviewed SHA. That is not evidence that checks failed or did not run; it means current required-check enforcement and runner availability were not independently established through the available interface.

## 2.3 Independent numerical reproduction

For the `LogBeta` investigation, the current:

- Lanczos coefficients;
- `PROB_LogGamma` arrangement;
- `PROB_LogGammaDelta` arrangement;
- `PROB_LogBeta` dispatch rule;
- binary64 evaluation order

were reproduced independently and compared with `mpmath` at 100 decimal digits.

This was not a symbolic critique alone. The current numerical path was evaluated for concrete accepted inputs and across a small logarithmic sweep of both-small, unbalanced shapes.

## 2.4 Scoring standard

A score of 10 requires all of the following:

- correct results throughout the documented and accepted public domain;
- explicit and internally consistent contracts;
- no known silent wrong-result path;
- complete deterministic regression coverage;
- current-source execution in CI;
- reproducible independent accuracy evidence;
- evidence tied to the source revision that produced it;
- documentation synchronized with authoritative metadata;
- recorded operational and performance baselines.

---

# 3. Hard repository metrics

## 3.1 Production source size

| Module | Lines | Primary responsibility |
|---|---:|---|
| [`M_STATS_PROBDIST_CORE.bas`](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/blob/e0ee3cb308e2d1956055262870649afb3ff61e9c/src/M_STATS_PROBDIST_CORE.bas) | **911** | Shared constants, predicates, guarded arithmetic, stable elementary functions |
| [`M_STATS_PROBDIST_SPECIALFUNCS.bas`](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/blob/e0ee3cb308e2d1956055262870649afb3ff61e9c/src/M_STATS_PROBDIST_SPECIALFUNCS.bas) | **1,580** | Gamma, beta, combinatorial, and inverse special-function kernels |
| [`M_STATS_PROBDIST_NORMALFAMILY.bas`](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/blob/e0ee3cb308e2d1956055262870649afb3ff61e9c/src/M_STATS_PROBDIST_NORMALFAMILY.bas) | **3,663** | Standard Normal, Normal, and Lognormal |
| [`M_STATS_PROBDIST_TFAMILY.bas`](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/blob/e0ee3cb308e2d1956055262870649afb3ff61e9c/src/M_STATS_PROBDIST_TFAMILY.bas) | **3,363** | Student t, Chi-square, and F |
| [`M_STATS_PROBDIST_CONTINUOUS.bas`](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/blob/e0ee3cb308e2d1956055262870649afb3ff61e9c/src/M_STATS_PROBDIST_CONTINUOUS.bas) | **4,952** | Gamma, Beta, Exponential, Weibull, continuous Uniform |
| [`M_STATS_PROBDIST_DISCRETE.bas`](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/blob/e0ee3cb308e2d1956055262870649afb3ff61e9c/src/M_STATS_PROBDIST_DISCRETE.bas) | **8,156** | Six discrete distributions |
| **Production VBA total** | **22,625** |  |
| [`M_STATS_PROBDIST_TEST.bas`](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/blob/e0ee3cb308e2d1956055262870649afb3ff61e9c/tests/M_STATS_PROBDIST_TEST.bas) | **6,149** | Consolidated regression harness |
| **Production plus primary tests** | **28,774** |  |

These counts are physical lines in the reviewed Git blobs, including comments and blank lines. The codebase is intentionally heavily documented, so line count should not be interpreted as executable-code count.

## 3.2 Public worksheet-facing surface

| Family | Public UDFs |
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

The discrete count includes both ordinary PMF and LogPMF functions.

## 3.3 Regression structure

The consolidated harness registers **96 test-section procedures**:

| Suite | Registered sections |
|---|---:|
| Core and Special Functions | 7 |
| Normal and Lognormal | 19 |
| Student t, Chi-square, and F | 18 |
| Continuous distributions | 27 |
| Discrete distributions | 25 |
| **Total** | **96** |

This count measures registered test sections, not individual assertions. The module maintains assertion counters dynamically and reports the executed total at runtime.

## 3.4 Numerical-assurance assets

| Evidence | Current size / state |
|---|---:|
| Active accuracy contracts | **132** |
| Main benchmark observations | **1,471** |
| Generated contract verdicts | **132 PASS** |
| FAIL | **0** |
| KNOWN LIMITATION contract state | **0** |
| CHARACTERIZATION ONLY | **0** |
| PENDING | **0** |
| Numerical limitations registered separately | **2** |
| GitHub Actions workflows | **2** |

The two registered numerical boundaries are:

- the mitigated incomplete-beta extreme-shape issue for F CDF/SF/inverse;
- the characterized relative-accuracy degradation of Normal-family survival tails beyond the tight central regime.

The generated summary correctly distinguishes zero open `KNOWN LIMITATION` contract states from the separate limitations register.

---

# 4. Weighted scorecard

| Area | Weight | Score | Weighted contribution |
|---|---:|---:|---:|
| Functional correctness | 18% | **8.7** | 1.566 |
| Numerical robustness | 17% | **9.1** | 1.547 |
| Architecture and modularity | 10% | **9.5** | 0.950 |
| Public API design | 8% | **9.4** | 0.752 |
| Error handling and diagnostics | 7% | **9.3** | 0.651 |
| Regression testing | 10% | **9.3** | 0.930 |
| External accuracy assurance | 11% | **8.7** | 0.957 |
| CI and release engineering | 8% | **8.1** | 0.648 |
| Documentation | 5% | **8.0** | 0.400 |
| Maintainability and repository hygiene | 5% | **8.7** | 0.435 |
| Performance engineering | 1% | **8.3** | 0.083 |
| **Total** | **100%** |  | **8.919 / 10** |

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

The score sits immediately below 9 because the confirmed `LogBeta` path is a silent wrong-result defect inside the currently accepted public Beta domain, while the two assurance findings weaken the reliability of a green accuracy-gate verdict.

---

# 5. Component scores

| Component | Score | Assessment |
|---|---:|---|
| `M_STATS_PROBDIST_CORE` | **9.4** | Strong single-source numerical primitives, careful overflow contracts, and good separation between finiteness and project magnitude policy |
| `M_STATS_PROBDIST_SPECIALFUNCS` | **8.5** | Sophisticated reusable kernels, but the current `LogBeta` dispatch violates the delta kernel's documented precondition |
| `M_STATS_PROBDIST_NORMALFAMILY` | **9.5** | Excellent direct-tail design, deep-tail inverse contracts, stable intervals, and guarded location-scale arithmetic |
| `M_STATS_PROBDIST_TFAMILY` | **9.2** | Strong density/tail/inverse design and a measured F envelope; source accuracy prose is not fully synchronized with contracts |
| `M_STATS_PROBDIST_CONTINUOUS` | **8.9** | Broad and well structured, with strong Uniform and Weibull engineering; Beta inherits the `LogBeta` defect |
| `M_STATS_PROBDIST_DISCRETE` | **9.4** | Complete six-family layer, direct tails, LogPMFs, explicit domains, stable mass kernels, and external contracts for the full surface |
| `M_STATS_PROBDIST_TEST` | **9.3** | Large, consolidated, dependency-ordered harness with strong regression design; it does not cover the confirmed tiny/tiny `LogBeta` regime |
| External benchmark framework | **8.5** | Broad, regime-aware, independently held out, and all contracts frozen; completeness parsing and source-provenance gaps remain |
| Excel/PowerShell CI | **9.0** | Imports the full source set, runs every suite, captures counters, and limits untrusted fork execution |
| Documentation | **8.0** | Extensive and technically valuable, but source comments, benchmark prose, and generator comments contain material stale claims |

---

# 6. Principal strengths

## 6.1 Coherent numerical architecture

The dependency design is clear and professionally structured:

```text
CORE
  -> SPECIALFUNCS
      -> distribution-family modules
          -> consolidated tests
              -> Excel CI
```

`CORE` and `SPECIALFUNCS` use `Option Private Module`, allowing project-scoped reuse without exposing implementation helpers in the worksheet Function Wizard.

The architecture avoids the two most common VBA numerical-library failures:

- copying private versions of the same predicate or constant into multiple modules;
- embedding distribution-independent special functions inside one public family.

The module headers make numerical ownership explicit, and the implementation generally follows that ownership.

## 6.2 Strong guarded-arithmetic contract

The shared core provides explicit `Try` routines for:

- exponentiation;
- addition;
- multiplication;
- division;
- affine reconstruction;
- standardization.

Predictable overflow returns `False` to the caller, which is then mapped to `#NUM!`. Negative exponential underflow is classified as a valid zero. This is substantially safer than relying on a broad `On Error` handler at the outer UDF layer.

The distinction between:

- true IEEE-754 finiteness;
- the coarse `1E100` representational guard;
- kernel-specific convergence or accuracy evidence

is conceptually correct and explicitly documented.

## 6.3 Direct-tail design

The library treats survival probability as a first-class numerical quantity.

Examples include:

- Standard Normal, Normal, and Lognormal direct survival;
- Student t, Chi-square, F, Gamma, Beta, Exponential, Weibull, and Uniform direct survival;
- all six discrete families;
- direct inverse survival for the Normal and Lognormal families.

This design protects small upper-tail probabilities from cancellation and protects tiny inverse-tail inputs from loss in `1 - q`.

The current contract registry also separates central, split-boundary, and deep-tail inverse regimes rather than making one unsupported universal claim.

## 6.4 Paired special-function arguments

The incomplete-beta API receives both `X` and its complement from the caller rather than reconstructing one by subtraction. This is a high-value design choice for:

- Student t near its center;
- F with extreme ratios;
- Beta tails;
- Binomial and Negative Binomial tails.

The incomplete-gamma layer similarly exposes direct lower and upper regularized functions.

## 6.5 Measured F envelope

The F CDF, survival, and inverse do not rely on local continued-fraction convergence as a proxy for output accuracy.

A dedicated study showed that unguarded errors can grow materially while the local convergence condition still passes. The public implementation therefore enforces a measured `df <= 1E5` envelope for the incomplete-beta-backed F surface, while leaving the closed-form F density outside that restriction.

This is the correct pattern:

1. characterize the numerical regime;
2. determine whether silent wrong output is possible;
3. enforce a public boundary only where evidence supports it.

## 6.6 Complete discrete layer

The discrete module is a substantial and technically coherent addition.

Notable strengths include:

- Loader-style Binomial and Poisson log-mass kernels;
- LogPMF functions that remain useful after ordinary PMF underflow;
- direct incomplete-beta and incomplete-gamma tails;
- explicit lower-bound integer inverse semantics;
- Negative Binomial reuse of the stable Binomial mass structure;
- Hypergeometric log-mass construction and near-tail ratio summation;
- explicit population and summation ceilings;
- signed Discrete Uniform bounds;
- real-threshold step behavior for Discrete Uniform CDF/SF;
- corrected adjacent-step quantile checks;
- stable factored moment formulas;
- external accuracy and independent holdout coverage for all six families.

## 6.7 Consolidated regression harness

The test module provides one authoritative counter set and one final verdict.

Its design includes:

- dependency-ordered suites;
- exact constant checks;
- direct kernel tests;
- support-edge tests;
- inverse minimality and round-trips;
- complement identities;
- deep-tail regressions;
- error-code assertions;
- diagnostic-status checks;
- named regression cases tied to historical numerical failure modes.

The PowerShell runner imports every production module plus the tests and executes every suite, including the Discrete suite.

## 6.8 Mature contract registry

The benchmark framework is materially stronger than a collection of hand-written examples.

It contains:

- machine-readable contracts;
- explicit regimes;
- appropriate error metrics;
- independent reference generation;
- high/low export reconstruction;
- dedicated studies;
- independent holdouts;
- a strict generated verdict summary;
- a separate numerical-limitations register.

All 132 current contracts are marked validated and frozen. This is a significant assurance asset.

---

# 7. Findings summary

| ID | Severity | Area | Finding |
|---|---|---|---|
| P1-01 | P1 | Numerical correctness | `PROB_LogBeta` calls `PROB_LogGammaDelta` outside its documented precondition and can be materially inaccurate when both unbalanced shapes are extremely small |
| P1-02 | P1 | Accuracy gate | Active contracts can PASS with `ERROR` observations, malformed references, or incomplete tail-residual rows |
| P1-03 | P1 | Evidence provenance | Committed tight-accuracy observations are not bound to the VBA source revision that allegedly produced them |
| P2-01 | P2 | Documentation governance | Source comments, benchmark prose, and generator comments materially disagree with the authoritative contract registry and current assurance behavior |
| P2-02 | P2 | Public-domain governance | Some special-function-backed surfaces accept parameters far beyond the measured accuracy regime; successful return is explicitly not an accuracy claim |
| P2-03 | P2 | Maintainability | The 8,156-line Discrete module and 6,149-line test module are approaching a reviewability boundary |
| P3-01 | P3 | Performance evidence | No reproducible timing baseline is committed |
| P3-02 | P3 | Static assurance | No automated source/API inventory, duplicate-procedure, line-continuation, or generated-document freshness gate was identified |
| P3-03 | P3 | Operational assurance | Current required-check enforcement and self-hosted-runner availability were not independently evidenced for the reviewed SHA |

---

# 8. Detailed findings

## P1-01 — `PROB_LogBeta` violates the delta kernel's precondition for tiny unbalanced shapes

### Severity

**P1 — release-blocking numerical correctness defect**

### Current source contract

[`PROB_LogGammaDelta`](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/blob/e0ee3cb308e2d1956055262870649afb3ff61e9c/src/M_STATS_PROBDIST_SPECIALFUNCS.bas#L272-L368) states:

```text
LargeArg >= 1
Increment > 0
```

and documents measured accuracy for increments and large arguments within its studied regime.

[`PROB_LogBeta`](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/blob/e0ee3cb308e2d1956055262870649afb3ff61e9c/src/M_STATS_PROBDIST_SPECIALFUNCS.bas#L371-L460) orders the two arguments and selects the delta path whenever:

```vb
If SmallArg / LargeArg < PROB_LOGBETA_STABLE_RATIO Then
    PROB_LogBeta = _
        PROB_LogGamma(SmallArg) - _
        PROB_LogGammaDelta(LargeArg, SmallArg)
End If
```

The caller does **not** require `LargeArg >= 1`.

Public Beta validation accepts any strictly positive finite shapes inside the broad project magnitude guard. There is no positive lower bound that prevents both shapes from being far below one.

### Confirmed counterexamples

#### Counterexample 1

```text
Alpha = 1E-12
Beta  = 9.9E-14
Small / Large = 0.099
```

The stable-delta path is selected.

| Quantity | Value |
|---|---:|
| 100-digit reference `LogBeta` | `30.03805722019758` |
| Current binary64 route | `30.03805921300609` |
| Absolute error | `1.9928085E-6` |
| Direct three-log-gamma identity error | approximately `7.1E-15` |
| Resulting Beta-density relative error at `x = 0.5` | approximately `-1.9928E-6` |

#### Counterexample 2

```text
Alpha = 1E-16
Beta  = 9.9E-18
Small / Large = 0.099
```

| Quantity | Value |
|---|---:|
| 100-digit reference `LogBeta` | `39.24839759217376` |
| Current binary64 route | `39.23941402347159` |
| Absolute error | `8.9835687E-3` |
| Resulting Beta-density relative error at `x = 0.5` | approximately `+9.024E-3` |

The density error is approximately **0.9024%**.

A sweep over:

```text
LargeArg = 1E-16 ... 1
ratio    = 0.001, 0.01, 0.05, 0.099
```

found:

```text
maximum current stable-route absolute LogBeta error  ~8.98E-3
maximum direct-identity absolute LogBeta error        ~1.42E-14
```

### Why this happens

The delta arrangement was designed to avoid subtracting two enormous log-gamma values when `LargeArg` is genuinely large.

When `LargeArg` is extremely small:

- the Lanczos shifted denominators enter a different conditioning regime;
- the stated validation evidence no longer applies;
- the stable-delta route is not intrinsically superior;
- the ordinary three-log-gamma identity is well conditioned because none of the three values is enormous.

The problem is not the existence of the delta kernel. The problem is dispatching to it outside the domain for which it was designed and measured.

### Current test gap

The regression harness explicitly tests some subunit structural identities for `PROB_LogGammaDelta`, but the representative arguments are around `0.2`, `0.5`, and `0.55`. It does not include both-small inputs near `1E-12` or `1E-16`.

The main `PROB_LogBeta` contract currently reports `108/108 PASS`, with a worst absolute log error of `5.93E-14`. The confirmed counterexamples therefore sit outside the committed grid.

### Recommended correction

The minimum safe dispatch is to enforce the documented delta precondition:

```vb
If LargeArg >= 1# Then
    If SmallArg / LargeArg < PROB_LOGBETA_STABLE_RATIO Then
        PROB_LogBeta = _
            PROB_LogGamma(SmallArg) - _
            PROB_LogGammaDelta(LargeArg, SmallArg)
        Exit Function
    End If
End If

PROB_LogBeta = _
    PROB_LogGamma(A) + _
    PROB_LogGamma(B) - _
    PROB_LogGamma(A + B)
```

A lower crossover may be possible, but it should be introduced only after a dedicated seam study establishes a measured boundary.

### Required tests and evidence

Add:

1. exact regression cases at:
   - `(1E-12, 9.9E-14)`;
   - `(1E-16, 9.9E-18)`;
   - symmetric argument order;
2. a logarithmic tiny/tiny sweep;
3. Beta density, CDF, survival, and inverse cases using both-small shapes;
4. a separate contract regime such as:
   - `tiny_unbalanced`;
5. a seam study around the selected `LargeArg` lower boundary.

### Acceptance criteria

- No accepted tiny/tiny case should use a kernel outside its documented precondition.
- `PROB_LogBeta` absolute error should meet an explicit frozen contract in the new regime.
- Downstream Beta density/CDF/SF/inverse contracts should include both-small shapes.
- The production source, test registry, generator, exporter, contracts, and documentation should change together.

---

## P1-02 — The accuracy gate can PASS incomplete or invalid active evidence

### Severity

**P1 — release-gate integrity defect**

### Current behavior

In [`compute_errors.py`](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/blob/e0ee3cb308e2d1956055262870649afb3ff61e9c/benchmark/compute_errors.py#L54-L101):

```python
def parse_observed(s):
    s = s.strip()
    if s == "" or s.upper() == "ERROR":
        return None
```

Ordinary error measurement then does:

```python
o = parse_observed(...)
if o is None:
    continue
```

Malformed reference values are also silently skipped:

```python
try:
    ref = Decimal(...)
except InvalidOperation:
    continue
```

The strict completeness check later counts only literal blank observations:

```python
n_missing = sum(
    1 for r in matched
    if not r["observed_vba"].strip()
)
```

An `ERROR` token is nonblank, so it is not counted as missing.

The `tail_probability_residual` branch is evaluated before this common completeness check and exits with `continue`, so active tail-residual contracts do not receive the ordinary strict-row validation.

### Confirmed synthetic reproduction

Using the current parser and ordinary measurement logic:

```text
row 1: observed = 1.0,   reference = 1.0
row 2: observed = ERROR, reference = 1000
```

produces:

```text
blank-row count       = 0
usable observations   = 1
worst measured error  = 0
threshold              = 1E-12
verdict                = PASS
displayed points       = 1/2
```

The contract can therefore pass even though half of its intended evidence failed to produce a numeric result.

The same class of issue affects:

- malformed references;
- `ERROR` observations;
- tail-residual rows that are blank or `ERROR`;
- potentially malformed required arguments in tail evaluators.

### Impact

The strict gate's user-facing guarantee is stronger than its actual implementation.

A green result can mean:

> all **usable** rows met the threshold

rather than:

> every required row was successfully produced, parsed, and evaluated.

For a numerical release gate, that distinction is material.

### Recommended correction

Validate every matched active row before dispatching to a measure-specific evaluator.

A robust preflight should require:

- nonblank observation;
- observation not equal to `ERROR`;
- successful Decimal parsing of every `hi;lo` part;
- successful Decimal parsing of the reference;
- all required arguments present and parseable;
- evaluator availability.

Conceptually:

```python
def validate_active_rows(contract, rows):
    failures = []

    for index, row in enumerate(rows):
        try:
            observed = parse_observed_strict(row["observed_vba"])
            reference = Decimal(row["reference"].strip())
            validate_required_arguments(contract, row)
        except Exception as exc:
            failures.append((index, str(exc)))

    return failures
```

If any required row fails preflight, the contract should be:

- `PENDING` with a blocking exit code; or
- `FAIL` if the project treats export/evidence corruption as a failure.

It should never be silently omitted.

### Required unit tests

Add release-gate tests for:

1. blank ordinary observation;
2. `ERROR` ordinary observation;
3. malformed `hi;lo` observation;
4. malformed reference;
5. blank tail-residual observation;
6. `ERROR` tail-residual observation;
7. missing tail argument;
8. a mixed valid/invalid contract that must not PASS.

### Acceptance criteria

- `usable rows == matched rows` for every active contract.
- The points column must never show a PASS with `n / total` where `n < total`.
- Every invalid evidence row must be listed in the gate output.
- The gate must exit nonzero when any active row is unevaluated.

---

## P1-03 — Accuracy evidence is not bound to the source that generated it

### Severity

**P1 — release provenance defect**

### Current workflow

The hosted [`accuracy-gate.yml`](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/blob/e0ee3cb308e2d1956055262870649afb3ff61e9c/.github/workflows/accuracy-gate.yml#L1-L59):

- checks out the repository;
- runs `compute_errors.py` against the committed grid;
- is triggered only by:
  - `benchmark/**`;
  - the workflow file itself.

It is not triggered by `src/**` or `tests/**`.

The main grid schema contains:

```text
function
vba_kernel
claim
metric
arg1 ... arg4
reference
observed_vba
regime
evidence_set
```

It does not contain:

- source commit SHA;
- module blob hashes;
- test-module hash;
- exporter hash;
- Excel version/build;
- export timestamp linked to source;
- a signed or generated observation manifest.

### Consequence

A numerical algorithm can change in `src/**` while:

- the committed observations remain unchanged;
- the hosted accuracy workflow does not run;
- the old observations continue to satisfy every contract;
- the repository still presents a green tight-accuracy summary.

The Excel regression workflow does run current source, but its role is behavioral regression with intentionally broader tolerances. It is not a substitute for regenerating the 132 tight external contracts.

### Recommended target design

The strongest design is a two-stage workflow:

```text
Stage 1 — self-hosted Windows/Excel
    import exact current source
    export observations
    write source and environment manifest
    upload artifact

Stage 2 — hosted Python
    download artifact
    verify manifest against checked-out source
    evaluate every contract
    publish summary
```

The observation manifest should include at least:

```text
repository
commit_sha
branch_or_ref
module_path -> blob_sha
test_module_sha
exporter_sha
Excel version
Excel build
Office bitness
export timestamp
grid schema version
contract registry sha
```

### Minimum acceptable hardening

If automatic regeneration is not yet practical:

1. trigger the accuracy workflow on `src/**`;
2. add a committed manifest with source SHA and module hashes;
3. make the analyzer reject stale or mismatched manifests;
4. require both:
   - Excel regression;
   - strict accuracy gate
   in branch protection;
5. document the observation-generation commit in the generated summary.

### Acceptance criteria

A green accuracy result should prove:

> These exact observations were produced by this exact source revision under this recorded environment, and every active row was evaluated.

---

## P2-01 — Numerical documentation has materially drifted from authoritative contracts

### Severity

**P2 — documentation and contract-governance defect**

### Confirmed inconsistencies

#### Stirling error

[`M_STATS_PROBDIST_SPECIALFUNCS.bas`](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/blob/e0ee3cb308e2d1956055262870649afb3ff61e9c/src/M_STATS_PROBDIST_SPECIALFUNCS.bas#L492-L515) states:

```text
absolute error <= 3E-17
```

The authoritative contract is:

```text
1E-13 absolute
```

and records that the prior `3E-17` claim was overfit to its grid, with an independent holdout worst near `3.57E-14`.

#### F inverse

[`M_STATS_PROBDIST_TFAMILY.bas`](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/blob/e0ee3cb308e2d1956055262870649afb3ff61e9c/src/M_STATS_PROBDIST_TFAMILY.bas#L108-L123) states:

```text
F quantile <= 5.9E-13 relative error
```

The authoritative inverse contracts are:

```text
2E-10 relative quantile error
2E-10 relative tail residual
```

inside the measured F envelope.

#### Benchmark README completeness wording

[`benchmark/README.md`](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/blob/e0ee3cb308e2d1956055262870649afb3ff61e9c/benchmark/README.md#L60-L84) says empty observations are excluded so a partial export produces a partial summary.

The current analyzer actually blocks blank ordinary active rows, while still incorrectly skipping `ERROR` rows and bypassing completeness for tail residuals. The prose therefore describes neither the intended strict policy nor the exact current implementation.

#### Reference generator comments

[`generate_reference_values.py`](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/blob/e0ee3cb308e2d1956055262870649afb3ff61e9c/benchmark/generate_reference_values.py#L1-L30) repeats the stale Stirling claim.

Its discrete section comments still describe only Binomial, Poisson, and Geometric and state that no discrete contract is active, even though the current registry contains validated-and-frozen contracts for six discrete families.

### Impact

This drift creates multiple competing sources of truth:

- source comments;
- generator comments;
- benchmark prose;
- contract CSV;
- generated summary.

For a numerical library, stale accuracy text is not cosmetic. It can lead users to rely on thresholds that the independent evidence does not support.

### Recommended correction

Make `accuracy_contracts.csv` authoritative for all published measured thresholds.

Use generated documentation for:

- contract counts;
- family lists;
- thresholds;
- provenance;
- limitations.

The repository already contains `render_contract_table.py`. Extend CI to:

```text
python benchmark/render_contract_table.py --write
git diff --exit-code benchmark/README.md
```

Also add a lightweight source-comment consistency check, or replace exact duplicated threshold claims in source headers with stable references such as:

```text
See benchmark/accuracy_contracts.csv for the authoritative,
regime-specific measured accuracy contract.
```

### Acceptance criteria

- No published measured threshold disagrees with the contract registry.
- No generator comment claims inactive coverage for active families.
- Generated documentation is verified in CI.
- Numerical limitations and accuracy contracts remain distinct.

---

## P2-02 — Some accepted parameter domains extend beyond measured accuracy evidence

### Severity

**P2 — public numerical-governance risk**

### Current design

The project correctly distinguishes the broad `1E100` guard from measured accuracy.

The T-family source explicitly states that:

- F CDF/SF/inverse enforce a measured `df <= 1E5` envelope;
- Student t and Chi-square are accuracy-validated to roughly `1E9`;
- larger degrees of freedom up to `1E100` may be accepted and attempted;
- a successful result there does not imply contract accuracy.

This is honest documentation, and an arbitrary hard cutoff should not be introduced without evidence.

### Residual risk

A worksheet user can still receive a finite successful value from a region for which the project expressly makes no accuracy claim.

That creates three distinguishable states:

1. invalid mathematical input;
2. valid and accuracy-contracted input;
3. valid, computed, but not accuracy-contracted input.

The public API returns a number for states 2 and 3 without a machine-readable distinction unless the user has read the documentation.

### Recommended approach

Do not replace the broad guard with arbitrary kernel limits.

Instead:

1. extend measured envelope studies for large Student t and Chi-square regimes;
2. determine whether successful but inaccurate values occur;
3. enforce a limit only where evidence shows a silent wrong-result risk;
4. otherwise expose a documented capability indicator or validation helper, for example:
   - `K_STATS_*_IsWithinAccuracyContract`;
   - a public contract table;
   - a Status warning for VBA callers, if worksheet semantics remain acceptable.

### Acceptance criteria

- Every public function clearly distinguishes accepted mathematical domain from measured accuracy domain.
- Any enforced limit is supported by a committed study.
- Silent success outside a measured range is characterized before it is described as robust.

---

## P2-03 — Module size is approaching a reviewability boundary

### Severity

**P2 — maintainability and review-risk issue**

### Current scale

```text
M_STATS_PROBDIST_DISCRETE.bas   8,156 lines
M_STATS_PROBDIST_TEST.bas       6,149 lines
```

The modules remain coherent, but their size raises practical costs:

- manual review becomes slower;
- merge conflicts become more likely;
- test registration is easier to overlook;
- unrelated family changes share one file;
- future random-variate or array work could blur numerical responsibilities.

### Recommended boundary

There is no urgent need to split merely for aesthetics. A split should preserve public API compatibility and numerical ownership.

Before adding major new surfaces, consider:

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

or another evidence-based family boundary.

Random variates and array/range wrappers should **not** be appended to this module:

```text
M_STATS_RNG_CORE
M_STATS_PROBDIST_RANDOM
M_STATS_PROBDIST_ARRAY
```

The scalar numerical kernels should remain single-source.

### Acceptance criteria

- No public `K_STATS_*` signature changes solely because of a file split.
- No numerical formula is duplicated.
- CI imports the new module set automatically.
- API inventory and test registration are generated or statically checked.

---

## P3-01 — No reproducible performance baseline

### Severity

**P3 — performance-evidence gap**

No committed timing harness was identified for:

- common CDF/SF calls;
- inverse solvers;
- large-count PMFs;
- Hypergeometric summation;
- import/compile/test duration;
- comparison with `WorksheetFunction` marshalling where relevant.

The code is generally efficient in design, but “fast” and “performance-oriented” remain qualitative without an environment-recorded baseline.

### Recommended benchmark set

Record:

- Excel version/build;
- 32-bit or 64-bit Office;
- CPU;
- Windows version;
- calculation mode;
- warm-up policy;
- repetitions;
- median and percentile timings.

Suggested cases:

```text
1,000,000 Standard Normal CDF calls
100,000 Normal inverse calls
100,000 Beta CDF calls in balanced and unbalanced regimes
100,000 Binomial/Poisson PMF calls
Hypergeometric CDF near the summation budget
complete regression-suite wall time
```

Performance tests should be nonblocking initially and become regression gates only after stable baselines exist.

---

## P3-02 — Static source and generated-artifact checks are incomplete

### Severity

**P3 — repository-hygiene gap**

No automated static gate was identified for:

- duplicate VBA procedure names;
- duplicate module names;
- malformed exported-module headers;
- broken line continuations;
- public UDF inventory;
- test-section registration;
- generator/exporter dispatch completeness;
- source-comment versus contract consistency;
- generated README freshness.

These checks do not require Excel and can run cheaply on every pull request.

### Recommended static checker

Add a Python script that:

1. parses exported `.bas` modules;
2. verifies `Attribute VB_Name`;
3. inventories public procedures;
4. rejects duplicate names;
5. verifies expected module dependency inventory;
6. compares public UDFs with generated API documentation;
7. checks every `Test_*` section is registered;
8. validates benchmark exporter dispatch for every contracted function;
9. regenerates contract tables and fails on diff.

---

## P3-03 — Operational enforcement was not independently evidenced

### Severity

**P3 — operational assurance not confirmed**

The workflow definitions are strong, but the combined-status interface returned no status records for the reviewed SHA.

The review therefore could not independently establish:

- whether both workflows executed for the reviewed commit;
- whether the self-hosted Excel runner was online;
- whether checks are required by branch protection;
- whether direct pushes can bypass them.

### Recommendation

Document and expose:

- required status checks;
- runner health/availability process;
- release evidence linking workflow run IDs;
- branch-protection expectations;
- policy for documentation-only versus numerical changes.

---

# 9. Remediation roadmap

## Phase 1 — Release blockers

### 1. Correct `PROB_LogBeta` dispatch

- enforce `LargeArg >= 1#` for the current delta path;
- add tiny/tiny regression and benchmark regimes;
- regenerate all affected Beta evidence.

### 2. Make the accuracy gate row-complete

- reject `ERROR`;
- reject malformed observations and references;
- preflight tail-residual rows;
- require `usable == matched`;
- add negative gate tests.

### 3. Bind evidence to source

- add commit and module hashes;
- trigger on `src/**`;
- reject stale manifests;
- connect Excel export artifacts to hosted analysis.

## Phase 2 — Contract and documentation hardening

### 4. Generate accuracy prose from contracts

- remove stale exact thresholds from manually maintained text;
- make generated documentation a CI check;
- update generator comments and family inventories.

### 5. Characterize broad accepted domains

- study large Student t and Chi-square regimes;
- enforce limits only where measured evidence justifies them;
- clearly label computed-but-uncontracted regions.

## Phase 3 — Maintainability and operations

### 6. Introduce static exported-VBA checks

- API inventory;
- duplicate procedures;
- test registration;
- benchmark dispatch completeness;
- generated-file freshness.

### 7. Establish performance baselines

- record deterministic timing cases and environment;
- publish nonblocking trend artifacts.

### 8. Prepare module boundaries for future expansion

- keep random generation and array APIs separate;
- consider a discrete split before materially enlarging the current 8,156-line module.

---

# 10. Release-readiness checklist

A release candidate for the complete advertised surface should satisfy:

```text
[ ] Import all current production modules into a clean workbook
[ ] Debug -> Compile VBAProject
[ ] Run Test_STATS_PROBDIST_RunAll
[ ] Confirm RESULT: ALL TESTS PASSED
[ ] Run Excel/VBA CI on the exact release SHA
[ ] Regenerate observations from the exact release source
[ ] Verify source/module hashes in the observation manifest
[ ] Run the strict accuracy gate
[ ] Confirm every active contract has usable == matched observations
[ ] Review numerical_limitations.csv
[ ] Review all generated documentation diffs
[ ] Record workflow run IDs and Excel environment
[ ] Tag the exact release commit
```

For the reviewed revision, the first four items were not independently executed in this review environment. The committed evidence and source were inspected, but that is not a substitute for release execution in desktop Excel.

---

# 11. Final conclusion

The repository is substantially more sophisticated than a typical VBA statistical library. Its numerical architecture, direct-tail design, guarded arithmetic, discrete implementation, error policy, and validation assets are all professional strengths.

The codebase is close to an advanced production-grade standard, but the standard must be applied to the complete accepted domain and to the evidence pipeline itself.

The decisive issues are narrow and concrete:

- one accepted Beta regime currently returns silently inaccurate values;
- one strict gate can still pass incomplete evidence;
- the tight benchmark evidence is not tied to the source revision it is used to support.

Correcting those issues would remove the principal barriers to a score above 9 and to a much stronger production-readiness conclusion.

> **Final rating: 8.9 / 10 — advanced professional numerical library, requiring targeted correctness and release-evidence hardening before unconditional production use.**

---

# Appendix A — Reviewed files

## Production VBA

- [`src/M_STATS_PROBDIST_CORE.bas`](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/blob/e0ee3cb308e2d1956055262870649afb3ff61e9c/src/M_STATS_PROBDIST_CORE.bas)
- [`src/M_STATS_PROBDIST_SPECIALFUNCS.bas`](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/blob/e0ee3cb308e2d1956055262870649afb3ff61e9c/src/M_STATS_PROBDIST_SPECIALFUNCS.bas)
- [`src/M_STATS_PROBDIST_NORMALFAMILY.bas`](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/blob/e0ee3cb308e2d1956055262870649afb3ff61e9c/src/M_STATS_PROBDIST_NORMALFAMILY.bas)
- [`src/M_STATS_PROBDIST_TFAMILY.bas`](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/blob/e0ee3cb308e2d1956055262870649afb3ff61e9c/src/M_STATS_PROBDIST_TFAMILY.bas)
- [`src/M_STATS_PROBDIST_CONTINUOUS.bas`](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/blob/e0ee3cb308e2d1956055262870649afb3ff61e9c/src/M_STATS_PROBDIST_CONTINUOUS.bas)
- [`src/M_STATS_PROBDIST_DISCRETE.bas`](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/blob/e0ee3cb308e2d1956055262870649afb3ff61e9c/src/M_STATS_PROBDIST_DISCRETE.bas)

## Tests and Excel CI

- [`tests/M_STATS_PROBDIST_TEST.bas`](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/blob/e0ee3cb308e2d1956055262870649afb3ff61e9c/tests/M_STATS_PROBDIST_TEST.bas)
- [`ci/Run-ExcelVbaTests.ps1`](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/blob/e0ee3cb308e2d1956055262870649afb3ff61e9c/ci/Run-ExcelVbaTests.ps1)
- [`.github/workflows/excel-vba-regression.yml`](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/blob/e0ee3cb308e2d1956055262870649afb3ff61e9c/.github/workflows/excel-vba-regression.yml)

## Accuracy framework

- [`.github/workflows/accuracy-gate.yml`](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/blob/e0ee3cb308e2d1956055262870649afb3ff61e9c/.github/workflows/accuracy-gate.yml)
- [`benchmark/generate_reference_values.py`](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/blob/e0ee3cb308e2d1956055262870649afb3ff61e9c/benchmark/generate_reference_values.py)
- [`benchmark/M_STATS_PROBDIST_ACCURACYEXPORT.bas`](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/blob/e0ee3cb308e2d1956055262870649afb3ff61e9c/benchmark/M_STATS_PROBDIST_ACCURACYEXPORT.bas)
- [`benchmark/compute_errors.py`](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/blob/e0ee3cb308e2d1956055262870649afb3ff61e9c/benchmark/compute_errors.py)
- [`benchmark/accuracy_contracts.csv`](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/blob/e0ee3cb308e2d1956055262870649afb3ff61e9c/benchmark/accuracy_contracts.csv)
- [`benchmark/probability_accuracy_grid.csv`](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/blob/e0ee3cb308e2d1956055262870649afb3ff61e9c/benchmark/probability_accuracy_grid.csv)
- [`benchmark/accuracy_summary.md`](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/blob/e0ee3cb308e2d1956055262870649afb3ff61e9c/benchmark/accuracy_summary.md)
- [`benchmark/numerical_limitations.csv`](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/blob/e0ee3cb308e2d1956055262870649afb3ff61e9c/benchmark/numerical_limitations.csv)
- [`benchmark/render_contract_table.py`](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/blob/e0ee3cb308e2d1956055262870649afb3ff61e9c/benchmark/render_contract_table.py)
- [`benchmark/README.md`](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/blob/e0ee3cb308e2d1956055262870649afb3ff61e9c/benchmark/README.md)

## Repository documentation

- [`README.md`](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/blob/e0ee3cb308e2d1956055262870649afb3ff61e9c/README.md)
- [`docs/EXCEL_VBA_CI.md`](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/blob/e0ee3cb308e2d1956055262870649afb3ff61e9c/docs/EXCEL_VBA_CI.md)

---

# Appendix B — Scoring rationale by category

## Functional correctness — 8.7

Positive factors:

- broad correct family coverage;
- explicit support behavior;
- direct tails;
- exact discrete inverse semantics;
- guarded overflow;
- extensive committed evidence.

Deduction:

- confirmed silent `LogBeta` defect in an accepted shape regime.

## Numerical robustness — 9.1

Positive factors:

- strong tail orientation;
- log-domain formulas;
- paired arguments;
- Loader mass kernels;
- F accuracy envelope;
- exact-integer policy.

Deduction:

- one kernel dispatched outside its validated conditioning regime.

## Architecture and modularity — 9.5

Positive factors:

- clear one-way dependencies;
- project-private core and special functions;
- public wrapper separation;
- consolidated tests.

Deduction:

- Discrete and Test modules are reaching a practical reviewability boundary.

## Public API design — 9.4

Positive factors:

- consistent naming;
- consistent `Variant`/`CVErr` behavior;
- direct survival;
- useful LogPMFs;
- documented parameterization.

Deduction:

- no machine-readable distinction between accuracy-contracted and computed-but-uncontracted success.

## Error handling and diagnostics — 9.3

Positive factors:

- deliberate `#NUM!` versus `#VALUE!`;
- optional Status;
- no modal UI;
- valid-underflow policy.

Deduction:

- release-gate evidence errors are not treated as strictly as production numerical errors.

## Regression testing — 9.3

Positive factors:

- 96 registered test sections;
- direct kernel tests;
- error contracts;
- complete discrete registration;
- named regressions.

Deduction:

- no tiny/tiny unbalanced `LogBeta` coverage.

## External accuracy assurance — 8.7

Positive factors:

- 132 frozen contracts;
- 1,471 observations;
- regime-aware metrics;
- holdouts;
- dedicated studies;
- limitations register.

Deductions:

- incomplete `ERROR`/tail-row strictness;
- evidence not bound to source.

## CI and release engineering — 8.1

Positive factors:

- real Excel execution;
- isolated workbook;
- complete module import;
- machine-readable counters;
- same-repository PR guard;
- hosted strict analyzer.

Deductions:

- source/evidence provenance gap;
- current required-check enforcement not independently evidenced.

## Documentation — 8.0

Positive factors:

- extensive procedure-level comments;
- numerical rationale;
- public API and parameterization documentation;
- generated contract table capability.

Deductions:

- materially stale accuracy claims and generator descriptions.

## Maintainability and hygiene — 8.7

Positive factors:

- house style;
- explicit ownership;
- comments and provenance;
- machine-readable contracts.

Deductions:

- very large modules;
- missing static API/registration/generated-file checks.

## Performance engineering — 8.3

Positive factors:

- efficient algorithm selection;
- direct kernels;
- avoidance of WorksheetFunction marshalling;
- bounded search and summation policies.

Deduction:

- no reproducible timing evidence.
