# Detailed Code Review — VBA Probability Distributions

> **Repository:** `danielep71/VBA-PROBABILITY-DISTRIBUTIONS`  
> **Branch reviewed:** `main`  
> **Commit reviewed:** [`d943500b39945f243db96f8a752e76277930d732`](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/commit/d943500b39945f243db96f8a752e76277930d732)  
> **Review date:** 18 July 2026  
> **Review type:** Static source, numerical-design, testing, CI, benchmark, and documentation review  
> **Recommended repository path:** `docs/CODE_REVIEW_2026-07-18.md`

---

## 1. Review basis and limitations

A direct `git clone` was attempted, but the execution environment could not resolve `github.com`. The current `main` branch was therefore inspected through the authenticated GitHub connector, including the latest commit history, current committed files, and the full change set since the previous reviewed baseline.

The review covers:

- [`src/M_STATS_PROBDIST_CORE.bas`](../src/M_STATS_PROBDIST_CORE.bas)
- [`src/M_STATS_PROBDIST_SPECIALFUNCS.bas`](../src/M_STATS_PROBDIST_SPECIALFUNCS.bas)
- [`src/M_STATS_PROBDIST_NORMALFAMILY.bas`](../src/M_STATS_PROBDIST_NORMALFAMILY.bas)
- [`src/M_STATS_PROBDIST_TFAMILY.bas`](../src/M_STATS_PROBDIST_TFAMILY.bas)
- [`src/M_STATS_PROBDIST_CONTINUOUS.bas`](../src/M_STATS_PROBDIST_CONTINUOUS.bas)
- [`tests/M_STATS_PROBDIST_TEST.bas`](../tests/M_STATS_PROBDIST_TEST.bas)
- [`ci/Run-ExcelVbaTests.ps1`](../ci/Run-ExcelVbaTests.ps1)
- [`.github/workflows/excel-vba-regression.yml`](../.github/workflows/excel-vba-regression.yml)
- [`benchmark/`](../benchmark/)
- [`README.md`](../README.md)
- related repository documentation and policy files.

This report does **not** claim that the current Excel regression suite was executed during the review. The conclusions distinguish among:

- confirmed behavior supported by committed benchmark data;
- defects visible from the implementation;
- design risks requiring additional execution evidence;
- documentation or release-engineering inconsistencies.

---

# 2. Executive assessment

## Overall repository score: **8.9 / 10**

## Production numerical library score: **9.0 / 10**

The repository is now a serious numerical-computing project rather than a conventional VBA utility collection. Its strongest characteristics are:

- clear numerical layering;
- direct tail and inverse-tail APIs;
- disciplined `CVErr` contracts;
- cancellation-resistant elementary primitives;
- reusable incomplete-beta and incomplete-gamma kernels;
- extensive deterministic regression testing;
- a credible Excel/PowerShell CI architecture;
- an increasingly reproducible high-precision benchmark process;
- exceptional source-level documentation.

Several earlier weaknesses have been corrected:

- Normal and Lognormal density reconstruction now uses the logarithmic domain;
- predictable location-scale overflow is guarded;
- the `1E100` constant is now described as a coarse parameter-magnitude guard rather than a convergence boundary;
- `.gitattributes` has been moved to the repository root;
- CI failure names are now retained and surfaced;
- the accuracy harness has expanded across the public distribution families;
- the `PROB_LogBeta` imbalance problem has been measured rather than guessed.

The principal remaining issue is now well established:

> `PROB_LogBeta` has a confirmed accuracy gap for moderately unbalanced non-special arguments. The general three-log-gamma identity cancels, while the existing one-term asymptotic is too crude until the arguments become extremely unbalanced. No threshold change alone can meet the stated accuracy target across the middle band.

This is a **P1 numerical correctness issue** for affected Beta and asymmetric F calculations. The current source and benchmark documentation acknowledge the limitation, which is good governance, but documentation does not remove the defect.

The benchmark infrastructure also contains a material verdict-policy inconsistency: after introducing a two-part `hi;lo` representation intended to recover the full `Double`, the analyzer still suppresses some measured accuracy misses as “below harness precision.” This weakens the benchmark as an enforceable release gate.

---

# 3. Detailed scorecard

| Area | Weight | Score | Assessment |
|---|---:|---:|---|
| Functional correctness | 18% | **8.8** | Broadly strong; confirmed `PROB_LogBeta` middle-band error remains user-reachable. |
| Numerical robustness | 18% | **8.6** | Excellent tail, log-domain, and guarded-arithmetic design; LogBeta and limiting-standardization behavior remain incomplete. |
| Architecture and modularity | 11% | **9.6** | Excellent separation of Core, Special Functions, public distribution families, tests, CI, and benchmarks. |
| Public API design | 9% | **9.3** | Consistent naming, direct survival/inverse-survival functions, explicit parameterization, worksheet-safe returns. |
| Error handling and diagnostics | 8% | **9.1** | Strong predictable-versus-unexpected error policy; some limiting distribution cases are still classified as errors. |
| Regression testing | 11% | **9.4** | Exceptional deterministic suite for VBA, including named regressions, deep tails, exact error codes, and full-range cases. |
| CI and release engineering | 7% | **8.7** | Strong self-hosted Excel design and improved failure logging; no successful current-run evidence was available during review. |
| Accuracy benchmark framework | 8% | **8.2** | Major advance, but verdict-floor logic, small calibration grids, and claim derivation need tightening. |
| Documentation | 5% | **8.7** | Rich and unusually transparent; several accuracy tables and study documents are inconsistent or misplaced. |
| Maintainability | 4% | **9.0** | Disciplined house style and strong contracts; high duplication creates synchronization risk. |
| Performance engineering | 1% | **8.6** | Appropriate algorithms and fast paths, but no comprehensive reproducible timing baseline. |

### Weighted result

```text
Overall score: 8.9 / 10
```

### Score interpretation

| Score | Interpretation |
|---:|---|
| 9.5–10.0 | Exceptional, independently reproducible, mature release-grade library |
| 9.0–9.4 | Advanced and highly credible, with limited remaining gaps |
| 8.0–8.9 | Strong professional project with one or more material unresolved issues |
| 7.0–7.9 | Good implementation requiring substantial hardening |
| Below 7.0 | Significant correctness, architecture, or governance weaknesses |

The repository is at the top end of the **strong professional project** category. Resolving the LogBeta regime and making benchmark verdicts enforceable would move it clearly above 9.

---

# 4. Component-level scores

| Component | Score | Summary |
|---|---:|---|
| `M_STATS_PROBDIST_CORE` | **9.4** | Excellent shared numerical contract, guarded arithmetic, and explicit magnitude semantics. |
| `M_STATS_PROBDIST_SPECIALFUNCS` | **8.6** | Sophisticated kernel layer; `PROB_LogBeta` is the major remaining production defect. |
| `M_STATS_PROBDIST_NORMALFAMILY` | **9.3** | Excellent API and improved density reconstruction; a few limiting-value semantics remain debatable. |
| `M_STATS_PROBDIST_TFAMILY` | **9.1** | Strong Student t, Chi-square, and F algorithms; asymmetric F inherits LogBeta risk. |
| `M_STATS_PROBDIST_CONTINUOUS` | **9.0** | Broad and well designed; Beta public paths inherit the LogBeta accuracy gap. |
| `M_STATS_PROBDIST_TEST` | **9.4** | Outstanding deterministic regression harness and failure contract. |
| Excel/PowerShell CI | **8.7** | Strong architecture and security posture; operational evidence and richer failure detail remain desirable. |
| Accuracy benchmark framework | **8.2** | Correct overall architecture, but acceptance logic and coverage methodology need refinement. |
| README and technical documentation | **8.7** | Excellent presentation and transparency, with current consistency defects. |

---

# 5. Major improvements since the previous baseline

The current branch is materially stronger than the previously reviewed baseline.

## 5.1 Normal and Lognormal densities now use log-domain reconstruction

The earlier direct forms:

```vb
NormalPDF(Z) / StdDev
```

and:

```vb
NormalPDF(Z) / (X * StdDevLog)
```

could independently underflow numerator and denominator even when the final ratio was representable.

The current implementation reconstructs the logarithm of the density and then calls `PROB_TryExp`. This is the correct numerical architecture:

```text
Normal:
    log f(x) = -0.5 z² - log(σ) - 0.5 log(2π)

Lognormal:
    log f(x) = -0.5 z² - log(x) - log(σlog) - 0.5 log(2π)
```

The benchmark now includes a deep-underflow Lognormal density regression where the old expression returned zero or faulted but the true density remained representable.

**Assessment:** important and correct production hardening.

## 5.2 Guarded affine and standardization primitives are now shared

Core now includes guarded helpers for:

```text
Offset + Scale × Value
(Value - Location) / Scale
```

These centralize predictable overflow behavior and prevent repeated ad hoc arithmetic in public wrappers.

**Assessment:** excellent architectural improvement.

## 5.3 The `1E100` constant has been reframed honestly

The constant is now named and documented as a parameter-magnitude guard, not as a mathematical, accuracy, or convergence boundary.

That is the right policy. Hard per-kernel cutoffs would create a false sharp boundary where the real behavior is parameter-dependent and fuzzy. The current approach appropriately:

- rejects only extremely broad regimes outside the declared contract;
- attempts eligible cases;
- relies on iterative Try-contracts to report genuine non-convergence;
- avoids returning partial sums.

**Assessment:** the documentation now reflects the actual numerical contract.

## 5.4 CI now records failed assertion names

The test module accumulates failures in `mFailureLog`, and the PowerShell runner retrieves and writes them to the CI artifact.

**Assessment:** meaningful improvement over count-only failure reporting.

## 5.5 The accuracy harness now covers the wider public surface

The benchmark framework now spans special functions, T-family distributions, the Normal/Lognormal family, and the continuous distributions.

The two-part `hi;lo` observation format is a thoughtful attempt to preserve the full VBA `Double` rather than relying on one approximately 15-digit formatted literal.

**Assessment:** major governance improvement, although the verdict policy must now be aligned with that stronger representation.

## 5.6 The LogBeta problem was benchmarked before changing production code

This is exemplary numerical-development discipline.

The study prevented an incorrect change: widening the existing one-term asymptotic switch would have produced large errors for non-integer small arguments. The benchmark established that the problem cannot be solved by threshold movement alone.

**Assessment:** one of the strongest aspects of the current project process.

---

# 6. Findings summary

| ID | Severity | Area | Finding |
|---|---|---|---|
| P1-01 | P1 | Numerical correctness | Confirmed `PROB_LogBeta` accuracy gap for moderately unbalanced arguments |
| P1-02 | P1 | Benchmark governance | “Below harness precision” can mask measured misses despite full-Double `hi;lo` reconstruction |
| P2-01 | P2 | Benchmark methodology | Several acceptance bounds are derived from the same small grid used to demonstrate compliance |
| P2-02 | P2 | Public semantics | Standardization overflow is treated uniformly as `#NUM!` instead of returning distribution limits where sign is known |
| P2-03 | P2 | CI diagnostics | Failure-log documentation promises actual/expected/tolerance, but the buffer currently stores only the assertion name |
| P2-04 | P2 | CI operations | Workflow exists, but no successful current workflow result was available as review evidence |
| P2-05 | P2 | Documentation | `benchmark/logbeta_study/README.md` appears to contain the general benchmark README rather than the study-specific report |
| P2-06 | P2 | Documentation | Benchmark bound tables are inconsistent with the generated grid and summary |
| P2-07 | P2 | Documentation | Student t exposure to the LogBeta gap is described too broadly |
| P2-08 | P2 | Testing | The confirmed LogBeta gap is characterized in a study but not integrated into the authoritative regression/release verdict |
| P3-01 | P3 | Maintainability | Source headers, README, Wiki, benchmark claims, and generated summaries duplicate contract information |
| P3-02 | P3 | Benchmark tooling | Fixed inverse-reference brackets will limit future extreme-probability grid expansion |
| P3-03 | P3 | Documentation | Several procedure and module `UPDATED` fields do not reflect recent material changes |

---

# 7. Detailed findings

## P1-01 — Confirmed `PROB_LogBeta` accuracy gap for moderately unbalanced arguments

### Status

**Confirmed production defect.**

### Current implementation

`PROB_LogBeta` selects among:

1. special half-integer routes;
2. a one-term unbalanced approximation:

```text
LogGamma(Small) - Small × Log(Large)
```

when:

```text
Small / Large <= 1E-15
```

3. the defining identity:

```text
LogGamma(A) + LogGamma(B) - LogGamma(A + B)
```

otherwise.

### Measured behavior

The committed VBA study establishes that:

- the general identity degrades through cancellation as imbalance increases;
- the one-term asymptotic is accurate only in the extremely unbalanced regime;
- the one-term asymptotic has material truncation error at moderate ratios;
- no threshold can select an accurate answer throughout the middle band.

Representative best-of-current-method errors include approximately:

| `Small / Large` | Best achievable from current two methods |
|---:|---:|
| `1E-1` | `8.9E-16` |
| `1E-2` | `1.4E-14` |
| `1E-3` | `2.9E-13` |
| `1E-7` | `5.2E-10` |
| `1E-10` | `1.9E-12` |
| `1E-13` | `1.3E-15` |

For an accuracy target around `5E-15`, the broad middle band does not meet contract.

### Public impact

The defect is directly relevant to:

- Beta density;
- Beta CDF and survival through incomplete-beta normalization;
- Beta inverse calculations;
- F distribution calculations with disparate degrees of freedom.

The impact on Student t is more limited than the current documentation suggests:

- large-DF Student t normalization commonly reaches the `B = 0.5` shortcut;
- exact low-degree branches avoid the generic path;
- intermediate non-special degrees may still require verification.

### Correct remediation

A threshold change is insufficient.

The missing abstraction is a stable log-gamma increment:

```vb
PROB_LogGammaDelta(LargeArg, Increment)
```

computing:

```text
LogGamma(LargeArg + Increment) - LogGamma(LargeArg)
```

without subtracting two large independently rounded values.

Recommended three-regime structure:

```text
Balanced arguments
    → direct LogGamma identity

Moderately unbalanced arguments
    → cancellation-free Lanczos log-gamma difference

Strongly unbalanced arguments
    → validated multi-term digamma/Bernoulli expansion
```

The user's measured four-term expansion reaching approximately `1E-16` at ratios at or below about `1E-4` is strong evidence for the third regime. The residual region near `1E-2` to `1E-4` should be handled by the stable difference kernel or a sufficiently validated higher-order expansion.

### Required tests before merge

- direct `PROB_LogBeta` grid across ratios `1E0` to `1E-18`;
- several non-integer `Small` values;
- multiple absolute `Large` scales;
- switch-seam continuity;
- public Beta density/CDF/survival cases;
- highly asymmetric F cases;
- symmetry `LogBeta(A,B) = LogBeta(B,A)`;
- comparison among direct, stable-difference, and asymptotic paths.

### Recommendation

Treat this as the next production release gate. Do not add additional distributions before resolving it.

---

## P1-02 — Benchmark verdict policy can mask measured accuracy misses

### Status

**Confirmed benchmark-governance inconsistency.**

### Current design

The benchmark documentation states that VBA observations are exported as:

```text
hi;lo
```

and reconstructed by summing two components, so the analyzer can recover the full `Double` beyond a one-literal 15-digit formatting floor.

That is a meaningful improvement.

### Current inconsistency

The generated summary still marks some rows as:

```text
⚠️ below harness precision
```

when the measured error exceeds the stated claim, including examples such as:

- Standard Normal survival around `1.5E-14` against a `5E-15` claim;
- deep-tail Lognormal density around `2.36E-14` against a `5E-15` claim.

If the `hi;lo` scheme truly reconstructs the full observed `Double`, a blanket `1E-14` “measurement floor” is no longer logically consistent. A measured miss should be handled in one of three ways:

1. the implementation fails the claim;
2. the claim is explicitly relaxed;
3. the benchmark quantifies a specific uncertainty interval showing the miss is not distinguishable.

The current “below harness precision” classification acts as a non-failing exception without a quantified uncertainty model.

### Risk

The accuracy report can state that all functions pass while containing rows above their declared thresholds.

That weakens the report as a release gate and can create false confidence.

### Recommendation

Remove the blanket floor for `hi;lo` observations.

Use exact or near-exact observed-value reconstruction and calculate error with `Decimal` or `mpmath.mpf`.

A verdict should be:

```text
PASS
FAIL
PENDING / UNMEASURED
```

“Below harness precision” should be reserved for a formally demonstrated measurement limitation, not inferred only because the claim is tight.

For subnormal values, use an ULP-aware or absolute-error criterion in addition to relative error. The deep Lognormal density case may be numerically excellent in absolute/ULP terms while missing a relative target because of subnormal spacing. That should be stated explicitly rather than suppressed by a generic floor.

---

## P2-01 — Acceptance bounds are partly calibrated from the same grid used to pass them

### Current documentation

The benchmark README states that several continuous-distribution bounds were set from the measured worst-case error over the tested grid.

### Problem

This mixes:

- characterization data;
- acceptance criteria;
- validation data.

A threshold selected after seeing the maximum error on the same small grid is not an independent accuracy contract.

For example, if the worst measured value is `1.70E-14` and the bound is set to `2E-14`, the table demonstrates only that the selected points fit the selected bound.

### Recommendation

Separate the process:

1. **Development grid**  
   Used to design algorithms and set candidate tolerances.

2. **Frozen acceptance contract**  
   Versioned thresholds justified from development evidence.

3. **Independent validation grid**  
   Different points, denser seams, extreme values, and random reproducible cases.

4. **Regression grid**  
   Permanent named cases for every discovered defect.

Do not call a data-derived bound “published accuracy” unless the relevant source comments actually publish it.

---

## P2-02 — Standardization overflow is always an error, even when a distribution limit is known

### Current behavior

`PROB_TryStandardize` returns Boolean success/failure. When:

```text
(X - Mean) / StdDev
```

overflows, the sign is lost and public wrappers generally return `#NUM!`.

### Mathematical alternatives

For distribution functions, the sign of an infinite limiting z-score determines an exact result:

| Function | Positive standardized overflow | Negative standardized overflow |
|---|---:|---:|
| Normal density | `0` | `0` |
| Normal CDF | `1` | `0` |
| Normal survival | `0` | `1` |
| Z-score | `#NUM!` | `#NUM!` |

### Recommendation

Return an overflow classification or sign:

```vb
Enum PROB_StandardizeOutcome
    PROB_StandardizeInvalid = 0
    PROB_StandardizeSuccess = 1
    PROB_StandardizePositiveOverflow = 2
    PROB_StandardizeNegativeOverflow = 3
End Enum
```

This is not required to resolve an immediate wrong finite value, but it would make the public probability API more mathematically complete.

---

## P2-03 — CI failure detail is less rich than its documentation claims

### Current implementation

`mFailureLog` appends:

```text
TestName
```

The CI bridge retrieves that text and writes it to the result artifact.

### Documentation claim

The injected function's comment says the runner can surface:

```text
assertion name, actual, expected and tolerance
```

### Mismatch

The buffer contains only the assertion name unless the assertion helper embeds all detail into `TestName`.

### Recommendation

Store a structured failure line in each assertion helper:

```text
name=<...>;actual=<...>;expected=<...>;abs_error=<...>;rel_error=<...>;tolerance=<...>
```

At minimum, append the exact line already printed to the Immediate window.

---

## P2-04 — Excel CI exists, but operational evidence is still incomplete

### Strengths

The workflow:

- uses a self-hosted Windows/Excel runner;
- excludes untrusted fork code;
- creates an isolated workbook;
- imports the current modules;
- injects a CI-only result bridge;
- validates assertion counters;
- returns non-zero on failure;
- uploads a result artifact;
- cleans up COM references.

This is strong design.

### Remaining evidence gap

No successful current workflow run was available during this review.

### Recommendation

Record and link:

- first successful run;
- runner Excel version/build;
- assertion total;
- artifact retention;
- expected status-check name for branch protection.

Then make the workflow a required branch-protection check.

---

## P2-05 — `benchmark/logbeta_study/README.md` appears overwritten

The latest commit changes `benchmark/logbeta_study/README.md` from a study-specific document into content headed “Accuracy benchmarks,” substantially duplicating the parent benchmark README.

This loses the clearest place for:

- study purpose;
- exact ratio grid;
- measured table;
- tested small values;
- branch and general-method curves;
- conclusions invalidating a threshold-only fix;
- proposed multi-term/stable-difference remediation.

### Recommendation

Restore a dedicated LogBeta study report. It should include:

- the original study design;
- actual measured VBA results;
- integer versus non-integer Small behavior;
- best-of-method envelope;
- confirmed middle-band failure;
- current production status;
- proposed remediation and acceptance criteria.

---

## P2-06 — Benchmark documentation contains inconsistent bounds

Examples:

- the generated grid/summary gives `Beta_Cumulative` a `2E-14` bound;
- the benchmark README groups Beta CDF with functions described as `5E-15`;
- the summary reports a measured Beta CDF error above `5E-15` but below `2E-14`;
- the general text says bounds are taken verbatim from source comments, while the continuous-module bounds are explicitly derived from measurements because the source publishes none.

### Recommendation

Generate the benchmark README table from one authoritative machine-readable contract file, for example:

```yaml
Beta_Cumulative:
  metric: relative
  threshold: 2e-14
  provenance: benchmark-contract
  domain: balanced-to-moderately-unbalanced
```

Do not duplicate thresholds manually across:

- source comments;
- Python grid generator;
- benchmark README;
- generated summary;
- main README;
- Wiki.

---

## P2-07 — Student t exposure to the LogBeta limitation is described too broadly

The documentation currently states that the LogBeta middle-band issue propagates generally to Student t.

That requires qualification.

Student t includes protective paths:

- exact closed forms at selected low degrees of freedom;
- a `Beta(a, 0.5)` normalization where the special half-integer path applies for much of the large-DF regime;
- direct central/tail arrangements.

### Recommendation

State:

> The defect directly affects arbitrary Beta and asymmetric F calculations. Student t should be tested across non-special real degrees of freedom, but several important Student t regimes use exact or half-integer-specialized paths and are not automatically exposed to the generic LogBeta failure.

---

## P2-08 — The LogBeta study is not part of the authoritative release verdict

The study has successfully identified a production limitation, but it is separate from:

```vb
Test_STATS_PROBDIST_RunAll
```

and from the main benchmark summary.

### Risk

A release can report:

```text
RESULT: ALL TESTS PASSED
```

while the known LogBeta middle-band defect remains unchanged.

### Recommendation

Until the algorithm is fixed, add one explicit expected-limitation test or release-gate check that prevents a blanket “all numerical contracts passed” interpretation.

After the fix, move representative seam and public-path cases into:

- `M_STATS_PROBDIST_TEST`;
- the main external accuracy grid;
- the generated accuracy summary.

---

# 8. Module-by-module review

## 8.1 `M_STATS_PROBDIST_CORE` — **9.4 / 10**

### Strengths

- true finiteness separated from parameter-magnitude policy;
- guarded add, multiply, divide, exponentiation, affine transform, and standardization;
- explicit valid-underflow policy;
- reusable `Log1p` and `Expm1`;
- carefully represented constants;
- project-visible but worksheet-hidden surface through `Option Private Module`;
- honest documentation of the coarse magnitude guard.

### Improvements recommended

- return sign/classification from standardization overflow;
- rename legacy predicates such as `PROB_IsWithinSupportedMagnitude` in a later compatibility-conscious pass, since the constant is now a “guard” but the predicate still says “supported”;
- ensure all recent additions have current `UPDATED` metadata.

### Verdict

Excellent numerical infrastructure. No major architectural redesign is needed.

---

## 8.2 `M_STATS_PROBDIST_SPECIALFUNCS` — **8.6 / 10**

### Strengths

- clear Boolean Try-contract;
- no partial answers after non-convergence;
- paired incomplete-beta arguments;
- direct smaller-tail evaluation;
- incomplete-gamma series/continued-fraction split;
- safeguarded inverse solvers;
- stable half-log-gamma difference;
- Loader-style Stirling correction and LogChoose support;
- unusually transparent provenance.

### Principal weakness

`PROB_LogBeta` does not provide stable accuracy in the moderately unbalanced regime.

### Improvements recommended

- implement stable `PROB_LogGammaDelta`;
- retain half-integer shortcuts;
- use direct identity only where benchmarked as well-conditioned;
- use multi-term asymptotic only where validated;
- include iteration count and final residual in failure messages;
- integrate LogBeta seam tests into the main benchmark.

### Verdict

The module remains sophisticated, but one central normalization kernel prevents a higher score.

---

## 8.3 `M_STATS_PROBDIST_NORMALFAMILY` — **9.3 / 10**

### Strengths

- complete Standard Normal, Normal, and Lognormal API;
- direct survival and inverse-survival functions;
- stable interval probabilities;
- deep-tail Normal evaluation;
- raw fast inverse helper for Monte Carlo use;
- guarded affine and standardization paths;
- log-domain Normal and Lognormal densities;
- explicit support behavior at `X <= 0` for Lognormal functions.

### Remaining issues

- limiting standardized overflow is returned as an error rather than saturated probability;
- the deep subnormal Lognormal density should use an ULP/absolute benchmark metric rather than a strict ordinary relative metric;
- general Normal retains a conservative finite-magnitude restriction even though arithmetic is now better guarded.

### Verdict

One of the strongest modules in the project after the latest fixes.

---

## 8.4 `M_STATS_PROBDIST_TFAMILY` — **9.1 / 10**

### Strengths

- stable logarithmic Student t density;
- direct Student t tails;
- exact low-degree formulas;
- safeguarded inverse iteration;
- direct Gamma P/Q use for Chi-square;
- stable F logistic argument pair;
- log-domain F inverse reconstruction;
- explicit `#NUM!` failure contracts.

### Remaining issues

- asymmetric F calculations inherit LogBeta normalization risk;
- benchmark grids should include much more disparate `DF1/DF2`;
- Student t documentation should distinguish protected half-integer/exact paths from generic Beta exposure.

### Verdict

Highly mature. Most remaining risk is inherited from Special Functions.

---

## 8.5 `M_STATS_PROBDIST_CONTINUOUS` — **9.0 / 10**

### Strengths

- clear parameterization;
- direct Gamma and Beta tails;
- log-domain Gamma density;
- stable Exponential/Weibull CDFs using `Expm1`;
- guarded inverse reconstruction;
- stable Weibull moments;
- full finite-range Uniform formulas;
- exact support-edge policies;
- distinction between shape guards and full-range scale/rate/bound inputs.

### Remaining issues

- Beta density/CDF/survival/inverse inherit the LogBeta middle-band defect;
- the main deterministic suite does not yet include representative public unbalanced-Beta failures;
- continuous-distribution benchmark thresholds are not yet independently validated.

### Verdict

Excellent module design with one inherited special-function limitation.

---

## 8.6 `M_STATS_PROBDIST_TEST` — **9.4 / 10**

### Strengths

- consolidated authoritative result;
- direct Core and Special Function testing;
- exact constant verification;
- complement and symmetry identities;
- inverse round-trips;
- deep-tail tests;
- support-edge cases;
- full-range Uniform tests;
- exact `#NUM!` versus `#VALUE!` assertions;
- regression registry;
- CI failure buffer.

### Improvements recommended

- store full assertion diagnostics in `mFailureLog`;
- add fixed-seed generated monotonicity grids;
- add LogBeta middle-band cases;
- add public Beta/F cases that intentionally exercise the problematic normalization regime;
- separate expected known limitations from green release criteria.

### Verdict

Exceptional by VBA standards and strong by general numerical-library standards.

---

## 8.7 Excel/PowerShell CI — **8.7 / 10**

### Strengths

- correct use of a self-hosted runner with desktop Excel;
- isolated temporary workbook;
- source modules imported in dependency order;
- CI-only bridge avoids production-test API pollution;
- fork security control;
- counter consistency checks;
- COM cleanup;
- artifact upload;
- improved failure-name output.

### Improvements recommended

- demonstrate and link a successful run;
- add actual/expected/tolerance detail;
- make the check required for `main`;
- write Excel version/build and commit SHA into the artifact;
- optionally preserve the temporary workbook on failure for forensic review;
- add explicit compile-stage logging where feasible.

---

## 8.8 Accuracy benchmark framework — **8.2 / 10**

### Strengths

- independent high-precision reference generation;
- separation of reference and observed paths;
- direct survival references;
- two-part observed serialization;
- broad public-function coverage;
- committed environment metadata;
- generated summary;
- dedicated LogBeta study.

### Improvements recommended

- remove or rigorously justify the blanket precision floor;
- fail the process on genuine misses;
- use exact high-precision arithmetic throughout analysis;
- separate development and validation grids;
- expand parameter grids substantially;
- use ULP-aware metrics near subnormal values;
- integrate benchmark execution into CI;
- restore the study-specific README;
- centralize accuracy-contract metadata.

---

# 9. Documentation and repository review

## Strengths

The project has an unusually strong documentation stack:

- premium repository README;
- module-level architecture comments;
- procedure contracts;
- numerical provenance;
- parameterization;
- error policy;
- benchmark methodology;
- CI setup instructions;
- contributing and security files.

The project is transparent about numerical limitations, which materially increases trust.

## Current documentation defects

1. The LogBeta study README appears overwritten.
2. Beta CDF accuracy bounds are inconsistent across files.
3. “Taken verbatim from source comments” is not true for measurement-derived continuous bounds.
4. Student t exposure to LogBeta is overgeneralized.
5. CI documentation should distinguish “workflow present” from “runner verified.”
6. Multiple recent source changes retain older `UPDATED` dates.

## Recommendation

Create one machine-readable API and numerical-contract manifest and generate:

- README capability tables;
- benchmark threshold tables;
- Wiki API index;
- public-function counts;
- documentation cross-checks.

---

# 10. Recommended release gates

## Release Gate A — Resolve LogBeta

- implement `PROB_LogGammaDelta`;
- validate direct/stable/asymptotic regime selection;
- add multi-term expansion;
- close the `1E-2`–`1E-4` residual region;
- add public Beta and asymmetric F regressions;
- update source accuracy claims only after independent validation.

## Release Gate B — Make benchmark verdicts enforceable

- remove blanket “below precision” suppression for `hi;lo` observations;
- use exact Decimal/mpmath analysis;
- add ULP-aware subnormal metrics;
- return non-zero on measured claim failure;
- separate characterization and validation grids;
- regenerate an internally consistent summary.

## Release Gate C — Operationalize CI

- bring the self-hosted runner online;
- record a successful run;
- make the status check required;
- include full failure diagnostics;
- preserve forensic artifacts on failure.

## Release Gate D — Documentation normalization

- restore `benchmark/logbeta_study/README.md`;
- reconcile Beta CDF bounds;
- qualify Student t exposure;
- centralize contract metadata;
- refresh module dates and dependency lists.

## Release Gate E — Generated/property testing

After the deterministic and benchmark gates are stable:

- monotonicity;
- CDF/survival complement;
- inverse round-trips;
- distribution identities;
- support invariants;
- fixed-seed random parameter grids;
- persisted failing seeds.

---

# 11. Suggested priority order

| Priority | Action | Reason |
|---:|---|---|
| 1 | Fix `PROB_LogBeta` middle-band accuracy | Confirmed P1 production defect |
| 2 | Correct benchmark verdict-floor logic | Required for trustworthy release claims |
| 3 | Add LogBeta public-path regressions | Prevent recurrence and prove remediation |
| 4 | Verify and require Excel CI | Converts testing from process guidance into enforcement |
| 5 | Normalize benchmark documentation | Current tables conflict |
| 6 | Improve limiting standardized-overflow behavior | Mathematical API completeness |
| 7 | Add generated/property tests | Broader assurance after core contracts stabilize |

---

# 12. Final verdict

This repository is one of the strongest pure-VBA numerical projects likely to be encountered in an open-source setting.

It demonstrates:

- serious understanding of floating-point failure modes;
- strong separation between numerical kernels and worksheet wrappers;
- unusual attention to tails and inverse functions;
- explicit and consistent error semantics;
- regression-oriented development;
- willingness to benchmark hypotheses before changing algorithms;
- transparent documentation of known limitations.

The latest development cycle is particularly encouraging: the LogBeta investigation did not merely confirm a preconceived fix. It disproved a tempting threshold-only change and exposed the need for a better kernel. That is exactly how numerical software should be developed.

The project is not yet fully release-certified because:

- the LogBeta middle-band defect is confirmed and unresolved;
- the benchmark can still suppress measured misses;
- CI has not yet been evidenced as an active required check;
- some accuracy contracts are calibrated and documented inconsistently.

> **Current classification: 8.9/10 overall; 9.0/10 for the production library. Advanced, credible, and professionally engineered, with one confirmed special-function defect and a benchmark-verdict policy that should be corrected before making global accuracy claims.**

---

## Appendix A — Recommended LogBeta acceptance matrix

A corrected `PROB_LogBeta` implementation should be validated across:

### Ratio grid

```text
1
5E-1
2E-1
1E-1
5E-2
2E-2
1E-2
5E-3
2E-3
1E-3
...
1E-18
```

### Non-integer Small values

```text
0.2
0.7
0.8
1.3
1.5
2.5
5.75
10.25
```

### Absolute Large scales

```text
1E1
1E2
1E4
1E8
1E12
1E20
1E50
```

### Recorded methods

```text
high-precision reference
current direct identity
current one-term asymptotic
candidate multi-term asymptotic
candidate stable Lanczos difference
actual dispatched PROB_LogBeta
```

### Public-path checks

```text
Beta density
Beta CDF
Beta survival
Beta inverse CDF
F CDF
F survival
F inverse CDF
selected non-special Student t cases
```

---

## Appendix B — Recommended benchmark verdict schema

```text
PASS
    measured error <= contractual threshold

FAIL
    measured error > contractual threshold

PENDING
    no observed value or incomplete grid

KNOWN LIMITATION
    explicitly excluded domain documented in the contract

CHARACTERIZATION ONLY
    point is informative but not part of the acceptance contract
```

Avoid using “below harness precision” as a general non-failing category once the observed `Double` is reconstructed exactly enough to evaluate the claim.

---

## Appendix C — Review confidence

| Area | Confidence |
|---|---|
| Architecture and API review | High |
| LogBeta finding | High — supported by committed VBA study and user-provided measured analysis |
| Normal/Lognormal density review | High — implementation and regression changes are visible |
| CI architecture review | High |
| CI operational status | Medium — workflow source reviewed, successful execution not observed |
| Benchmark-policy finding | High |
| Full runtime correctness of all VBA paths | Medium — Excel suite not executed in this review environment |
