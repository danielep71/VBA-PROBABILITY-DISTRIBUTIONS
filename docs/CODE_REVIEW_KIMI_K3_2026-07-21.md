# Independent Code Review — `VBA-PROBABILITY-DISTRIBUTIONS`

| Field | Value |
|---|---|
| Repository | `danielep71/VBA-PROBABILITY-DISTRIBUTIONS` |
| Reviewed commit | `bdc3b928bf548da3b269b657bf959c4c2b55d0a4` (`main`) |
| Review date | 2026-07-21 |
| Reviewer | Kimi K3 (Moonshot AI), independent automated code review |
| Review basis | Cold re-read of the repository contents, static metric extraction, inspection of source/test/benchmark/CI/documentation files, and execution of the reproducible Python accuracy gate |
| Execution boundary | This environment cannot run desktop Excel/VBA. I therefore did **not** execute `Test_STATS_PROBDIST_RunAll` inside Excel. I did run the Python accuracy gate and its degradation self-test successfully. |

> **Overall weighted score: 8.7 / 10** (exact weighted value: **8.71 / 10**).  
> Verdict: a serious, unusually disciplined native-VBA numerical library with strong evidence practices, but with a current assurance gap around the discrete family, CI coverage, LogPMF tests, and several stale or partly misleading documentation/artifact details.

---

## 1. Executive summary

This repository implements a native Excel-VBA probability library rather than wrapping `Application.WorksheetFunction`. The production surface consists of **88 public `K_STATS_*` worksheet/VBA functions** across Normal/Lognormal, Student t / Chi-square / F, Gamma / Beta / Exponential / Weibull / Uniform, and Binomial / Poisson / Geometric. Beneath that surface are **30 project-scoped `PROB_*` kernels** in `CORE` and `SPECIALFUNCS`, plus private family kernels.

The strongest part of the project is its numerical culture. The code repeatedly chooses the right boring solution: direct survival functions instead of `1 - CDF`, paired incomplete-beta arguments `(X, Y)` so complements are not reconstructed by cancellation, Boolean `Try*` kernels that never return partial sums, guarded arithmetic through `PROB_TryAdd/TryMultiply/TryDivide/TryExp`, stable `Log1p`/`Expm1` primitives, Loader-style binomial and Poisson mass arrangements, explicit supported domains, and measured rejection where a kernel can converge locally yet be inaccurate.

The reproducible Python side is also strong. At the reviewed commit:

- `benchmark/compute_errors.py` runs in strict mode and exits `0`.
- The generated summary reports **76 PASS**, **0 FAIL**, **0 KNOWN LIMITATION**, **0 CHARACTERIZATION ONLY**, **0 PENDING**.
- `benchmark/test_gate_degradation.py` exits `0` and confirms that active tail-residual contracts become blocking `PENDING` items when the reference helper is unavailable.
- The contract file has **76 active contracts**: **71 validated and frozen**, **5 measured provisional**.
- The main grid has **724 rows**, of which **716 have non-empty `observed_vba`**.

The review score is not higher because the evidence and automation are not uniformly applied to the current public surface:

1. **All 24 discrete `K_STATS_*` functions are outside the accuracy-contract regime.** There are no Binomial/Poisson/Geometric rows in `accuracy_contracts.csv` or `probability_accuracy_grid.csv`.
2. **The Excel CI path currently omits the discrete module and suite.** `ci/Run-ExcelVbaTests.ps1` imports five production modules plus the test module, but not `src/M_STATS_PROBDIST_DISCRETE.bas`; the injected CI bridge runs Core, NormalFamily, TFamily and Continuous, but not `RunDiscreteSuite`.
3. **The three public `_LogPMF` functions are not referenced by the VBA test module at all.** The test module references **85 of 88** public `K_STATS_*` functions; the missing three are `K_STATS_Binomial_LogPMF`, `K_STATS_Poisson_LogPMF`, and `K_STATS_Geometric_LogPMF`.
4. Several documentation and artifact details are stale or easy to misread: the README tree names an older commit and files that are not present, the test-module header still describes four suites while the code runs five, survival headers still emphasize deep-tail stability without pointing at the registered relative-accuracy limitation, and the benchmark README still says the harness covers 66 functions while the current contract file covers 69 function names.

No blocking numerical-correctness defect was identified in the committed benchmark evidence. The strict gate is green, and the static read found the implementation broadly consistent with its documented algorithms. The deductions are about coverage, evidence scope, automation truthfulness, and documentation precision.

---

## 2. Hard-number snapshot

Numbers below were extracted from the checked-out tree at `bdc3b928`.

| Metric | Current value |
|---|---:|
| Tracked files | 96 |
| `.bas` files | 24 |
| Python files | 24 |
| Markdown files | 22 |
| CSV files | 13 |
| Workflow/config YAML files | 3 |
| PowerShell files | 1 |
| Macro-enabled demo workbook | 1 |
| Production source modules | 6 |
| Production source lines | 19,112 |
| Production source code/comment/blank lines | 6,177 / 11,396 / 1,539 |
| Test module lines | 5,616 |
| Test module code/comment/blank lines | 2,030 / 3,082 / 504 |
| Benchmark `.bas` lines | 3,818 |
| Public `K_STATS_*` functions | 88 |
| Project-scoped `PROB_*` public kernels | 30 |
| Public + private procedures in `src/*.bas` | 181 |
| Test procedures | 94 total: 6 public runners + 88 section procedures |
| Static assertion statement lines in tests | 638 |
| `Option Explicit` in VBA files | 24 / 24 |
| `MsgBox` in executable code | 0 |
| `Application.WorksheetFunction` executable use | 0 |
| Accuracy contracts | 76 active |
| Contract provenance | 71 validated and frozen; 5 measured provisional |
| Main accuracy grid rows | 724 |
| Grid rows with non-empty observed values | 716 |
| Grid rows for Binomial/Poisson/Geometric | 0 |
| Public `K_STATS_*` with accuracy contracts | 64 / 88 |
| Public `K_STATS_*` referenced by tests | 85 / 88 |
| Strict accuracy gate result | PASS, exit 0 |
| Gate degradation self-test | PASS, exit 0 |

Interpretation notes:

- The production comment ratio is about **59.6% of production source lines**. That is unusually high and mostly useful: procedure headers carry purpose, inputs, returns, error policy, dependencies, numerical method, and update dates.
- The static assertion count is a count of assertion statement lines, not an executed assertion total. The actual executed total requires an Excel run, which was outside this environment.
- The repository contains a prior review document at `docs/CODE_REVIEW_FABLE5_2026-07-21.md` scoring **9.1/10** against a different commit (`ecf3e45`). This document is independent and does not reuse that score.

---

## 3. Scoring rubric

Scale: **10** = exceptional for the problem domain; **9** = excellent with only research-level ceilings; **8** = strong professional quality with material gaps; **7** = solid but with gaps a serious user must manage; **6 or below** = deficient.

| # | Category | Weight | Score | Weighted | Main reason |
|---:|---|---:|---:|---:|---|
| 1 | Numerical correctness & methodology | 20% | **9.1** | 1.820 | Correct methods are used and measured; deductions for measured ceilings and uneven evidence depth. |
| 2 | Verification & benchmark evidence | 15% | **8.5** | 1.275 | Strict gate passes and contracts are unusually explicit; discrete family and some grid rows remain outside evidence. |
| 3 | Testing & CI execution | 12% | **7.6** | 0.912 | Broad local harness, but LogPMF untested and the Excel CI path omits the discrete module/suite. |
| 4 | Robustness & error contract | 10% | **8.8** | 0.880 | Consistent `Variant`/`CVErr`, `Try*` guards, explicit domains; policy trade-offs and heavy `GoTo` idiom remain. |
| 5 | API design & Excel integration | 10% | **8.7** | 0.870 | Clear names, Excel-compatible parameterization, direct tail APIs; no vector/array API and incomplete discrete catalogue. |
| 6 | Code quality & maintainability | 10% | **8.7** | 0.870 | Highly consistent house style and single-source constants; large modules and benchmark duplication. |
| 7 | Documentation | 10% | **7.6** | 0.760 | Candid and extensive, but several stale references and header claims need alignment. |
| 8 | Scope & completeness | 8% | **7.2** | 0.576 | Strong continuous coverage; discrete coverage stops at three distributions and lacks evidence parity. |
| 9 | Reproducibility & process | 5% | **7.0** | 0.350 | Public Python gate is reproducible; Excel execution and export remain manual/self-hosted. |
| 10 | Repository hygiene & governance | 5% | **8.0** | 0.400 | License/security/templates/gitignore are good; stale tree references and duplicated study helpers. |
|  | **Overall** | **100%** |  | **8.71** | Rounded: **8.7 / 10** |

---

## 4. What is genuinely strong

### 4.1 The architecture is layered correctly

The dependency direction is clean:

```text
Worksheet/VBA callers
  -> K_STATS_* family modules
    -> SPECIALFUNCS kernels
      -> CORE constants, predicates and guarded arithmetic
```

`CORE` owns shared constants and elementary numerical primitives. `SPECIALFUNCS` owns distribution-independent incomplete-beta/gamma and log-gamma machinery. Family modules own parameterization, validation, support edges, tail orientation, and worksheet error mapping. This is the right separation for a numerical VBA library.

The use of `Option Private Module` in `CORE` and `SPECIALFUNCS` is also correct: kernels remain project-visible to sibling modules but are hidden from the worksheet Function Wizard.

### 4.2 The code avoids the classic Excel/VBA numerical traps

Concrete examples from the cold read:

- `PROB_Log1p` and `PROB_Expm1` use compensated forms rather than naive `Log(1 + X)` and `Exp(X) - 1`, preserving relative accuracy near zero.
- `PROB_TryExp`, `PROB_TryAdd`, `PROB_TryMultiply`, `PROB_TryDivide`, `PROB_TryStandardize`, and `PROB_TryAffineTransform` convert overflow into Boolean failure instead of letting VBA runtime error 6 escape as an unexpected `#VALUE!`.
- `PROB_TryBetaRegularized` takes both `X` and `Y = 1 - X` from callers and never reconstructs the complement internally. Its header correctly explains why this matters for Student t near zero.
- `PROB_TryBetaContinuedFraction` returns `False` on non-convergence and explicitly never returns a partial sum.
- `PROB_LogGammaDelta` forms `LogGamma(LargeArg + Increment) - LogGamma(LargeArg)` as one stable expression; `PROB_LogBeta` dispatches to it in the unbalanced regime instead of subtracting two large log-gammas.
- Binomial and Poisson masses use Loader-style Stirling-error/deviance arrangements. The discrete module also exposes `_LogPMF` entry points for log-domain work.
- The F family enforces a measured envelope: `PROB_F_MAX_DF = 100000#`, and `PROB_F_ValidateEnvelope` rejects larger degrees of freedom for CDF/SF/inverse while deliberately leaving closed-form density unrestricted.

These are not cosmetic choices. They are the difference between a worksheet formula library and a numerical library.

### 4.3 Failure behavior is explicit and Excel-native

The public contract is consistent: worksheet-facing functions return `Variant`, valid results return as `Double`, predictable numerical/domain failures return `CVErr(xlErrNum)`, unexpected runtime failures return `CVErr(xlErrValue)`, valid exponential underflow returns `0`, and no public function raises `MsgBox`.

Static counts support the policy: **0 `MsgBox` calls** in executable source/test code, **482 `CVErr` occurrences** across VBA files, and **0 executable `Application.WorksheetFunction` uses**.

### 4.4 The benchmark gate is real and currently green

I ran:

```text
cd benchmark
python3 compute_errors.py
```

Observed output:

```text
accuracy_summary.md: FAIL=0 KNOWN_LIMITATION=0 CHARACTERIZATION_ONLY=0 PENDING=0 [gate: strict]
  gate passed (exit 0).
```

I also ran:

```text
python3 test_gate_degradation.py
```

Observed result: exit `0`, with the self-test confirming that a missing `_ibeta` helper turns the two active tail-residual contracts into blocking `PENDING` items and makes the gate exit non-zero.

This matters because the release gate is not merely a report generator; it has a tested failure mode.

---

## 5. Findings register

Severity scale: **High** = blocks or materially weakens the assurance story; **Medium** = real gap a maintainer should schedule; **Low** = polish/consistency; **Info** = acceptable trade-off or scope note.

| ID | Severity | Area | Finding | Evidence | Recommendation |
|---|---|---|---|---|---|
| K3-01 | **High** | Excel CI coverage | The Excel CI runner does not import or execute the discrete layer. `ci/Run-ExcelVbaTests.ps1` imports `CORE`, `SPECIALFUNCS`, `NORMALFAMILY`, `TFAMILY`, `CONTINUOUS`, and the test module, but not `src/M_STATS_PROBDIST_DISCRETE.bas`. The injected bridge runs `RunCoreSuite`, `RunNormalFamilySuite`, `RunTFamilySuite`, and `RunContinuousSuite`, but not `RunDiscreteSuite`. | `ci/Run-ExcelVbaTests.ps1` source list and CI bridge; `tests/M_STATS_PROBDIST_TEST.bas` contains `RunDiscreteSuite`. | Add `src\M_STATS_PROBDIST_DISCRETE.bas` to `$sourceFiles` and call `RunDiscreteSuite` in the CI bridge. Until then, do not describe the Excel workflow as running the complete harness. |
| K3-02 | **Medium** | Public API test coverage | The three public `_LogPMF` functions are not referenced by the VBA test module. The test module references **85/88** public `K_STATS_*` functions; the missing three are Binomial, Poisson and Geometric LogPMF. | `grep`/`K_STATS` cross-reference of `src/*.bas` against `tests/M_STATS_PROBDIST_TEST.bas`. | Add direct LogPMF tests: finite log mass where PMF underflows, exact-zero impossible outcomes, support edges, and PMF/LogPMF consistency where PMF is representable. |
| K3-03 | **Medium** | Accuracy evidence | The external accuracy contract/grid regime excludes the entire discrete family. All **24** missing contract-covered `K_STATS_*` functions are Binomial/Poisson/Geometric. | `accuracy_contracts.csv`, `probability_accuracy_grid.csv`, and public API mapping: 64/88 `K_STATS_*` under contract; 0 discrete rows. | Add discrete grid rows and contracts, especially large-`n` Binomial and large-mean Poisson paths through incomplete beta/gamma; then holdout-validate and freeze. |
| K3-04 | **Medium** | Benchmark grid hygiene | The main grid contains 724 rows but only 716 observed values. Eight unobserved rows belong to `Lognormal_Variance` and `Lognormal_StdDev`; the summary still reports PASS with `2/6` points. Four rows use obsolete function names `Lognormal_ParamMeanLog` / `Lognormal_ParamStdDevLog` and are ignored by the contract join. | `probability_accuracy_grid.csv` rows around Lognormal moments; `compute_errors.py` ignores empty observed values and reports measured `n` over matched rows. | Either require full observation for active-contract grid rows in strict mode or mark partially observed contracts explicitly as partial. Remove or regenerate obsolete rows. |
| K3-05 | **Medium** | Documentation truthfulness | Several docs are stale relative to the checked commit: the README repository-structure section says it reflects commit `e77d0a75...`, names `CODE_REVIEW_CHATGPT5.5_2026-07-19.md` and `CODE_REVIEW_FABLE5_2026-07-19.md`, while the actual docs folder contains `CODE_REVIEW_FABLE5_2026-07-21.md` and `EXCEL_VBA_CI.md`. The test-module header still lists four suites and omits discrete, while the code runs five. Benchmark README still says the harness covers 66 functions; current contracts cover 69 function names. | `README.md` repository-structure block; `tests/M_STATS_PROBDIST_TEST.bas` header vs `RunDiscreteSuite`; `benchmark/README.md` vs `accuracy_contracts.csv`. | Regenerate the repository tree and benchmark prose from the current commit. Make stale commit references impossible by deriving them from release tooling or removing them. |
| K3-06 | **Low** | Survival documentation | Survival headers still say the routines stay accurate deep into the tail without pointing to `SurvivalTailRel`. The limitation register is honest and measured, but header-only readers get a stronger impression than the contract file supports. | `K_STATS_NormalStandard_Survival` header; `benchmark/numerical_limitations.csv` and `benchmark/survival_boundary/survival_boundary_summary.md`. | Add one line to each affected survival header: tight relative bounds are domain-restricted; deeper-tail relative degradation is characterized in `SurvivalTailRel`. |
| K3-07 | **Low** | Deep-tail inverse evidence | Inverse-normal contracts are green, but the main grid’s inverse points reach only probability `0.999` / `0.01` style regions, while the README demonstrates `K_STATS_NormalStandard_InverseSurvival(1E-18)`. The private inverse header itself notes tail-quantile relative error is bounded by the tail CDF approximation beyond the split. | `probability_accuracy_grid.csv` inverse-normal rows; `PROB_NormalInvCDF` accuracy note; README quick-start example. | Add extreme inverse-survival grid points or attach an explicit accuracy caveat to the 1E-18 example. |
| K3-08 | **Low** | Benchmark duplication | `_ibeta.py` is identical in three study folders. Several benchmark study folders also contain duplicate `.bas` files with the same `Attribute VB_Name` under different filenames. | SHA-256 comparison of benchmark helpers and `.bas` files. | Single-source `_ibeta.py` where practical; otherwise document that duplicate study macros are intentional export bundles and must never be imported together. |
| K3-09 | **Low** | Gate/report wording | `accuracy_summary.md` reports `KNOWN LIMITATION: 0` while `numerical_limitations.csv` contains two limitation rows. The tally is about contract verdicts, not the limitation register, but a casual reader can misread it. | `accuracy_summary.md` verdict tally; `numerical_limitations.csv`. | Add one sentence under the tally: limitation-register entries are separate from contract verdict states. |
| K3-10 | **Info** | Scope | Negative Binomial, Hypergeometric, Discrete Uniform, multivariate distributions, random variates and array/vector entry points are absent. | README roadmap and public API inventory. | Keep as roadmap. Highest value: complete the discrete family and bring it to evidence parity. |
| K3-11 | **Info** | Policy trade-off | Overflowing standardization returns `#NUM!` even where a mathematical limit value of 0 or 1 would be defensible. | Core guarded-standardization policy and family wrappers. | Acceptable as a documented policy. Revisit only if users report friction. |

---

## 6. Category detail

### 6.1 Numerical correctness & methodology — 9.1/10

The numerical methodology is the strongest part of the repository. The implementation is not a collection of approximations pasted into VBA; it is arranged around known failure modes.

Positive evidence:

- Direct survival functions are exposed across the catalogue, and the incomplete-gamma upper tail is evaluated as `Q` rather than recovered as `1 - P`.
- Incomplete beta receives paired complementary arguments and never forms `1 - X` internally.
- Iterative kernels return Boolean failure and leave results non-contractual on failure.
- The incomplete-beta inverse solves on the smaller tail and returns both `X` and `Y`, avoiding a later destructive complement in callers such as F.
- `PROB_LogGammaDelta` and the unbalanced `PROB_LogBeta` branch remove the catastrophic cancellation in the naive three-log-gamma identity.
- The discrete module uses Loader-style mass arrangements and routes Binomial/Poisson CDFs through the already-shared incomplete-beta/gamma kernels.
- The F envelope rejects a region where local convergence can still be inaccurate, which is exactly how a governed library should behave.

Deductions:

- The measured ceilings are real: unbalanced Beta/F carry looser thresholds than balanced regimes, and the survival tail has a characterized relative-accuracy degradation.
- Deep-tail inverse-normal evidence is narrower than the most visible README example.
- The discrete family uses sound methods but lacks the same external contract evidence as the continuous families.

### 6.2 Verification & benchmark evidence — 8.5/10

The benchmark design is unusually explicit: machine-readable contracts, regime-aware thresholds, provenance states, a generated summary, a two-part `hi;lo` observed-value format to preserve VBA doubles through CSV, Decimal-based error computation, and a gate with tested degradation behavior.

Current state:

- 76 active contracts; all 76 generated summary rows are PASS.
- 71 contracts are validated and frozen; 5 are measured provisional.
- The strict gate exits 0 on the committed grid.
- The degradation self-test exits 0 and proves active tail contracts cannot silently become non-blocking when the evaluator is missing.

Deductions:

- The contract regime covers 69 function names but excludes all discrete distributions.
- The main grid has 8 rows without observed values and 4 obsolete rows; strict mode still passes because contracts with at least some observations pass.
- The Excel export middle remains manual, so the public CI gate verifies committed evidence, not a fresh Excel execution of the current VBA.

### 6.3 Testing & CI execution — 7.6/10

The local VBA harness is substantial: 94 test procedures and 638 static assertion statement lines, with helper types for absolute/relative closeness, error classification, unit-interval checks, and round-trips. The suite order is dependency-aware and the discrete suite exists in `RunAll`.

Deductions are material:

- K3-01: the Excel CI path omits discrete import/execution.
- K3-02: LogPMF functions are public but untested in the VBA harness.
- The suite remains value-centric. There is no randomized/property-based fuzzing over guard edges, signed zero, denormals, or systematic argument-permutation sweeps.
- Because this review environment cannot run desktop Excel, I could not independently confirm the executed assertion count or green VBA run; I inspected the harness and ran the Python-side gate only.

### 6.4 Robustness & error contract — 8.8/10

The error contract is one of the library’s best features. Public functions are worksheet-safe, validation distinguishes invalid domains from numerical failure, kernels do not validate caller preconditions they do not own, and status diagnostics are available through `ByRef Status` without depending on `Application.StatusBar`.

Deductions:

- The code uses `GoTo` heavily for VBA structured flow. This is idiomatic VBA error handling, but it is still a maintainability and audit cost: 449 `GoTo` occurrences across VBA files.
- Some policy choices are defensible but not universally preferred, such as `#NUM!` on overflowing standardization rather than returning a limit value.

### 6.5 API design & Excel integration — 8.7/10

The API is clear and migration-friendly. Names state exactly what they compute, parameterization follows Excel conventions including Excel’s own rate/scale inconsistencies, and the library adds direct survival and inverse-survival pairs that Excel lacks.

Deductions:

- The discrete catalogue stops at Binomial, Poisson and Geometric.
- There are no array/vectorized entry points, which is understandable in VBA but relevant for Monte Carlo workloads.
- LogPMF is present but currently under-protected by tests.

### 6.6 Code quality & maintainability — 8.7/10

The house style is rigorous: `Option Explicit` everywhere, structured banners, explicit declarations, `Double` literals, consistent labels, and header fields that usually match the code. Shared constants are held once, and the Lanczos coefficients are single-sourced for `PROB_LogGamma` and `PROB_LogGammaDelta`.

Deductions:

- Production modules are large: `CONTINUOUS` is 4,953 lines, `DISCRETE` 4,667, `NORMALFAMILY` 3,635, `TFAMILY` 3,364. The style mitigates this, but navigation cost is real.
- Benchmark helper duplication remains, including identical `_ibeta.py` files and duplicate macro filenames with the same VBA module name.

### 6.7 Documentation — 7.6/10

The documentation is extensive and candid about scope, parameterization, validation boundaries, error policy, CI constraints, and the difference between regression tolerance and measured accuracy. The limitations register is especially good.

Deductions:

- Stale commit/file references in the README tree.
- Test-module header omits discrete despite code including it.
- Benchmark prose still cites 66 functions while current contracts cover 69 names.
- Survival headers and the 1E-18 inverse-survival example need tighter evidence pointers or caveats.

### 6.8 Scope & completeness — 7.2/10

Continuous coverage is strong and coherent. The discrete layer is useful but incomplete: Binomial, Poisson and Geometric are present; Negative Binomial, Hypergeometric and Discrete Uniform are not. There are no multivariate distributions, random variate generation, parameter estimation, or goodness-of-fit utilities. The roadmap states these plainly, so this is not misrepresentation; it is scope.

### 6.9 Reproducibility & process — 7.0/10

The public Python accuracy gate is reproducible and currently green. That is the strongest reproducibility asset.

The weaker parts are structural: the Excel/VBA regression workflow needs a self-hosted Windows runner with desktop Excel and Trust Center access; outsiders cannot reproduce it on GitHub-hosted runners. The benchmark’s observed-value export is manual. The SciPy cross-check is an artifact, but it is not wired into the public gate and the script requires explicit `--grids`.

### 6.10 Repository hygiene & governance — 8.0/10

The repository has MIT license, security policy, contributing guide, code of conduct, issue/PR templates, workflows, a Python-aware `.gitignore`, and no tracked `__pycache__`/`.pyc` artifacts. The demo workbook is present as a binary example, appropriately cautioned in the README.

Deductions: stale structure references and benchmark duplication.

---

## 7. Recommended priority order

1. **Fix Excel CI for the discrete layer.** Import `M_STATS_PROBDIST_DISCRETE.bas` and call `RunDiscreteSuite` in the injected CI bridge. This is the highest-value repair because it aligns automation with the current public surface.
2. **Add LogPMF regression tests.** Cover finite log mass below PMF underflow, exact-zero cases, support edges, and consistency with PMF where representable.
3. **Bring discrete functions under accuracy contracts.** Start with Binomial/Poisson PMF, CDF, SF and inverse in the regimes that reuse incomplete beta/gamma; then holdout-validate and freeze.
4. **Clean the main grid.** Remove obsolete `Lognormal_ParamMeanLog` / `Lognormal_ParamStdDevLog` rows, fill or explicitly mark unobserved `Lognormal_Variance` / `Lognormal_StdDev` rows, and decide whether strict mode should require full observation for active-contract rows.
5. **Refresh documentation from the current commit.** Regenerate the repository tree, update benchmark function counts, align the test-module header with five suites, and add survival/inverse-tail caveat pointers.
6. **Single-source benchmark helpers or document intentional duplication.** Prevent `_ibeta.py` drift.
7. **Add systematic edge fuzzing.** Property-style sweeps around validation guards, exact-integer truncation boundaries, support edges, signed zero, and denormal-adjacent inputs would complement the value-centric suite.

---

## 8. Bottom line

This is a high-quality repository by the standards of Excel VBA and, in parts, by the standards of open-source numerical software generally. The implementation shows real numerical judgment, and the committed Python gate is green with explicit per-regime contracts.

The reason this review lands at **8.7/10** rather than above 9 is that the current assurance story is not yet uniform across the public API. The continuous families are strongly evidenced; the discrete family is implemented and locally tested but lacks contract/grid coverage, LogPMF lacks direct tests, and the Excel CI path currently does not execute the discrete suite. Those are fixable gaps. Closing K3-01, K3-02 and K3-03 would materially raise confidence and likely justify a score above 9 in a follow-up review.

*Review produced independently by Kimi K3 against commit `bdc3b928bf548da3b269b657bf959c4c2b55d0a4`. It reflects the repository state at that commit only and is not a certification for regulated, financial, actuarial, engineering, or safety-critical use.*
