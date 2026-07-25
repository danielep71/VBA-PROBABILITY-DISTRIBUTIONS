# Independent Code Review — `VBA-PROBABILITY-DISTRIBUTIONS`

| Field | Value |
|---|---|
| Repository | `danielep71/VBA-PROBABILITY-DISTRIBUTIONS` |
| Reviewed commit | `6516f0991e56434e9dc23453ed14c65528d8e22b` (`main`) |
| Snapshot source | `codeload.github.com` archive of the reviewed commit |
| Review date | 2026-07-23 |
| Reviewer | Kimi K3 (Moonshot AI), independent automated code review |
| Review basis | Cold re-read of the current source, tests, benchmark files, CI scripts and documentation; fresh metric extraction; execution of the reproducible Python accuracy gate and gate-degradation self-test |
| Execution boundary | This environment cannot run desktop Excel/VBA. I did **not** execute `Test_STATS_PROBDIST_RunAll` inside Excel. I did run the Python gate successfully. The Excel workflow remains self-hosted Windows/Excel. |

> **Overall weighted score: 9.4 / 10** — exact weighted value **9.426 / 10**.  
> Verdict: the project is now in a strong, coherent evidence state. The earlier assurance gaps are closed: every public `K_STATS_*` function is referenced by tests and covered by an active accuracy contract; the grid is fully observed; Discrete Uniform is now in both the VBA harness and the external contract regime; stale artifacts and benchmark duplication have been removed; deep-tail inverse evidence has been added. The remaining deductions are small: summary wording around the separate limitations register, nine measured-provisional contracts, the size of the discrete module, and the inherent self-hosted Excel reproducibility boundary.

---

## 1. Executive summary

At commit `6516f099`, the repository has reached the state its documentation claims: a native Excel-VBA probability library with a broad continuous catalogue, a complete six-family discrete catalogue, direct tail APIs, guarded numerical kernels, a consolidated VBA regression harness, and a machine-readable external accuracy regime.

The current public surface is **112 worksheet/VBA `K_STATS_*` functions**. The test module references **112/112** of them. The accuracy-contract regime covers **112/112** public UDFs. The contract file contains **132 active contracts** across **117 function names**: **123 validated and frozen**, **9 measured provisional**. The main accuracy grid contains **1,471 rows**, and **1,471 rows have non-empty `observed_vba` values**. The generated summary reports **132 PASS**, **0 FAIL**, **0 KNOWN LIMITATION**, **0 CHARACTERIZATION ONLY**, **0 PENDING**.

I ran the reproducible checks against the reviewed snapshot:

```text
cd benchmark
python compute_errors.py
```

Result:

```text
accuracy_summary.md: FAIL=0 KNOWN_LIMITATION=0 CHARACTERIZATION_ONLY=0 PENDING=0 [gate: strict]
  gate passed (exit 0).
```

I also ran:

```text
python test_gate_degradation.py
```

Result: exit `0`. The self-test confirms that when the `_ibeta` helper is unavailable, the two active tail-residual contracts become blocking `PENDING` items and the strict gate exits non-zero.

The most important change since the earlier same-day review is synchronization. The README now says the external pipeline covers all six discrete families; the benchmark README says the generated block is authoritative and derived from `accuracy_contracts.csv`; `README_old.md` and the duplicate root `environment.txt` are gone; the test `RunAll` header says five suites; survival headers now point to `SurvivalTailRel`; deep-tail inverse-normal rows and contracts have been added; benchmark helper duplication has been removed.

No blocking numerical correctness defect was found in the committed evidence or the static read.

---

## 2. Hard-number snapshot

Extracted from the checked-out archive of `6516f099`.

| Metric | Current value |
|---|---:|
| Files in archive | 88 |
| `.bas` files | 19 |
| Python files | 22 |
| Markdown files | 21 |
| CSV files | 13 |
| Workflow/config YAML files | 3 |
| PowerShell files | 1 |
| Macro-enabled demo workbook | 1 |
| Production source modules | 6 |
| Production source lines | 22,631 |
| Production source code/comment/blank lines | 7,943 / 12,693 / 1,995 |
| Test module lines | 6,150 |
| Test module code/comment/blank lines | 2,357 / 3,246 / 547 |
| Benchmark `.bas` lines | 2,840 |
| Public `K_STATS_*` functions | 112 |
| Public `K_STATS_*` referenced by tests | 112 / 112 |
| Public `K_STATS_*` with accuracy contracts | 112 / 112 |
| Project-scoped `PROB_*` public kernels | 30 |
| Test procedures | 120 total: 6 public runners + 114 section/helper procedures |
| Static assertion statement lines in tests | 758 |
| `Option Explicit` in VBA files | 19 / 19 |
| `MsgBox` in executable source/test code | 0 |
| Executable `Application.WorksheetFunction` use | 0 |
| Accuracy contracts | 132 active |
| Contract function names | 117 |
| Contract provenance | 123 validated and frozen; 9 measured provisional |
| Main accuracy grid rows | 1,471 |
| Grid rows with non-empty observed values | 1,471 |
| Empty observed grid rows | 0 |
| Duplicate `.bas` hash groups under `benchmark` | 0 |
| Duplicate `_ibeta.py` study helpers | 0 |
| Strict accuracy gate | PASS, exit 0 |
| Gate degradation self-test | PASS, exit 0 |

Discrete-family evidence is complete for the current public discrete surface:

| Family | Public UDFs | Contracts | Grid rows | Provenance |
|---|---:|---:|---:|---|
| Binomial | 8 | 8 | 210 | 8/8 validated and frozen |
| Poisson | 8 | 8 | 60 | 8/8 validated and frozen |
| Geometric | 8 | 8 | 72 | 8/8 validated and frozen |
| Negative Binomial | 8 | 8 | 210 | 8/8 validated and frozen |
| Hypergeometric | 8 | 8 | 70 | 8/8 validated and frozen |
| Discrete Uniform | 8 | 8 | 89 | 8/8 validated and frozen |

The nine remaining measured-provisional contracts are all green; they are evidence-polish items rather than failures:

| Contract | Threshold | Evidence | Measured note |
|---|---:|---|---|
| `ChiSquare_Density.all.output_rel` | 1E-13 | density_helpers | worst measured 2.37E-14 |
| `F_Density.all.output_rel` | 1E-13 | density_helpers | worst measured 4.81E-14 |
| `Normal_IntervalProbability.all.output_rel` | 1E-14 | density_helpers | worst measured 1.17E-15 |
| `Lognormal_ParametersFromMeanStdDev.param_meanlog.output_rel` | 5E-15 | density_helpers | worst measured 1.29E-16 |
| `Lognormal_ParametersFromMeanStdDev.param_stddevlog.output_rel` | 5E-15 | density_helpers | worst measured 1.17E-16 |
| `NormalStandard_InverseSurvival.split_boundary.output_rel` | 1E-10 | main grid | measured worst 5.5E-11 at the CDF split |
| `NormalStandard_InverseCumulative.split_boundary.output_rel` | 1E-10 | main grid | measured worst 5.5E-11 at the CDF split |
| `Normal_InverseSurvival.split_boundary.output_rel` | 2E-10 | main grid | measured worst 1.0E-10 at the CDF split |
| `Lognormal_InverseSurvival.split_boundary.output_rel` | 5E-10 | main grid | measured worst 3.8E-10 at the CDF split |

Interpretation notes:

- The production comment ratio is about **56.1% of production source lines**. The comments are mostly load-bearing: provenance, preconditions, error policy, regime thresholds, limitations and update dates.
- The static assertion count is a count of assertion statement lines, not an executed assertion total. A desktop Excel run is required for the executed count.
- The archive contains no `README_old.md` and no duplicate root `environment.txt`.

---

## 3. Scoring rubric

Scale: **10** = exceptional for the problem domain; **9** = excellent with only bounded or research-level gaps; **8** = strong professional quality with material gaps; **7** = solid but with gaps a serious user must manage.

| # | Category | Weight | Score | Weighted | Main reason |
|---:|---|---:|---:|---:|---|
| 1 | Numerical correctness & methodology | 20% | **9.4** | 1.880 | Correct methods, stable tails, guarded arithmetic, explicit domains; deep-tail behavior is now evidenced rather than merely claimed. |
| 2 | Verification & benchmark evidence | 18% | **9.5** | 1.710 | All public UDFs under active contracts; grid fully observed; strict gate and degradation self-test pass; nine contracts remain provisional. |
| 3 | Testing & CI execution | 12% | **8.8** | 1.056 | Tests reference all public UDFs and CI imports/runs discrete; still self-hosted Excel and not independently executed here. |
| 4 | Robustness & error contract | 10% | **8.9** | 0.890 | Consistent `Variant`/`CVErr`, Boolean `Try*` kernels, predictable failure classification; heavy `GoTo` idiom remains. |
| 5 | API design & Excel integration | 10% | **8.9** | 0.890 | Clear naming, Excel-compatible parameterization, direct tails and LogPMF; no vector/array API. |
| 6 | Code quality & maintainability | 10% | **8.4** | 0.840 | Excellent house style and duplication cleanup; `DISCRETE` remains very large at 8,157 lines. |
| 7 | Documentation | 8% | **8.4** | 0.672 | README/benchmark prose are now synchronized; summary still needs a limitation-register clarification. |
| 8 | Scope & completeness | 8% | **8.5** | 0.680 | Broad continuous catalogue plus six discrete families; still no multivariate distributions, RNG, or array API. |
| 9 | Reproducibility & process | 6% | **7.6** | 0.456 | Public Python gate is reproducible; Excel execution/export remains manual/self-hosted and the hosted gate checks committed observations. |
| 10 | Repository hygiene & governance | 4% | **8.8** | 0.352 | Stale artifacts and duplicate helpers removed; good templates/license/security/gitignore remain. |
|  | **Overall** | **100%** |  | **9.426** | Rounded: **9.4 / 10** |

---

## 4. What is genuinely strong

### 4.1 Public API, tests, contracts and grid now agree

The strongest current property is coherence. At the reviewed commit:

- every public `K_STATS_*` function is referenced by the test module;
- every public `K_STATS_*` function is represented in `accuracy_contracts.csv`;
- every grid row has an observed value;
- the generated summary is current;
- the strict gate passes;
- the gate’s own degradation behavior is tested.

That is the correct end state for an evidence-led numerical library: public surface, regression suite, machine-readable contracts, observed grid and generated verdict all point to the same set of functions.

### 4.2 The numerical kernel layering remains correct

The dependency direction remains clean:

```text
Worksheet/VBA callers
  -> K_STATS_* family modules
    -> SPECIALFUNCS kernels
      -> CORE constants, predicates and guarded arithmetic
```

`CORE` owns shared constants and elementary numerical primitives. `SPECIALFUNCS` owns log-gamma/log-beta/incomplete-beta/incomplete-gamma kernels. Family modules own parameterization, support edges, validation, tail orientation and worksheet error mapping. `Option Private Module` in `CORE` and `SPECIALFUNCS` correctly hides kernels from the worksheet Function Wizard while leaving them project-visible.

### 4.3 The code still avoids the classic Excel/VBA numerical traps

Concrete examples from the fresh read:

- `PROB_Log1p` and `PROB_Expm1` use compensated forms rather than naive `Log(1 + X)` and `Exp(X) - 1`.
- `PROB_TryExp`, `PROB_TryAdd`, `PROB_TryMultiply`, `PROB_TryDivide`, `PROB_TryStandardize`, and `PROB_TryAffineTransform` turn overflow into Boolean failure instead of unexpected runtime errors.
- `PROB_TryBetaRegularized` takes both `X` and `Y = 1 - X` from callers and never reconstructs the complement internally.
- `PROB_TryBetaContinuedFraction` returns `False` on non-convergence and explicitly never returns a partial sum.
- `PROB_LogGammaDelta` and the unbalanced `PROB_LogBeta` branch remove catastrophic cancellation in the naive three-log-gamma identity.
- Binomial and Poisson use Loader-style Stirling-error/deviance arrangements; Negative Binomial and Hypergeometric reuse stable log-mass/combinatorial kernels.
- Discrete Uniform implements signed inclusive bounds, real-threshold step CDF/SF behavior, a corrected lower-bound quantile and cancellation-safe moments.
- The F family enforces a measured envelope: `PROB_F_MAX_DF = 100000#`, and degrees of freedom above that are rejected for CDF/SF/inverse while closed-form density remains unrestricted.

### 4.4 Failure behavior is explicit and Excel-native

Public functions return `Variant`; valid results return as `Double`; predictable numerical/domain failures return `CVErr(xlErrNum)`; unexpected runtime failures return `CVErr(xlErrValue)`; valid exponential underflow returns `0`; no public function raises `MsgBox`.

Static counts support the policy: **0 `MsgBox` calls** in executable source/test code, **530 `CVErr` occurrences** across VBA files, and **0 executable `Application.WorksheetFunction` uses**.

---

## 5. Findings register

Severity scale: **High** = blocks or materially weakens assurance; **Medium** = real gap a maintainer should schedule; **Low** = polish/consistency; **Info** = acceptable trade-off or scope note.

| ID | Severity | Area | Finding | Evidence | Recommendation |
|---|---|---|---|---|---|
| K3-6516f099-01 | **Low** | Gate/report wording | `accuracy_summary.md` reports `KNOWN LIMITATION: 0` while `numerical_limitations.csv` contains two limitation rows. The tally is about contract verdicts, not the limitation register, but a casual reader can misread it. | `accuracy_summary.md` verdict tally; `numerical_limitations.csv`. | Add one sentence under the tally: limitation-register entries are separate from contract verdict states. |
| K3-6516f099-02 | **Medium** | Module size | `M_STATS_PROBDIST_DISCRETE.bas` is 8,157 lines, larger than the other five production modules combined. The style is consistent, but navigation and merge risk are now the main maintainability issue. | Fresh VBA line metrics: `DISCRETE` 8,157 lines; all production source 22,631 lines. | Consider splitting the discrete module by family while preserving the public API, e.g. `DISCRETE_CORE`, `DISCRETE_CLASSICAL`, `DISCRETE_EXTENDED`. |
| K3-6516f099-03 | **Low** | Provisional contracts | Nine contracts remain measured provisional. All pass, but they are not yet holdout-frozen like the rest of the regime. | `accuracy_contracts.csv`: 123 validated/frozen, 9 measured provisional. | Either holdout-validate and freeze them, or document why density-helper and split-boundary contracts are intended to remain provisional. |
| K3-6516f099-04 | **Info** | Process | The hosted accuracy gate evaluates the committed observation grid; it does not execute Excel or regenerate observations from current VBA source. The Excel regression workflow is self-hosted and not reproducible by outsiders. | `accuracy-gate.yml`, `excel-vba-regression.yml`, README caution note. | Keep the caution. For releases, require both the strict Python gate and a fresh self-hosted Excel run before tagging. |
| K3-6516f099-05 | **Info** | Scope | There are still no multivariate distributions, random variate generation, parameter estimation, goodness-of-fit utilities or array/vector entry points. | Public API inventory and README positioning. | Acceptable as stated scope. If the discrete catalogue is now considered complete, say so explicitly and mark these as separate future products rather than gaps. |
| K3-6516f099-06 | **Info** | Fuzz/property testing | The suite remains value-centric. It is broad, but not randomized/property-based around guard edges. | Test harness design and static assertion inventory. | Optional hardening: add deterministic seeded sweeps around exact-integer truncation boundaries, support edges, signed zero, denormal-adjacent inputs and split-boundary neighborhoods. |

---

## 6. Category detail

### 6.1 Numerical correctness & methodology — 9.4/10

The numerical methodology remains excellent. The implementation is arranged around known failure modes rather than around worksheet convenience.

Positive evidence:

- Direct survival functions are exposed across the catalogue, and incomplete-gamma upper tails are evaluated as `Q` rather than recovered as `1 - P`.
- Incomplete beta receives paired complementary arguments and never forms `1 - X` internally.
- Iterative kernels return Boolean failure and leave results non-contractual on failure.
- The incomplete-beta inverse solves on the smaller tail and returns both `X` and `Y`, avoiding a later destructive complement in callers such as F.
- `PROB_LogGammaDelta` and the unbalanced `PROB_LogBeta` branch remove the catastrophic cancellation in the naive three-log-gamma identity.
- The discrete module uses Loader-style mass arrangements for Binomial/Poisson and stable log-mass/combinatorial arrangements for Negative Binomial/Hypergeometric.
- Discrete Uniform handles signed bounds, real thresholds, exact support cardinality and corrected quantile jumps.
- The F envelope rejects a region where local convergence can still be inaccurate.
- Deep-tail inverse-normal behavior is now evidenced by `deep_tail` contracts to `q = 1E-300`, with a separate measured `split_boundary` regime where relative accuracy dips and recovers.

Deductions:

- The measured ceilings are real: unbalanced Beta/F carry looser thresholds than balanced regimes, survival-tail relative accuracy degrades beyond the central region, and inverse split-boundary contracts are intentionally looser than deep-tail contracts.

### 6.2 Verification & benchmark evidence — 9.5/10

This category is now very strong. The benchmark design is explicit and covers the full public surface: machine-readable contracts, regime-aware thresholds, provenance states, generated summary, two-part `hi;lo` observed values to preserve VBA Doubles through CSV, Decimal-based error computation, and a gate with tested degradation behavior.

Current state:

- 132 active contracts; all 132 generated summary rows are PASS.
- 123 contracts are validated and frozen; 9 are measured provisional.
- The strict gate exits 0 on the committed grid.
- The degradation self-test exits 0 and proves active tail contracts cannot silently become non-blocking when the evaluator is missing.
- Grid rows are fully observed: 1,471/1,471.

Remaining deduction: nine contracts remain measured provisional rather than validated/frozen.

### 6.3 Testing & CI execution — 8.8/10

The local VBA harness is substantial: 120 test procedures and 758 static assertion statement lines. The suite order is dependency-aware and includes all six discrete families. The public cross-reference is complete: 112/112 public UDFs are referenced.

The CI runner imports `src/M_STATS_PROBDIST_DISCRETE.bas`, and the injected bridge calls `RunDiscreteSuite`. That structural gap is closed.

Deductions:

- I could not execute desktop Excel in this environment, so I verified the harness statically and ran the Python gate only.
- The Excel workflow remains self-hosted and not reproducible by outside contributors on GitHub-hosted runners.
- The suite remains value-centric; deterministic seeded fuzzing would still add assurance around guard edges.

### 6.4 Robustness & error contract — 8.9/10

The error contract is one of the library’s best features. Public functions are worksheet-safe, validation distinguishes invalid domains from numerical failure, kernels do not validate caller preconditions they do not own, and status diagnostics are available through `ByRef Status` without depending on `Application.StatusBar`.

Deductions:

- The code uses `GoTo` heavily for VBA structured flow. This is idiomatic VBA error handling, but it remains an audit cost: 518 `GoTo` occurrences across VBA files.
- Some policy choices are defensible but not universally preferred, such as `#NUM!` on overflowing standardization rather than returning a limit value.

### 6.5 API design & Excel integration — 8.9/10

The API is clear and migration-friendly. Names state exactly what they compute, parameterization follows Excel conventions including Excel’s own rate/scale inconsistencies, and the library adds direct survival, inverse-survival and LogPMF entry points that Excel lacks or makes awkward.

Deductions:

- There are no array/vectorized entry points, which matters for Monte Carlo workloads.
- The public surface is broad enough that the wiki/API reference must be kept tightly synchronized with source headers.

### 6.6 Code quality & maintainability — 8.4/10

The house style is rigorous: `Option Explicit` everywhere, structured banners, explicit declarations, `Double` literals, consistent labels, and header fields that usually match the code. Shared constants are held once, and the Lanczos coefficients are single-sourced for `PROB_LogGamma` and `PROB_LogGammaDelta`.

The duplication cleanup is real: benchmark `.bas` duplicate hash groups are gone and `_ibeta.py` study-helper duplicates are gone.

Remaining deduction: `DISCRETE` is now very large. It is well organized, but at 8,157 lines it increases review and merge cost.

### 6.7 Documentation — 8.4/10

The documentation is now synchronized with the evidence in the places that previously contradicted it. The README says all six discrete families are covered by the external pipeline and present in both harness and grid. The benchmark README correctly says the generated contract block is authoritative and derived from `accuracy_contracts.csv`.

Remaining deduction: the summary tally still needs one sentence preventing `KNOWN LIMITATION: 0` from being misread as “no entries in the limitations register.”

### 6.8 Scope & completeness — 8.5/10

The continuous catalogue is strong and coherent. The discrete catalogue now includes Binomial, Poisson, Geometric, Negative Binomial, Hypergeometric and Discrete Uniform, with PMF/LogPMF/CDF/SF/inverse/moments across all six.

Remaining omissions are explicit scope choices rather than defects: multivariate distributions, random variates, parameter estimation, goodness-of-fit utilities and array/vector APIs are absent.

### 6.9 Reproducibility & process — 7.6/10

The public Python accuracy gate is reproducible and currently green. That is the strongest reproducibility asset.

The weaker parts are structural: the Excel/VBA regression workflow needs a self-hosted Windows runner with desktop Excel and Trust Center access; outsiders cannot reproduce it on GitHub-hosted runners. The benchmark’s observed-value export remains manual. The hosted gate verifies committed observations, not a fresh Excel execution of current source.

### 6.10 Repository hygiene & governance — 8.8/10

The repository has MIT license, security policy, contributing guide, code of conduct, issue/PR templates, workflows, and a Python-aware `.gitignore`. No tracked `__pycache__`/`.pyc` artifacts were present in the archive before running the local gate.

The stale `README_old.md` and duplicate root `environment.txt` are gone. Benchmark duplicate helpers are gone.

---

## 7. Recommended priority order

1. **Add the limitation-register clarification to `accuracy_summary.md`.** This is a one-line documentation fix that prevents misreading `KNOWN LIMITATION: 0`.
2. **Freeze or explicitly classify the nine measured-provisional contracts.** All are green; the work is evidence finalization, not defect repair.
3. **Split the discrete module before the next expansion.** Keep the public API unchanged, but reduce the 8,157-line module into family modules or a discrete core plus family modules.
4. **Keep the release rule explicit.** A release should require both the reproducible Python strict gate and a fresh self-hosted Excel run, because the hosted gate does not execute Excel.
5. **Optionally add deterministic seeded fuzzing.** Focus on exact-integer truncation boundaries, support edges, signed zero, denormal-adjacent inputs and split-boundary neighborhoods.

---

## 8. Bottom line

This is a high-quality repository by the standards of Excel VBA and, in important respects, by the standards of open-source numerical software generally. The implementation shows real numerical judgment, and the current evidence chain is coherent: public API, tests, contracts, grid, generated summary and strict gate all agree.

The earlier major gaps are closed. Discrete distributions are no longer the weak flank: all six current discrete families are implemented, tested and covered by validated-and-frozen accuracy contracts. Deep-tail inverse-normal behavior now has explicit grid evidence and contracts. Stale documentation and duplicate artifacts have been cleaned up.

The remaining work is finish work: clarify the limitation tally, freeze the last provisional contracts, split the oversized discrete module, and keep the release process honest about the difference between the reproducible Python gate and the self-hosted Excel run. Those are small relative to the state of the code.

*Review produced independently by Kimi K3 against commit `6516f0991e56434e9dc23453ed14c65528d8e22b`. It reflects the repository state at that commit only and is not a certification for regulated, financial, actuarial, engineering, or safety-critical use.*
