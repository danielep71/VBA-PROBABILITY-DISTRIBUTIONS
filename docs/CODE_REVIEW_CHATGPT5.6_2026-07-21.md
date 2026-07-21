# Independent Code Review — VBA Probability Distributions

> **Repository:** `danielep71/VBA-PROBABILITY-DISTRIBUTIONS`  
> **Branch reviewed:** `main`  
> **Commit reviewed:** [`bdc3b928bf548da3b269b657bf959c4c2b55d0a4`](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/commit/bdc3b928bf548da3b269b657bf959c4c2b55d0a4)  
> **Review date:** 21 July 2026  
> **Reviewer:** OpenAI ChatGPT 5.6 Thinking — extra-high reasoning configuration  
> **Recommended repository path:** `docs/INDEPENDENT_CODE_REVIEW_2026-07-21.md`

---

## 1. Executive assessment

### Overall repository score: **8.7 / 10**

### Production numerical library score: **9.0 / 10**

### Numerical-assurance and release-engineering score: **7.9 / 10**

The repository is an advanced native-VBA numerical library with a coherent architecture, 88 worksheet-facing probability functions, reusable special-function kernels, direct survival APIs, safeguarded inverse solvers, explicit numerical failure contracts, a large deterministic VBA regression suite, and a substantial external benchmark framework.

The production code demonstrates strong floating-point awareness:

- direct upper-tail calculations rather than indiscriminate `1 - CDF`;
- direct inverse-survival functions where subtraction from one would lose the input probability;
- paired incomplete-beta arguments;
- direct lower and upper incomplete-gamma kernels;
- `Log1p`, `Expm1`, `LogExpm1`, and log-density reconstruction;
- guarded addition, multiplication, division, affine transforms, and standardization;
- Loader-style Binomial and Poisson mass calculations;
- explicit supported domains for discrete kernel-backed calculations;
- a stable Lanczos log-gamma difference for unbalanced LogBeta;
- explicit `CVErr(xlErrNum)` versus `CVErr(xlErrValue)` classification.

The repository is not yet release-certified across its complete advertised surface. Four findings materially affect correctness or assurance:

1. **`PROB_LogBeta` dispatches to `PROB_LogGammaDelta` outside that kernel's documented and validated precondition.** For two strongly unbalanced shape parameters that are both much smaller than one, an IEEE-754 reproduction of the committed formula shows absolute LogBeta error near `2E-6`, while the direct identity remains accurate near `1E-14`.

2. **The Excel/VBA CI runner omits `M_STATS_PROBDIST_DISCRETE.bas` and does not execute `RunDiscreteSuite`.** The local `RunAll` procedure includes the discrete suite, but the machine CI bridge manually enumerates only the older suites.

3. **The accuracy analyzer can report PASS for a partially populated contract.** Eight committed extreme-Lognormal rows have blank observations, yet the summary reports the corresponding variance and standard-deviation contracts as PASS with only `2/6` points measured.

4. **The hosted accuracy gate evaluates committed observations rather than the current VBA source and is not triggered by `src/**`.** The observed grid has no source-SHA or source-hash binding, so an algorithm change can leave the tight accuracy gate green without re-executing the current implementation.

The newly added discrete layer is well structured and covered by the local VBA suite, but its 24 public UDFs do not yet have external high-precision accuracy contracts.

### Independent verdict

> **A professionally designed and numerically serious VBA library. The production architecture is strong, but one confirmed small-shape LogBeta defect and several material assurance-pipeline gaps prevent an unconditional production-ready rating for the full repository at the reviewed commit.**

---

# 2. Review scope and methodology

## 2.1 Source retrieval

The exact repository files at the reviewed commit were fetched directly from GitHub and read as committed. The audit did not rely on a branch snapshot inferred from filenames or on copied source excerpts.

The review covered:

- all six production VBA modules;
- the consolidated VBA test module;
- the Excel/PowerShell CI runner;
- both GitHub Actions workflows;
- the accuracy-contract registry;
- the main benchmark grid and generated summary;
- the benchmark generator and Excel exporter;
- the independent holdout summaries;
- numerical-limitations records;
- the main README and benchmark documentation;
- repository configuration and hygiene files.

## 2.2 Execution boundary

The reviewer did not execute desktop Excel or compile the VBA project in this environment.

The review therefore distinguishes among:

- **confirmed source defects**, visible directly in the committed implementation;
- **confirmed assurance defects**, visible directly in scripts, workflows, and committed artifacts;
- **committed numerical evidence**, produced by the repository's own VBA/Python pipeline;
- **targeted independent numerical analysis**, reproducing a committed formula in IEEE-754 binary64 and comparing it with high-precision arithmetic;
- **operational status not independently evidenced**, such as the availability of the self-hosted Excel runner.

No GitHub commit statuses were returned for the reviewed SHA through the combined-status API, so required-check enforcement and current runner availability could not be independently confirmed.

---

# 3. Hard repository metrics

## 3.1 Production source size

| Module | Lines |
|---|---:|
| `M_STATS_PROBDIST_CORE.bas` | 913 |
| `M_STATS_PROBDIST_SPECIALFUNCS.bas` | 1,582 |
| `M_STATS_PROBDIST_NORMALFAMILY.bas` | 3,636 |
| `M_STATS_PROBDIST_TFAMILY.bas` | 3,365 |
| `M_STATS_PROBDIST_CONTINUOUS.bas` | 4,954 |
| `M_STATS_PROBDIST_DISCRETE.bas` | 4,668 |
| **Production VBA total** | **19,118** |
| `M_STATS_PROBDIST_TEST.bas` | 5,617 |
| **Production plus primary tests** | **24,735** |

Additional major artifacts:

| Artifact | Lines |
|---|---:|
| Main `README.md` | 1,442 |
| `compute_errors.py` | 253 |
| `M_STATS_PROBDIST_ACCURACYEXPORT.bas` | 395 |
| `Run-ExcelVbaTests.ps1` | 259 |
| Excel workflow | 64 |
| Hosted accuracy workflow | 59 |

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
| **Total** | **88** |

The discrete count includes ordinary mass and log-mass functions.

## 3.3 Numerical-assurance artifacts

| Evidence | Current size / result |
|---|---:|
| Active accuracy contracts | 76 |
| Main benchmark observations | 726 rows |
| Populated main observations | 718 |
| Blank main observations | 8 |
| Generated summary verdicts | 76 PASS |
| Older-family fresh holdout | 63 contracts, 63 PASS |
| Regime-specific fresh holdout | 10 contracts, 10 PASS |
| Numerical limitations registered | 2 |
| Hosted workflows | 2 |

The generated summary's all-PASS tally must be interpreted with caution because two active contracts are only partially populated.

---

# 4. Scoring methodology

A score of 10 requires all of the following:

- correct algorithms over the documented domain;
- explicit and internally consistent public contracts;
- reproducible independent evidence;
- complete active test coverage;
- CI that executes the current source;
- no known silent wrong-result path;
- documentation generated from, or continuously checked against, authoritative metadata.

## Weighted scorecard

| Area | Weight | Score | Weighted contribution |
|---|---:|---:|---:|
| Functional correctness | 18% | **8.8** | 1.584 |
| Numerical robustness | 17% | **9.0** | 1.530 |
| Architecture and modularity | 10% | **9.6** | 0.960 |
| Public API design | 8% | **9.3** | 0.744 |
| Error handling and diagnostics | 7% | **9.2** | 0.644 |
| Regression testing | 10% | **8.9** | 0.890 |
| External accuracy benchmarking | 11% | **7.8** | 0.858 |
| CI and release engineering | 8% | **7.5** | 0.600 |
| Documentation | 5% | **7.8** | 0.390 |
| Maintainability and hygiene | 5% | **8.9** | 0.445 |
| Performance engineering | 1% | **8.6** | 0.086 |
| **Total** | **100%** |  | **8.731 / 10** |

Rounded overall score:

```text
8.7 / 10
```

## Score interpretation

| Score | Interpretation |
|---:|---|
| 9.5–10.0 | Exceptional and independently release-certified |
| 9.0–9.4 | Advanced professional numerical library with limited gaps |
| 8.0–8.9 | Strong implementation requiring material hardening |
| 7.0–7.9 | Good foundation with significant correctness or assurance work |
| Below 7.0 | Major design, correctness, or governance deficiencies |

---

# 5. Component scores

| Component | Score | Assessment |
|---|---:|---|
| `M_STATS_PROBDIST_CORE` | **9.4** | Strong shared primitives and clear error contracts; a few policy and completeness issues remain. |
| `M_STATS_PROBDIST_SPECIALFUNCS` | **8.7** | Sophisticated kernels, but LogBeta dispatch violates the delta kernel's validated precondition for tiny shapes. |
| `M_STATS_PROBDIST_NORMALFAMILY` | **9.4** | Excellent tail design; Lognormal moments are now reconstructed in one log expression. |
| `M_STATS_PROBDIST_TFAMILY` | **9.2** | Strong Student t/Chi-square/F design and a measured F domain guard; source accuracy text is stale. |
| `M_STATS_PROBDIST_CONTINUOUS` | **9.1** | Broad, consistent and regime-aware; inherits the tiny-shape LogBeta defect. |
| `M_STATS_PROBDIST_DISCRETE` | **8.8** | Well-designed Loader/log-tail layer with explicit limits; external accuracy evidence is absent. |
| `M_STATS_PROBDIST_TEST` | **9.0** | Large and thoughtful local suite; module documentation is stale and CI does not execute the discrete section. |
| External benchmark framework | **7.8** | Strong contracts and holdouts, but partial rows can pass and the documented rebuild path is not reproducible. |
| Excel/PowerShell CI | **7.5** | Good isolated Excel design, but it omits the complete discrete production/test layer. |
| Documentation | **7.8** | Extensive and useful, but materially inconsistent with current contracts and current module inventory. |

---

# 6. Findings summary

| ID | Severity | Area | Finding |
|---|---|---|---|
| P1-01 | P1 | Numerical correctness | `PROB_LogBeta` uses the stable delta outside its precondition and is inaccurate for strongly unbalanced shapes that are both below one |
| P1-02 | P1 | CI coverage | Excel CI omits the discrete production module and the discrete test suite |
| P1-03 | P1 | Accuracy gate | Partially populated and `ERROR`-containing contracts can still receive PASS |
| P1-04 | P1 | Evidence provenance | Tight accuracy evidence is not bound to the current source revision |
| P2-01 | P2 | Assurance coverage | The 24-UDF discrete layer has no external high-precision contracts |
| P2-02 | P2 | Reproducibility | The documented benchmark generator/exporter cannot reproduce the committed main grid |
| P2-03 | P2 | Documentation | Source comments, benchmark README, contract table, and test headers have materially drifted |
| P2-04 | P2 | Contract governance | Five active contracts remain `measured provisional` without an independent holdout |
| P2-05 | P2 | Numerical domain | Broad source statements promise clean non-convergence outside approximate kernel ranges without sufficient evidence |
| P3-01 | P3 | API completeness | Standardization overflow loses sign instead of returning exact probability limits where possible |
| P3-02 | P3 | Discrete boundary | The exact-integer maximum constant is one lower than the documented `2^53 - 1` |
| P3-03 | P3 | Performance | No reproducible timing baseline is committed |
| P3-04 | P3 | Operational assurance | No current combined-status evidence was exposed for the reviewed commit |

---

# 7. Detailed findings

## P1-01 — LogBeta dispatch is inaccurate when both unbalanced shapes are tiny

### Current design

`PROB_LogBeta` orders its arguments and applies the stable-delta route whenever:

```vb
If SmallArg / LargeArg < PROB_LOGBETA_STABLE_RATIO Then
    PROB_LogBeta = _
        PROB_LogGamma(SmallArg) - _
        PROB_LogGammaDelta(LargeArg, SmallArg)
End If
```

The delta kernel documents:

```text
LargeArg >= 1
Increment > 0
```

Its committed validation evidence also emphasizes increments from approximately `0.25` upward and large arguments at or above ordinary scale.

The caller does not enforce `LargeArg >= 1`.

### Counterexample

For:

```text
A = 1E-12
B = 9.9E-14
```

both values are valid positive Beta shapes, and:

```text
Small / Large = 0.099
```

so the stable route is selected.

Reproducing the committed formulas and coefficients in IEEE-754 binary64 gives:

```text
high-precision LogBeta reference   30.03805722019758
current stable-route result        30.03805921300609
absolute error                     1.9928E-6
```

The direct three-log-gamma identity gives approximately:

```text
30.038057220197572
```

with error near `7.6E-15`.

### Impact

An absolute LogBeta error of `~2E-6` produces approximately the same relative error in a factor reconstructed as:

```text
Exp(-LogBeta)
```

This can affect:

- `K_STATS_Beta_Density`;
- Beta CDF and survival normalization;
- Beta inverse calculations.

The error is many orders of magnitude above the frozen unbalanced Beta contracts.

### Root cause

The stable Lanczos-difference formula is being used outside the domain for which its numerical behavior was established. When `LargeArg` is extremely small, denominators and the Lanczos series arrangement enter a conditioning regime not covered by the seam study.

### Recommended correction

Use the stable delta only when both conditions hold:

```text
Small / Large < 0.1
Large >= validated lower bound
```

A conservative immediate dispatch is:

```vb
If LargeArg >= 1# And _
   SmallArg / LargeArg < PROB_LOGBETA_STABLE_RATIO Then
```

For two subunit arguments, retain the direct identity, which is well conditioned because the three log-gamma values are not enormous.

A broader delta domain may be adopted only after a dedicated small-argument study.

### Required tests

Add direct kernel and public-path grids over:

```text
LargeArg:
1E-14, 1E-12, 1E-10, 1E-8, 1E-6, 1E-4, 1E-2, 0.1, 0.5, 0.99

Small/Large:
1E-8, 1E-4, 1E-2, 0.05, 0.099
```

Measure:

- absolute LogBeta error;
- symmetry;
- Beta density;
- Beta CDF;
- Beta survival;
- Beta inverse forward residual;
- continuity at the dispatch boundary.

---

## P1-02 — Excel CI does not include the discrete layer

### Current production and local-test state

The repository contains:

```text
src/M_STATS_PROBDIST_DISCRETE.bas
```

The local public entry point:

```vb
Test_STATS_PROBDIST_RunAll
```

calls:

```vb
RunDiscreteSuite
```

### Current CI runner

`ci/Run-ExcelVbaTests.ps1` imports:

```text
CORE
SPECIALFUNCS
NORMALFAMILY
TFAMILY
CONTINUOUS
TEST
```

It does not import:

```text
DISCRETE
```

The injected CI bridge calls:

```vb
RunCoreSuite
RunNormalFamilySuite
RunTFamilySuite
RunContinuousSuite
```

It does not call:

```vb
RunDiscreteSuite
```

### Impact

The machine result labeled as the complete regression suite excludes:

- 24 public discrete UDFs;
- discrete domain limits;
- Loader PMFs;
- LogPMFs;
- direct discrete tails;
- Binomial/Poisson inverses;
- Geometric formulas;
- discrete error-code behavior.

The CI bridge has diverged from the authoritative local `RunAll`.

### Recommended correction

Add:

```powershell
"src\M_STATS_PROBDIST_DISCRETE.bas"
```

to the import list.

Add:

```vb
RunDiscreteSuite
```

to the CI bridge.

Prefer eliminating duplicated suite enumeration. The injected bridge should call one shared internal routine that is also used by the public `RunAll`, or the production test module should expose a machine-readable entry point directly.

Add a minimum expected assertion count and explicit suite list to the artifact so accidental suite omission is immediately visible.

---

## P1-03 — Partial observations and `ERROR` rows can pass an active accuracy contract

### Current analyzer behavior

`parse_observed` returns `None` for:

```text
empty cell
ERROR
```

`measure_error` skips such rows.

A contract is marked PENDING only when **none** of its matched rows yields a usable observation.

If at least one row is numeric, the contract can PASS even when other rows are empty or `ERROR`.

### Committed example

The main grid contains eight extreme-magnitude Lognormal moment rows with empty `observed_vba` cells:

- four variance rows;
- four standard-deviation rows.

The generated summary reports:

```text
Lognormal_Variance: 2/6 PASS
Lognormal_StdDev:   2/6 PASS
```

The headline tally still states:

```text
FAIL: 0
PENDING: 0
```

### Why this is dangerous

A new boundary point can fail to execute, return `ERROR`, or simply remain unexported while the contract remains green based on older easy points.

This is exactly the situation a release gate must prevent.

### Recommended correction

For every active contract:

```text
usable observations must equal matched observations
```

unless the contract explicitly defines expected-error rows in a separate schema.

Required verdict rules:

| Condition | Verdict |
|---|---|
| All matched rows numeric and within threshold | PASS |
| Any numeric row exceeds threshold | FAIL |
| Any matched row empty | PENDING / incomplete |
| Any matched row `ERROR` when numeric result expected | FAIL |
| No matched rows | PENDING |
| Expected error row returns the required CVErr code | PASS under an error-contract schema |

The summary should report:

```text
numeric / expected / blank / error
```

rather than only `n/len(matched)`.

---

## P1-04 — Accuracy observations are not bound to the reviewed source

### Current architecture

The hosted Accuracy Gate:

- runs on `ubuntu-latest`;
- reads committed `probability_accuracy_grid.csv`;
- evaluates committed observations against contracts;
- does not execute VBA.

The workflow path filter covers:

```text
benchmark/**
```

but not:

```text
src/**
```

The grid contains no:

- commit SHA;
- source-module hashes;
- Excel build provenance per export;
- exporter version hash.

### Impact

A source algorithm can change while:

- the committed observed values remain unchanged;
- the hosted accuracy gate remains green;
- the workflow may not run at all because only `src/**` changed.

The self-hosted Excel regression suite provides useful protection, but it uses broader deterministic tolerances and does not independently remeasure all tight contracts.

### Recommended correction

Preferred design:

1. On the self-hosted Excel runner, generate a temporary reference grid.
2. Import the exact current source.
3. run the accuracy exporter;
4. run `compute_errors.py` in strict mode;
5. publish the generated summary and source manifest.

At minimum, commit a manifest such as:

```json
{
  "observed_at_commit": "...",
  "module_sha256": {
    "CORE": "...",
    "SPECIALFUNCS": "...",
    "NORMALFAMILY": "...",
    "TFAMILY": "...",
    "CONTINUOUS": "...",
    "DISCRETE": "..."
  },
  "excel_version": "...",
  "excel_build": "...",
  "exporter_sha256": "..."
}
```

The hosted gate should fail when the manifest does not match the current source.

The accuracy workflow should also trigger on:

```text
src/**
benchmark/**
```

---

## P2-01 — The discrete public surface has no external accuracy contracts

### Current state

The discrete module adds 24 UDFs:

- Binomial: 8;
- Poisson: 8;
- Geometric: 8.

The main README explicitly states that the discrete family is covered by the VBA regression suite but not yet by the complete external benchmark grid.

### Strength of the implementation

The code uses credible numerical arrangements:

- Loader's Stirling-error/deviance formulation;
- LogPMFs for underflowed masses;
- direct regularized incomplete-beta and incomplete-gamma identities;
- direct survival functions;
- smaller-tail integer inverse searches;
- `Log1p`/`Expm1` geometric formulas;
- explicit kernel-aligned hard limits.

### Assurance gap

There are no machine-readable external contracts for:

- PMF accuracy near the mode;
- PMF and LogPMF deep tails;
- large counts near `2^53`;
- CDF/SF accuracy near kernel limits;
- inverse quantile correctness;
- forward residuals for discrete inverses;
- moment formulas at extreme probabilities;
- count truncation semantics.

### Recommended external measures

For PMF and LogPMF:

```text
relative PMF error when PMF is representable
absolute log-mass error
```

For CDF and survival:

```text
relative error on the smaller tail
absolute complement residual
```

For inverse:

```text
quantile equality
CDF(k-1) < p <= CDF(k)
tail-oriented residual
```

Use independent references from at least two of:

- mpmath/high-precision summation;
- SciPy;
- R;
- recurrence-ratio summation;
- saddlepoint/Loader reference implementations.

---

## P2-02 — The documented benchmark rebuild path cannot reproduce the committed grid

### Documented workflow

The benchmark README describes:

```text
generate_reference_values.py
Excel accuracy exporter
compute_errors.py
```

as the reproducible path.

### Current generator limitations

`generate_reference_values.py` generates the original core/continuous/test-statistic grid. It does not regenerate the complete committed main grid, including:

- appended LogBeta seam observations;
- unbalanced Beta/F study observations;
- inverse-study observations;
- extreme Lognormal moment rows;
- density/helper rows;
- survival-boundary rows;
- discrete rows.

It also keys contracts only by:

```text
(function, regime)
```

even though one inverse function/regime can have multiple measures and thresholds.

### Current exporter limitations

`M_STATS_PROBDIST_ACCURACYEXPORT.bas` lacks dispatch cases for committed grid functions such as:

- `ChiSquare_Density`;
- `F_Density`;
- `Normal_IntervalProbability`;
- the current regime-labelled `Lognormal_ParametersFromMeanStdDev` rows;
- every discrete function.

A full rerun through the documented exporter would write `ERROR` for unsupported function names.

### Documentation mismatch

The exporter says a nonnumeric observed token is a failed point, while the current analyzer skips `ERROR` rows unless no numeric rows remain.

### Recommended architecture

Create one authoritative registry describing each benchmarked row:

```python
FUNCTION_REGISTRY = {
    "ChiSquare_Density": {...},
    "F_Density": {...},
    ...
}
```

Generate:

- reference rows;
- exporter dispatch;
- contract-to-grid validation;
- README coverage tables;

from shared metadata where practical.

Alternatively, split studies cleanly and have one deterministic merge script create the main grid from named evidence sets.

Add a CI `--check` mode that regenerates the grid structure without observations and compares:

- function;
- regime;
- inputs;
- references;
- evidence set;
- row count.

---

## P2-03 — Documentation and contracts have materially drifted

### Benchmark README

The benchmark README still states:

- 66 functions;
- only four covered modules;
- `PROB_StirlingError` absolute error `<= 3E-17`;
- F quantile error `<= 5.9E-13`;
- Standard Normal survival and Lognormal density at `5E-15`;
- misses below `1E-14` may be “below harness precision.”

Current contracts instead include:

- 76 contract rows;
- Stirling absolute threshold `1E-13`;
- F inverse quantile and tail thresholds `2E-10`;
- domain-restricted Normal/Lognormal survival contracts;
- no generic precision-floor exemption.

The generated contract table embedded in the README is stale.

### Special Functions source

`PROB_StirlingError` still publishes:

```text
absolute error <= 3E-17 for every N >= 0.5
```

The fresh older-family holdout measured approximately:

```text
3.57E-14
```

and the frozen contract is:

```text
1E-13
```

### T-family source

The module-level accuracy table still states:

```text
F quantile <= 5.9E-13
```

The frozen contract is:

```text
2E-10 quantile relative error
2E-10 tail-relative forward residual
```

within the enforced `df <= 1E5` envelope.

### Test module header

The test module's introductory inventory still lists five production modules and four suites, although:

- the Discrete module exists;
- `RunAll` calls `RunDiscreteSuite`;
- a fifth suite is present.

### Recommended correction

Treat `accuracy_contracts.csv` as authoritative.

Generate or continuously check:

- benchmark README contract table;
- source-level accuracy summary;
- public module inventory;
- test-suite inventory;
- UDF catalogue.

Add:

```text
python render_contract_table.py --check
```

and documentation consistency checks to hosted CI.

---

## P2-04 — Five active contracts are still provisional

The following contracts are active but marked `measured provisional`:

- Chi-square density;
- F density;
- Normal interval probability;
- Lognormal parameter output: MeanLog;
- Lognormal parameter output: StdDevLog.

They pass the main grid, but they are not included in the committed fresh holdout summaries.

This is acceptable during development, but an all-green summary should expose provenance so readers can distinguish:

```text
validated and frozen
measured provisional
characterization only
```

Before release, add fresh off-grid points and freeze or revise the thresholds.

---

## P2-05 — Kernel-range documentation overstates clean failure behavior

Core and Continuous documentation broadly state that parameters beyond approximate incomplete-beta/incomplete-gamma ranges are attempted and then return clean non-convergence rather than a wrong answer.

The F envelope study established an important counterexample pattern:

> A continued fraction can satisfy its local convergence criterion while its result has drifted outside the accuracy contract.

The F public API now mitigates this through a strict measured envelope. Equivalent blanket guarantees should not be made for other public beta/gamma consumers without evidence.

Recommended wording:

> Inputs outside the validated accuracy domain may fail with `#NUM!`; a successful return outside that domain is not accuracy-certified unless a separate contract states otherwise.

---

## P3-01 — Standardization overflow loses useful sign information

`PROB_TryStandardize` returns only success/failure.

For distribution functions, signed overflow has exact limiting behavior:

| Function | Positive standardized overflow | Negative standardized overflow |
|---|---:|---:|
| Normal density | 0 | 0 |
| Normal CDF | 1 | 0 |
| Normal survival | 0 | 1 |
| Z-score | `#NUM!` | `#NUM!` |

A richer result enum would improve mathematical completeness without changing ordinary finite behavior.

---

## P3-02 — Exact-integer maximum is one lower than documented

The discrete module documents the maximum exact count as:

```text
2^53 - 1 = 9,007,199,254,740,991
```

The constant is:

```vb
9.00719925474099E+15
```

which represents:

```text
9,007,199,254,740,990
```

The implementation is conservative by one integer, not unsafe.

Use a split expression that the VBE preserves, for example:

```vb
9.00719925474099E+15 + 1#
```

and add an exact boundary regression.

---

## P3-03 — No reproducible performance baseline

The algorithms appear appropriate for VBA:

- constant-time closed forms;
- bounded Newton/bisection;
- continued fractions and series with explicit caps;
- no `WorksheetFunction` marshalling;
- fast raw inverse-normal path;
- no volatile worksheet state.

However, the repository has no committed timing benchmark with:

- Excel version/build;
- Office bitness;
- CPU;
- warm-up policy;
- iteration counts;
- scalar-call throughput;
- tail and large-parameter cases.

A small reproducible benchmark would support performance claims and detect accidental slowdowns.

---

## P3-04 — Current required-check enforcement is not evidenced

The reviewed SHA returned no combined statuses through the repository API.

This does not prove the workflows are broken, but it means the audit could not verify:

- self-hosted runner availability;
- successful current Excel execution;
- required-check branch protection;
- current artifact retention.

Publish or link a successful run and document the required check names.

---

# 8. Module-by-module review

## 8.1 `M_STATS_PROBDIST_CORE` — **9.4 / 10**

### Strengths

- true finiteness separated from project magnitude policy;
- guarded primitive arithmetic;
- valid underflow distinguished from overflow;
- stable `Log1p`, `Expm1`, and `LogExpm1`;
- raw inverse-normal seed centralized;
- project-visible but worksheet-hidden scope;
- status-bar side effects disabled by default;
- detailed preconditions and numerical rationale.

### Concerns

- “supported magnitude” naming still mixes representational and accuracy concepts;
- standardization overflow loses sign;
- broad clean-nonconvergence wording is stronger than the evidence;
- the public-surface header should explicitly include `PROB_LogExpm1`.

### Verdict

A strong low-level numerical layer. No architectural rewrite is needed.

---

## 8.2 `M_STATS_PROBDIST_SPECIALFUNCS` — **8.7 / 10**

### Strengths

- Boolean Try-contracts;
- no partial sum returned after iteration exhaustion;
- direct P/Q incomplete-gamma functions;
- paired incomplete-beta arguments;
- stable half-step gamma difference;
- stable Lanczos gamma increment;
- regime-aware LogBeta;
- Loader Stirling and LogChoose kernels;
- safeguarded inverse beta and gamma solvers;
- centralized Lanczos coefficients.

### Concerns

- confirmed invalid dispatch for both-tiny unbalanced Beta shapes;
- source Stirling accuracy claim contradicts the frozen holdout contract;
- no explicit iteration count returned for diagnostics;
- successful results outside validated parameter ranges are not universally accuracy-certified.

### Verdict

Sophisticated and largely well structured, but the small-shape LogBeta dispatch is a real production defect.

---

## 8.3 `M_STATS_PROBDIST_NORMALFAMILY` — **9.4 / 10**

### Strengths

- complete Standard Normal/Normal/Lognormal surface;
- direct survival and inverse survival;
- stable interval probability;
- tail-aware CDF and inverse design;
- log-domain Normal and Lognormal densities;
- one-log reconstruction for Lognormal variance and standard deviation;
- explicit overflow and underflow behavior;
- stable parameter conversion through `Log1p`;
- optional diagnostics without modal UI.

### Current Lognormal moment design

Variance:

```text
Exp(LogExpm1(sigma²) + 2*mu + sigma²)
```

Standard deviation:

```text
Exp(0.5*LogExpm1(sigma²) + mu + 0.5*sigma²)
```

This is the correct architecture for finite-result cancellation.

### Concerns

- eight committed extreme-moment benchmark points remain unexported;
- the current gate masks that incompleteness;
- general Normal still applies a conservative `1E100` restriction despite guarded arithmetic;
- some procedure update dates are inconsistent.

### Verdict

One of the strongest modules in the repository.

---

## 8.4 `M_STATS_PROBDIST_TFAMILY` — **9.2 / 10**

### Strengths

- stable Student t density normalization;
- exact low-degree branches;
- direct upper tails;
- central-mass handling;
- safeguarded inverses;
- direct incomplete-gamma P/Q use;
- F logistic-pair construction;
- log-domain F inverse reconstruction;
- explicit `df <= 1E5` measured F envelope.

### Concerns

- source accuracy table is stale;
- Student t/Chi-square successful results beyond their measured range are explicitly not contract-certified, but this distinction should be more prominent in public UDF comments;
- F envelope policy should be included in each relevant public function's documentation, not only the module header.

### Verdict

Strong numerical design with a sensible strict-domain mitigation for F.

---

## 8.5 `M_STATS_PROBDIST_CONTINUOUS` — **9.1 / 10**

### Strengths

- clear Excel-compatible parameterization;
- direct Gamma/Beta survival;
- log-domain densities;
- guarded moment arithmetic;
- stable Exponential/Weibull left tails;
- stable Weibull large-shape moments;
- full-finite-range Uniform formulas;
- exact support-edge behavior;
- explicit balanced/unbalanced Beta contracts.

### Concerns

- public Beta functions inherit the both-tiny LogBeta defect;
- broad “clean non-convergence, not wrong answer” wording should be narrowed;
- no public local summary of the small-subunit Beta domain gap;
- some measured contracts live only in external metadata.

### Verdict

Broad, disciplined and production-oriented within its measured regimes.

---

## 8.6 `M_STATS_PROBDIST_DISCRETE` — **8.8 / 10**

### Strengths

- 24 coherent public UDFs;
- Binomial/Poisson ordinary and log mass;
- Loader-style central mass;
- direct CDF and survival identities;
- tail-oriented inverses;
- explicit exact-integer policy;
- hard kernel-backed limits;
- Geometric closed forms through `Log1p`/`Expm1`;
- support and endpoint handling;
- consistent `CVErr` and Status behavior;
- local regression coverage.

### Concerns

- no external accuracy contracts;
- no Excel CI execution;
- exact-integer maximum is one lower than documented;
- no generated evidence for near-limit CDF/SF/inverse behavior;
- source module update date predates the reviewed commit's latest changes.

### Verdict

A promising and thoughtfully implemented discrete layer whose assurance maturity lags the continuous library.

---

## 8.7 `M_STATS_PROBDIST_TEST` — **9.0 / 10**

### Strengths

- consolidated counters and verdict;
- dependency-ordered suites;
- exact error-code checks;
- known-value tests;
- identities and complements;
- inverse round-trips;
- deep-tail cases;
- support-edge cases;
- large-range and overflow regressions;
- detailed failure messages;
- machine-readable CI bridge support.

### Concerns

- header inventory omits Discrete;
- header says four suites despite five;
- CI manually bypasses the discrete suite;
- several new benchmark boundaries are not permanent VBA regressions;
- no explicit minimum assertion count is enforced.

### Verdict

Excellent local testing architecture, weakened by machine-run divergence.

---

# 9. External benchmark and contract review

## Strong design choices

The contract schema is appropriately normalized:

```csv
contract_id,function,regime,measure,metric,threshold,domain,provenance,status,evidence,notes
```

The framework correctly distinguishes:

- output error;
- quantile error;
- LogBeta absolute error;
- forward-tail inverse residual.

The two-part `hi;lo` export and `Decimal` reconstruction avoid a false 15-digit serialization floor.

Fresh holdouts are a major strength:

- 63 older-family contracts checked on fresh points;
- 10 regime-specific contracts checked on a dedicated holdout;
- threshold margins recorded.

The numerical-limitations register is also valuable. It currently records:

- mitigated extreme F behavior;
- characterized deep Normal/Lognormal survival relative-error degradation.

## Principal benchmark weaknesses

1. partial rows pass;
2. `ERROR` rows are skipped;
3. source revision is not bound to observations;
4. the generator cannot reconstruct the committed grid;
5. the exporter cannot execute every committed function label;
6. the benchmark README is stale;
7. the discrete layer is absent;
8. five active contracts remain provisional.

## Benchmark verdict

The design is advanced, but the all-green headline currently overstates completeness.

---

# 10. CI and release-engineering review

## Hosted Accuracy Gate

### Strengths

- GitHub-hosted and reproducible;
- strict mode;
- explicit reference dependency;
- degradation self-test;
- read-only permissions;
- benchmark-path trigger.

### Weaknesses

- operates on static observations;
- no source SHA binding;
- does not trigger on source changes;
- partial-row logic allows incomplete contracts to pass;
- no generated-document drift check;
- no discrete accuracy coverage.

## Self-hosted Excel Regression

### Strengths

- isolated temporary workbook;
- exact source import;
- Excel version/build logging;
- fork trust boundary;
- machine-readable counters;
- failure log retrieval;
- COM cleanup;
- artifact upload.

### Weaknesses

- omits Discrete source;
- omits Discrete suite;
- duplicates suite enumeration;
- no compile-stage assertion;
- no required minimum assertion count;
- current runner/status evidence unavailable.

## Recommended required checks

```text
1. Hosted static/contract/documentation check
2. Self-hosted complete Excel VBA regression
3. Self-hosted current-source external accuracy export and strict analysis
```

---

# 11. Documentation review

## Strengths

- extensive main README;
- public API catalogues;
- parameterization;
- worksheet equivalents;
- numerical provenance;
- explicit error policy;
- source-level preconditions;
- benchmark explanations;
- holdout evidence;
- numerical-limitations registry;
- contribution and security guidance.

## Material drift

The following should be corrected atomically:

- benchmark function count;
- benchmark module inventory;
- precision-floor language;
- Stirling threshold;
- F inverse thresholds;
- survival domain restrictions;
- generated contract table;
- test module suite inventory;
- accuracy exporter dependencies and supported functions;
- discrete external-assurance status;
- current CI coverage statement.

## Documentation verdict

The volume and intent are excellent. Current synchronization is not sufficiently reliable for a numerical contract.

---

# 12. Security and platform assessment

No high-severity security defect was identified in production numerical code.

Positive controls:

- no external DLL;
- no network access from production UDFs;
- no shell invocation from production modules;
- no modal UI from numerical UDFs;
- read-only workflow permissions;
- fork PR exclusion on the self-hosted Excel runner;
- isolated workbook;
- explicit COM cleanup;
- Python hosted gate separate from Excel trust requirements.

The self-hosted runner necessarily enables programmatic VBA project access and lowers automation security for its isolated Excel instance. It should be dedicated, patched, access-controlled, and unavailable to untrusted code.

---

# 13. Release-readiness assessment

## Suitable now

Within documented and measured domains, the code is suitable for:

- teaching and numerical demonstrations;
- controlled Excel/VBA model components;
- model-validation comparisons;
- quantitative prototyping;
- direct-tail calculations;
- internal governed libraries with independent application-specific validation.

## Blocking conditions for a full release-quality claim

Close at least:

```text
P1-01 both-tiny unbalanced LogBeta dispatch
P1-02 discrete omission from Excel CI
P1-03 partial/ERROR benchmark PASS behavior
P1-04 stale observation/source binding
```

Then:

- rerun the complete Excel suite;
- regenerate all external observations;
- run strict contracts;
- publish source/build manifests;
- require all checks on `main`.

## Regulated or high-stakes use

For banking, actuarial, engineering, or similar governed use:

- pin the full commit SHA;
- archive Excel version/build;
- archive the exact benchmark summary;
- include domain limits in model documentation;
- independently validate the parameter regimes actually used;
- treat provisional contracts as non-frozen;
- treat the small-shape Beta defect as an explicit exclusion until fixed.

---

# 14. Prioritized remediation plan

## Priority 1 — Correct tiny-shape LogBeta dispatch

1. Add a lower-bound condition to the stable-delta route.
2. Create a both-subunit seam study.
3. Add public Beta regressions.
4. extend the external grid.
5. update contract domains.

## Priority 2 — Make Excel CI truly complete

1. Import `M_STATS_PROBDIST_DISCRETE.bas`.
2. Execute `RunDiscreteSuite`.
3. eliminate duplicated suite lists.
4. enforce a minimum assertion count.
5. publish current run evidence.

## Priority 3 — Make strict accuracy mean complete accuracy

1. Fail on blank active rows.
2. Fail on `ERROR` where numeric output is expected.
3. show blank/error counts.
4. add an expected-error contract type.
5. add a CI degradation test for partial rows.

## Priority 4 — Bind evidence to source

1. add source hashes and commit SHA;
2. trigger accuracy checks on `src/**`;
3. execute current-source observations on Excel;
4. reject stale manifests.

## Priority 5 — Add external discrete contracts

Start with:

- LogPMF absolute error;
- PMF relative error;
- CDF/SF smaller-tail relative error;
- inverse integer inequalities;
- near-limit parameter cases.

## Priority 6 — Rebuild the benchmark pipeline

1. unify row registry;
2. unify exporter dispatch;
3. deterministic study merge;
4. regenerate docs;
5. add `--check` modes.

## Priority 7 — Synchronize documentation

Generate or validate all numerical tables and inventories from machine-readable metadata.

---

# 15. Final verdict

The repository demonstrates unusually strong numerical engineering for native VBA:

- clear layering;
- transparent algorithms;
- stable tail methods;
- reusable beta/gamma kernels;
- safeguarded inverses;
- stable continuous and discrete masses;
- explicit domains;
- substantial tests;
- meaningful holdout evidence;
- disciplined worksheet-error semantics.

The codebase is materially stronger than a typical VBA helper collection and is credible as a focused numerical library.

The current assurance headline is nevertheless stronger than the underlying evidence. A confirmed small-shape LogBeta dispatch defect affects valid Beta parameters; the CI excludes the entire discrete layer; the strict benchmark can pass partially populated contracts; and the committed observations are not cryptographically or operationally tied to the current source.

> **Overall score: 8.7 / 10**  
> **Classification: strong advanced numerical library requiring material release-pipeline hardening and one focused production correction.**

---

# Appendix A — Current representative accuracy evidence

| Contract | Main-grid worst | Threshold |
|---|---:|---:|
| Balanced Beta density | `2.78E-15` | `1E-14` |
| Unbalanced Beta density | `1.23E-12` | `4E-12` |
| Balanced Beta CDF | `5.43E-15` | `2E-14` |
| Unbalanced Beta CDF | `5.30E-11` | `1E-10` |
| Unbalanced Beta survival | `9.16E-11` | `2E-10` |
| Unbalanced Beta inverse quantile | `4.72E-11` | `1E-10` |
| Unbalanced Beta inverse tail residual | `3.94E-10` | `1E-9` |
| Validated F CDF | `2.59E-14` | `1.1E-10` |
| Validated F inverse quantile | `7.77E-11` | `2E-10` |
| Validated F inverse tail residual | `6.91E-11` | `2E-10` |
| `PROB_LogBeta` committed grid | `5.93E-14` absolute | `2E-13` |
| LogGamma | `3.77E-15` | `6.1E-14` |
| LogGamma half difference | `1.53E-14` | `2E-14` |
| Lognormal density | `2.36E-14` | `3E-14` |
| Standard Normal survival | `1.52E-14` | `2E-14` |
| Student t density | `1.93E-14` | `2E-14` |
| Student t inverse | `9.96E-14` | `3E-12` |
| Gamma CDF | `1.71E-14` | `2E-14` |
| Gamma survival | `1.60E-14` | `5E-14` |
| Weibull standard deviation | `4.12E-15` | `5E-15` |

These figures are committed evidence, not results independently rerun by the reviewer.

---

# Appendix B — Targeted LogBeta numerical check

## Formula reproduced

The audit reproduced the committed binary64 calculations for:

- Lanczos `PROB_LogGamma`;
- compensated `PROB_Log1p`;
- `PROB_LogGammaDelta`;
- `PROB_LogBeta` dispatch.

Reference values used 80-digit arithmetic.

## Representative result

```text
A = 1E-12
B = 9.9E-14
ratio = 0.099
```

| Method | LogBeta | Absolute error |
|---|---:|---:|
| High-precision reference | `30.03805722019758` | — |
| Current stable dispatch | `30.03805921300609` | `1.9928E-6` |
| Direct identity | `30.038057220197572` | `~7.6E-15` |

This is a targeted independent formula check. It should be confirmed in the actual VBA exporter and made a permanent regression.

---

# Appendix C — Suggested GitHub issues

1. `Fix PROB_LogBeta dispatch for strongly unbalanced shapes below one`
2. `Include M_STATS_PROBDIST_DISCRETE and RunDiscreteSuite in Excel CI`
3. `Fail accuracy contracts when any matched row is blank or ERROR`
4. `Bind benchmark observations to source SHA and module hashes`
5. `Add external accuracy contracts for Binomial, Poisson, and Geometric`
6. `Unify benchmark generation, study merge, and VBA export dispatch`
7. `Regenerate benchmark README from current accuracy_contracts.csv`
8. `Align Stirling and F inverse source comments with frozen contracts`
9. `Add independent holdout for the five provisional contracts`
10. `Correct discrete exact-integer maximum to 2^53 - 1`
11. `Add performance baseline with Excel build and CPU metadata`
12. `Publish and require current Excel and accuracy workflow checks`

---

# Appendix D — Evidence confidence

| Conclusion | Confidence |
|---|---|
| Repository metrics | High |
| Architecture and public API inventory | High |
| Tiny-shape LogBeta defect | High — direct formula reproduction plus precondition mismatch |
| Discrete CI omission | High — explicit runner and bridge source |
| Partial benchmark PASS defect | High — analyzer source plus committed `2/6` summary |
| Source/observation staleness risk | High — workflow and grid schemas |
| Main accuracy figures | High as committed evidence; not independently re-executed |
| Discrete numerical correctness | Medium-high from static review and local test design; external oracle absent |
| Current workflow operational status | Low-to-medium — no combined statuses returned |
| Performance assessment | Medium — algorithmic review only |
