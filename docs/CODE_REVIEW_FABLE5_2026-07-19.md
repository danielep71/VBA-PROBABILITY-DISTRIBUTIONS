# Independent Code Review — VBA-PROBABILITY-DISTRIBUTIONS

| | |
|---|---|
| **Repository** | `danielep71/VBA-PROBABILITY-DISTRIBUTIONS` |
| **Commit reviewed** | `b0c4b2f` |
| **Review date** | 2026-07-19 |
| **Reviewer** | Claude Fable 5 (Anthropic) — automated review with human-readable rationale |
| **Scope** | All VBA source modules, test suite, benchmark and verification infrastructure, CI, documentation, repository hygiene |
| **Method** | Full static read of every `.bas` module; metric extraction (API surface, assertion counts, error-contract usage); cross-checking documentation claims against machine-readable artifacts; verification that the accuracy-gate tooling runs and returns the exit codes it documents |

---

## 1. Executive summary

This is a VBA library of probability-distribution worksheet functions (Normal
family, Student-t family, Gamma/Beta/Exponential/Weibull/Uniform, plus the
special-function kernels beneath them) whose stated purpose is to exceed
Excel's built-in accuracy, particularly in deep tails where Excel's functions
collapse to zero.

The headline result of this review: **the library's engineering discipline is
far above what is typical for the VBA ecosystem, and its verification
infrastructure is above what is typical for open-source numerical libraries in
any language.** Accuracy is not asserted; it is contracted per function *and
per parameter regime* in a machine-readable file, measured against 50-digit
references on a 685-point grid, gated by a script with strict/development exit
codes, and — for the most recently calibrated bounds — validated on an
independent holdout grid before being frozen. A documented numerical
limitation that the library cannot currently meet is recorded honestly in a
separate limitations register instead of being hidden or spun.

The main weaknesses are scope (no discrete distributions — a significant gap
for a general-purpose statistics library), some residual duplication of
contract information across prose documentation, a benchmark whose oldest
bounds were calibrated on the same grid that now demonstrates compliance, and
CI that depends on a self-hosted Windows/Excel runner and therefore cannot be
independently re-run by outsiders.

**Overall weighted score: 8.6 / 10** (exact 8.625). Detailed rubric below.

---

## 2. Scoring rubric

Each category is scored 0–10 and weighted by its importance to a numerical
library intended for professional use. 10 = state of the art for the problem
domain; 7 = solid professional quality with identifiable gaps; 5 = adequate
but with weaknesses a serious user must work around; below 4 = deficient.

| # | Category | Weight | Score | Weighted |
|---|----------|:-----:|:-----:|:-------:|
| 1 | Numerical correctness & methodology | 20% | **9.5** | 1.90 |
| 2 | Verification & benchmark infrastructure | 15% | **9.5** | 1.43 |
| 3 | Error handling & robustness | 10% | **9.0** | 0.90 |
| 4 | Code quality & style consistency | 10% | **9.0** | 0.90 |
| 5 | Testing | 10% | **8.5** | 0.85 |
| 6 | API design & Excel integration | 10% | **8.5** | 0.85 |
| 7 | Documentation | 10% | **8.0** | 0.80 |
| 8 | Repository, CI & process | 5% | **7.0** | 0.35 |
| 9 | Completeness of scope | 5% | **6.0** | 0.30 |
| 10 | Maintainability & duplication control | 5% | **7.0** | 0.35 |
| | **Overall** | 100% | | **8.6** (exact 8.625) |

---

## 3. Category detail

### 3.1 Numerical correctness & methodology — 9.5/10 (weight 20%)

**Evidence examined.** The five source modules
(`CORE`, `SPECIALFUNCS`, `NORMALFAMILY`, `TFAMILY`, `CONTINUOUS`; ~14,350
lines) expose 64 public worksheet functions over ~56 private/public numerical
kernels. Techniques observed are the correct ones for double-precision work:

- Survival functions computed on the upper tail directly, never as `1 − CDF`,
  so deep-tail probabilities retain relative accuracy where a naive
  subtraction would return zero.
- Log-domain reconstruction of densities where numerator/denominator would
  underflow independently (e.g. lognormal density at magnitudes near 1E−305),
  with `Log1p`/`Expm1`-style primitives in `CORE` to avoid cancellation near
  zero.
- A dedicated stable log-gamma **difference** kernel (`PROB_LogGammaDelta`)
  computing `LogGamma(z+s) − LogGamma(z)` as a single expression, with the
  Lanczos series difference formed directly rather than by subtracting two
  independently evaluated series. `PROB_LogBeta` dispatches between the
  defining three-log-gamma identity (balanced arguments) and this stable
  difference (unbalanced arguments), which eliminates the catastrophic
  cancellation that the identity suffers when the shape ratio is small. This
  is the kind of numerically literate design normally seen in Boost.Math or
  SciPy internals, not in VBA.
- Absolute-vs-relative error metrics chosen per function on mathematical
  grounds (e.g. `StirlingError` and `PROB_LogBeta` are judged on absolute
  error because what propagates through `exp(·)` downstream is the absolute
  log-domain error; inverse functions are judged both on quantile error and on
  the forward-probability residual `|F(x) − p| / min(p, 1−p)`).
- Iterative kernels (incomplete gamma/beta, inverse solvers) return a clean
  non-convergence error outside their validated parameter range instead of
  silently degrading; the validated range is documented.

**Deductions.** −0.5: the accuracy of unbalanced Beta/F is bounded by the
~1E−15 relative accuracy of the underlying double-precision log-gamma
primitives, yielding contracted bounds of roughly 1E−12…2E−10 in that regime
rather than the 5E−15 headline achieved elsewhere. This is honestly
contracted, and reducing it (better coefficient sets, a direct log-beta
approximation, or extended-precision accumulation) is a research task rather
than a defect — but the ceiling exists and is why this category is not 10.

### 3.2 Verification & benchmark infrastructure — 9.5/10 (weight 15%)

This is the repository's most distinctive asset.

- **Machine-readable accuracy contracts** (`benchmark/accuracy_contracts.csv`,
  73 rows) with schema
  `contract_id, function, regime, measure, metric, threshold, domain,
  provenance, status, evidence, notes`. Real function names carry multiple
  regime-scoped contracts (balanced vs unbalanced Beta; validated-range F),
  and a single observation can satisfy several contracts (an inverse quantile
  feeds both a quantile-error and a tail-residual contract).
- **A 685-point main grid** with 50-digit mpmath references, exported from
  Excel via a two-part `hi;lo` encoding that preserves the full IEEE double
  across the CSV round-trip, and analyzed in `Decimal` arithmetic so reported
  errors are not distorted by float cancellation in the harness itself.
- **A release gate** (`compute_errors.py`) with five verdict states (PASS /
  FAIL / KNOWN LIMITATION / CHARACTERIZATION ONLY / PENDING) and granular exit
  codes (1 = blocking, 2 = incomplete, 0 = green), in strict and
  development modes. At the reviewed commit the gate is green: 73/73 PASS.
- **Independent holdout validation**: the most recently calibrated bounds
  (unbalanced Beta, inverse functions, the log-beta absolute bound) were
  confirmed on a 134-point holdout of fresh shapes, ratios and probabilities
  before their provenance was flipped to `validated and frozen`. Margins of
  4.6×–180× were observed.
- **An honest limitations register** (`numerical_limitations.csv`): F with an
  incomplete-beta shape parameter beyond ~1E7 degrades to ~4E−7 and is
  recorded as a known limitation *outside* the contracts, rather than as a
  contract the code cannot meet.
- Focused sub-studies (`logbeta_study/`, `delta_seam_study/`,
  `beta_f_unbalanced/`, `holdout/`) each with their own README, reference
  generator, export macro and analyzer — reproducible end to end.

**Deductions.** −0.5, for one methodological seam: 46 of 73 contracts carry
provenance `source claim` and 19 `measured and frozen`, where the frozen
values were calibrated on the same grid that now demonstrates compliance. Only
the 8 most recent contracts have been through the independent-holdout
discipline. The framework to close this exists (the holdout pattern); it has
simply not yet been applied retroactively to the older bounds.

### 3.3 Error handling & robustness — 9.0/10 (weight 10%)

- A uniform worksheet error contract: public functions return `Variant` and
  fail via `CVErr(xlErrNum)` / `CVErr(xlErrValue)` (287 `CVErr` sites);
  **zero** `MsgBox` calls exist in the source (the 17 textual occurrences are
  documentation lines asserting the policy). 69 structured `On Error GoTo`
  handlers with consistent label conventions.
- Input validation distinguishes domain errors (`#VALUE!`) from numeric-range
  failures (`#NUM!`), with validation messages that state the accepted domain.
- Guard constants (e.g. a parameter-magnitude guard at 1E100) are documented
  as representational bounds distinct from convergence guarantees, and
  non-convergence inside the representational range fails cleanly.

**Deductions.** −1.0: standardization overflow in extreme inputs returns
`#NUM!` even where the sign of the limit is known and a mathematically exact
limiting value (0 or 1) exists. This is a deliberate, documented policy choice
rather than a bug, but it is a place where a user computing e.g. a CDF at an
absurd-but-well-defined input receives an error where a limit value would be
defensible.

### 3.4 Code quality & style consistency — 9.0/10 (weight 10%)

- `Option Explicit` in 6/6 modules. A rigorously uniform house style: banner
  and ruler comment conventions, structured function headers (PURPOSE / INPUTS
  / RETURNS / ERROR POLICY / DEPENDENCIES), aligned declarations, explicit
  `Double` literals (`0#`, `1#`), `vbNullString`, consistent naming tiers
  (`K_STATS_` public, `PROB_` kernels, `Try*` boolean contracts).
- Kernels are single-purpose and composable; the special-function layer is
  cleanly separated from the distribution layer, and cross-module dependencies
  flow in one direction (distributions → special functions → core).
- No dead code or commented-out blocks were observed in the source modules.

**Deductions.** −1.0: some numeric constants (the Lanczos coefficient set) are
duplicated between two functions with a "must match" comment rather than
shared at module scope — a deliberate blast-radius tradeoff, but a divergence
risk; a small number of header `UPDATED` fields lag the code they describe;
one module-level constant name exceeds the declaration-alignment column,
breaking the otherwise-uniform alignment.

### 3.5 Testing — 8.5/10 (weight 10%)

- 76 test procedures, 514 assertions, wired into a single `RunAll` with
  per-module suites, pass/fail accounting, and a CI failure log whose entries
  embed the actual/expected/tolerance diagnostic in the test name.
- Tests include structural mathematical identities that validate kernels
  independently of any reference value (e.g. `Delta(z,0)=0`,
  `Delta(z,1)=Log z`, and the composition law
  `Delta(z,s+t)=Delta(z,s)+Delta(z+s,t)`), plus regression points with
  high-precision literals for previously failing regimes.
- Error-contract tests assert that invalid inputs produce the documented
  worksheet error type, not merely "an error".

**Deductions.** −1.5: coverage is value-centric; there is no randomized or
property-based fuzzing of the input space, and boundary sweeps (denormal
inputs, ±0, argument permutations at guard edges) are present but not
systematic. The suite verifies the *contracted* grid excellently; it explores
outside it modestly.

### 3.6 API design & Excel integration — 8.5/10 (weight 10%)

- Function names state exactly what they compute; the naming discipline is
  strong enough that optional flags that would change semantics have been
  (correctly) rejected in favor of separate names.
- Parameterization deliberately matches Excel's built-ins argument-for-
  argument, including Excel's own inconsistencies, which minimizes migration
  friction for the target user.
- The `Variant`-return / `CVErr` pattern is the correct Excel-native failure
  mode, and survival/inverse-survival pairs give tail-accurate entry points
  Excel lacks.

**Deductions.** −1.5: no discrete distributions (see 3.9), so the API is
incomplete as a general statistics surface; a handful of convenience gaps
(e.g. no vectorized/array entry points) are understandable in VBA but real.

### 3.7 Documentation — 8.0/10 (weight 10%)

- Root README, per-study READMEs, a benchmark README whose contract table is
  generated from the contract file between explicit `BEGIN/END generated`
  markers (with a `--write` refresh mode), CI documentation, contributing/
  security/community files, and a design note recording why a considered
  third numerical regime was dropped after measurement.
- Function headers are effectively reference documentation, including error
  policy and numerical method per function.
- The documentation is candid: it tells regulated-use readers to re-validate
  against an independent oracle, and it documents its own limitations.

**Deductions.** −2.0: contract information still appears in prose in source
headers and (per repository cross-references) wiki pages, which the generated
table cannot keep synchronized — drift between prose claims and the contract
file has occurred before and the structural fix (single-sourcing) currently
covers the benchmark README only. Four public functions
(`ChiSquare_Density`, `F_Density`, `Normal_IntervalProbability`,
`Lognormal_ParametersFromMeanStdDev`) have no benchmark rows, so their header
accuracy language is uncorroborated by the grid.

### 3.8 Repository, CI & process — 7.0/10 (weight 5%)

- GitHub Actions workflow driving a PowerShell harness that opens Excel,
  imports the modules, runs the suite and parses the counts — genuine
  end-to-end regression for a platform that is notoriously hard to automate.
  Issue/PR templates and community-health files are complete.
- **Deductions.** −3.0: the workflow requires a **self-hosted** Windows+Excel
  runner, so no outsider (and no public CI) can independently reproduce a
  green run; compiled Python `__pycache__/*.pyc` artifacts are committed under
  `benchmark/beta_f_unbalanced/`; the two-phase benchmark (Python references,
  manual Excel export, Python analysis) is documented but inherently manual in
  the middle, so the accuracy gate is not exercised by CI — only the unit
  suite is.

### 3.9 Completeness of scope — 6.0/10 (weight 5%)

Continuous coverage is strong: Normal/Lognormal family (23 UDFs), t-family
including Chi-square and F (12), and the general continuous module
(Gamma, Beta, Exponential, Weibull, Uniform — 29), with density / CDF /
survival / inverse / moments largely uniform across families.

**Deductions.** −4.0: **no discrete distributions.** Binomial, Poisson,
geometric, negative binomial, and hypergeometric are absent, which for a
library positioned as a general Excel statistics upgrade is the single
largest functional gap — a user cannot compute a binomial tail with this
library at all, accurately or otherwise. Multivariate and less-common
continuous families (logistic, Laplace, Pareto) are also absent but are
second-order next to the discrete gap.

### 3.10 Maintainability & duplication control — 7.0/10 (weight 5%)

- The benchmark chain is single-sourced (contract → grid → summary → README
  table) with drift cross-checks, which is exactly the right structure.
- **Deductions.** −3.0: accuracy prose in module headers and wiki duplicates
  contract facts outside that chain; the Lanczos constant duplication noted in
  3.4; the high-precision inverse reference helper hard-codes bracketing
  intervals (`(0,1)` for Beta, `(0,1E12)` for F), which will silently limit
  future extreme-probability grid expansion; `__pycache__` in version control.

---

## 4. Findings register

| ID | Severity | Area | Finding | Recommendation |
|----|:--:|------|---------|----------------|
| F-01 | Medium | Scope | No discrete distributions (binomial, Poisson, geometric, negative binomial, hypergeometric). | Highest-value next feature; the existing kernel/test/benchmark discipline transfers directly. |
| F-02 | Medium | Benchmark methodology | 65 of 73 contracts predate the independent-holdout discipline; their bounds were calibrated on the compliance grid itself. | Extend the holdout pattern retroactively (one consolidated fresh grid covering the older families), then flip provenance. |
| F-03 | Medium | CI | Accuracy gate not run in CI; unit suite requires a self-hosted Excel runner no outsider can reproduce. | Run `compute_errors.py` (strict) as a CI step on the committed grid — it is pure Python and needs no Excel; document the self-hosted constraint prominently. |
| F-04 | Low | Coverage | Four public UDFs have no benchmark rows (`ChiSquare_Density`, `F_Density`, `Normal_IntervalProbability`, `Lognormal_ParametersFromMeanStdDev`). | Add grid rows; all four have trivial mpmath references. |
| F-05 | Low | Hygiene | `__pycache__/*.pyc` committed under `benchmark/beta_f_unbalanced/`. | Delete and add `__pycache__/` to `.gitignore`. |
| F-06 | Low | Maintainability | Contract facts duplicated in module-header prose and wiki outside the generated chain. | Replace numeric claims in prose with pointers to `accuracy_contracts.csv`. |
| F-07 | Low | Maintainability | Lanczos coefficients duplicated across two kernels with a "must match" comment. | Promote to module-scope `Private Const` in a future touch of that module. |
| F-08 | Low | Benchmark tooling | Fixed inverse-search brackets in the reference helper limit future extreme-tail grids. | Parameterize the bracket or derive it from the target probability when such grids are added. |
| F-09 | Info | Documentation | A few header `UPDATED` fields lag recent changes. | Refresh opportunistically. |
| F-10 | Info | Robustness (policy) | Overflowing standardization returns `#NUM!` where the limit value is known. | Documented policy; acceptable as-is. Revisit only if users report friction. |

No high-severity findings. No numerical-correctness defects were identified at
the reviewed commit: the accuracy gate is green (73/73), the one known
numerical limitation is registered honestly, and spot-reading of the kernels
found the algorithms matching their documented methods.

---

## 5. Comparative context

Judged against its ecosystem: most published VBA statistical code has no
tests, no error contract, and no accuracy measurement whatsoever; this
repository has 514 assertions, a machine-readable regime-scoped accuracy
contract, and holdout-validated bounds. Judged against mainstream numerical
libraries (SciPy, Boost.Math), the *verification process* here is comparable
or stricter per function covered — the gap is breadth (families covered,
platform reach), which is inherent to the choice of VBA as the delivery
vehicle, not to the engineering.

## 6. Conclusion

**8.6 / 10.** A small library executed to an unusually high standard. Its
distinguishing feature is epistemic honesty: accuracy claims are contracts
with provenance, failures are gate states rather than footnotes, and the one
thing the library cannot currently do to its own standard is recorded in a
limitations register instead of prose hedging. The clear priorities are the
discrete-distribution gap (F-01), retroactive holdout validation of the older
bounds (F-02), and putting the pure-Python accuracy gate into CI (F-03).
Everything else is polish.

*This review was produced by Claude Fable 5 against commit `b0c4b2f` and
reflects the repository state at that commit only.*
