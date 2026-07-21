# Independent Code Review — VBA-PROBABILITY-DISTRIBUTIONS

| | |
|---|---|
| **Repository** | `danielep71/VBA-PROBABILITY-DISTRIBUTIONS` |
| **Commit reviewed** | `ecf3e45` |
| **Review date** | 2026-07-21 |
| **Reviewer** | Claude Fable 5 (Anthropic) — automated review with human-readable rationale |
| **Scope** | All VBA source modules, test suite, benchmark and verification infrastructure, CI, documentation, repository hygiene |
| **Method** | Full static read of every `.bas` module; metric extraction (API surface, assertion counts, error-contract usage); cross-checking documentation claims against machine-readable artifacts; execution of the accuracy-gate tooling to confirm it runs, evaluates every active contract, and returns the exit codes it documents |

---

## 1. Executive summary

This is a VBA library of probability-distribution worksheet functions — the
Normal family, the Student-t family (including Chi-square and F), the general
continuous families (Gamma, Beta, Exponential, Weibull, Uniform), a discrete
family (Binomial, Poisson, Geometric), and the special-function kernels beneath
them — whose stated purpose is to exceed Excel's built-in accuracy,
particularly in deep tails where Excel's functions collapse to zero.

The headline result of this review: **the verification process is now
essentially complete end to end.** Accuracy is contracted per function *and per
parameter regime* in a machine-readable file (76 contracts), measured against
high-precision references on a 724-point grid, gated by a script whose strict
mode cannot pass while any active contract is unevaluated, and — for 71 of the
76 contracts — validated on independent, off-grid holdout data before being
frozen. Where fresh data showed a bound did not generalize, the repository did
the honest thing in every case: thresholds were relaxed to measured levels,
domains were restricted to where the claim is actually true, and the residual
behavior was recorded in a limitations register with its own measurement study.
Two functions whose extreme-parameter path could silently return inaccurate
values are now guarded by a *measured* rejection envelope rather than a hope.

The engineering discipline is far above what is typical for the VBA ecosystem,
and the verification infrastructure is above what is typical for open-source
numerical libraries in any language. The main remaining weaknesses are scope
(the discrete family covers three of the five canonical distributions and is
the only family not yet under accuracy contracts), a benchmark whose middle
step (the Excel export) is inherently manual, and CI whose unit-test half
requires a self-hosted Windows/Excel runner that outsiders cannot reproduce —
though the accuracy gate itself now runs on a public runner that anyone can
verify.

**Overall weighted score: 9.1 / 10** (exact 9.05). Detailed rubric below.

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
| 4 | Code quality & style consistency | 10% | **9.5** | 0.95 |
| 5 | Testing | 10% | **9.0** | 0.90 |
| 6 | API design & Excel integration | 10% | **9.0** | 0.90 |
| 7 | Documentation | 10% | **8.5** | 0.85 |
| 8 | Repository, CI & process | 5% | **8.5** | 0.43 |
| 9 | Completeness of scope | 5% | **7.5** | 0.38 |
| 10 | Maintainability & duplication control | 5% | **8.5** | 0.43 |
| | **Overall** | 100% | | **9.1** (exact 9.05) |

---

## 3. Category detail

### 3.1 Numerical correctness & methodology — 9.5/10 (weight 20%)

**Evidence examined.** The six source modules
(`CORE`, `SPECIALFUNCS`, `NORMALFAMILY`, `TFAMILY`, `CONTINUOUS`, `DISCRETE`;
~19,100 lines) expose 88 public worksheet functions over ~91 private/public
numerical kernels. Techniques observed are the correct ones for
double-precision work:

- Survival functions computed on the upper tail directly, never as `1 − CDF`,
  so deep-tail probabilities retain accuracy where a naive subtraction would
  return zero. Notably, the *relative* accuracy of the shared standard-normal
  survival core in the deep tail has been measured rather than assumed: a
  dedicated z-sweep study locates exactly where the tight relative bound stops
  holding (~z ≈ 2.75–3.25), the contracts are domain-restricted to that
  measured region, and the deeper-tail behavior (absolute error steady at
  ~1E−17 while relative error grows to ~5E−10 by z = 6) is registered as a
  characterized limitation. This is the difference between claiming tail
  accuracy and knowing precisely which kind of tail accuracy one has.
- Log-domain reconstruction wherever independent factors could overflow or
  underflow while the final result is representable: the lognormal
  variance/StdDev are formed as a *single* logarithm through a dedicated
  `LogExpm1` primitive with one final exponential, so a finite moment is never
  lost to an intermediate `Exp` overflow, and zero is returned only when the
  moment genuinely underflows.
- A stable log-gamma **difference** kernel computing
  `LogGamma(z+s) − LogGamma(z)` with the Lanczos series difference formed
  directly rather than by subtracting two independently evaluated series.
  `PROB_LogBeta` dispatches between the defining three-log-gamma identity
  (balanced arguments) and this stable difference (unbalanced), eliminating
  the catastrophic cancellation the identity suffers at small shape ratios.
  The Lanczos coefficient set is held once, at module scope, and shared by
  both consumers — the single-source-of-truth arrangement one wants for
  constants that must remain bit-identical.
- The F distribution enforces a **measured accuracy envelope**: both degrees
  of freedom must stay within 1E5, a boundary located by sweeping the
  underlying incomplete-beta shape parameter in *both* orientations
  (one-large and both-large — the latter binds, degrading earlier) and finding
  where the continued fraction begins returning values outside contract while
  still satisfying its local convergence test. Beyond the envelope the public
  CDF/survival/inverse return a clean worksheet error rather than a silently
  inaccurate number; the closed-form density is deliberately exempt. Rejecting
  a computable-but-unreliable region on measured grounds is a governed-library
  decision made correctly.
- Absolute-vs-relative error metrics chosen per function on mathematical
  grounds (log-domain kernels judged on absolute error because that is what
  propagates through `exp(·)`; inverse functions judged both on quantile error
  and on the forward-probability residual `|F(x) − p| / min(p, 1−p)`).
- The discrete family routes large-n Binomial and large-mean Poisson tails
  through the already-validated regularized incomplete beta/gamma kernels
  rather than reimplementing summation, so the discrete CDFs inherit the
  accuracy of the continuous machinery; log-PMF entry points are provided for
  work in the log domain.

**Deductions.** −0.5: the accuracy of unbalanced Beta/F remains bounded by the
~1E−15 relative accuracy of the underlying double-precision log-gamma
primitives, yielding contracted bounds of roughly 1E−12…2E−10 in that regime
rather than the 5E−15 headline achieved elsewhere; and the deep-tail survival
path is honest about relative degradation rather than free of it. Both
ceilings are measured, contracted, and documented — reducing them (better
coefficient sets, extended-precision accumulation, a dedicated tail expansion)
is a research task rather than a defect, but the ceilings exist.

### 3.2 Verification & benchmark infrastructure — 9.5/10 (weight 15%)

This remains the repository's most distinctive asset, and its historical seam
has been closed.

- **Machine-readable accuracy contracts** (`benchmark/accuracy_contracts.csv`,
  76 rows) with schema
  `contract_id, function, regime, measure, metric, threshold, domain,
  provenance, status, evidence, notes`. Functions carry multiple regime-scoped
  contracts (balanced vs unbalanced Beta; envelope-restricted F; central-region
  vs characterized-tail survival), and a single observation can satisfy several
  contracts.
- **Provenance is now dominated by out-of-sample validation: 71 of 76
  contracts are `validated and frozen`**, confirmed on independent holdout
  grids of fresh shapes, ratios and probabilities that were never used to
  calibrate the thresholds. The consolidated older-family holdout (116 fresh
  points, 63 contracts) is the kind of retroactive validation most projects
  never do — and it visibly did its job: it caught one threshold that had been
  overfit to its calibration grid by three orders of magnitude, two thresholds
  a few percent too tight, one domain overclaim on the survival tail, and two
  contracts referencing functions that no longer exist. Every one was resolved
  by measurement (relax, restrict, or delete), not by quietly re-fitting.
- **A 724-point main grid** with high-precision references, exported from
  Excel via a two-part `hi;lo` encoding that preserves the full IEEE double
  across the CSV round-trip, and analyzed in `Decimal` arithmetic so reported
  errors are not distorted by float cancellation in the harness itself.
- **A hardened release gate** (`compute_errors.py`) with five verdict states
  and granular exit codes, in strict and development modes. The gate cannot
  pass while an active contract is unevaluated: if the high-precision
  reference helper is missing or broken, the affected contracts become
  blocking PENDING items with the exact import failure printed, and a
  committed self-test (`test_gate_degradation.py`) locks that behavior in.
  At the reviewed commit the strict gate is green with every active contract
  evaluated.
- **Independent-oracle cross-checking is an artifact, not an assertion**:
  `cross_check_scipy.py` regenerates `cross_check_scipy.md`, which confirms
  242 reference points against SciPy where SciPy is reliable and honestly
  marks the 8 extreme points beyond SciPy's range as resting on the
  reference's own 50-vs-80-digit self-consistency instead.
- **An honest limitations register** (`numerical_limitations.csv`) with two
  entries: the extreme-shape incomplete-beta limitation (now *mitigated* by
  the F envelope) and the characterized survival-tail relative degradation,
  each pointing at its measurement study.
- Eight focused sub-studies, each with its own reference generator, export
  macro, grid and analyzer — reproducible end to end, including the boundary
  studies that produced the F envelope and the survival-tail domain.

**Deductions.** −0.5: five contracts (the recently added density/moment
helpers) remain `measured provisional` — measured on the main grid but not yet
through the holdout discipline; and the benchmark's middle step (the Excel
export that fills `observed_vba`) is inherently manual, so CI verifies that
the *committed* evidence meets the contracts, not that the committed evidence
matches the current VBA. The first is a small queue item; the second is
structural to the platform and is documented as such.

### 3.3 Error handling & robustness — 9.0/10 (weight 10%)

- A uniform worksheet error contract: public functions return `Variant` and
  fail via `CVErr(xlErrNum)` / `CVErr(xlErrValue)` (443 `CVErr` sites across
  the library); **zero** `MsgBox` calls exist in executable code. 93
  structured `On Error GoTo` handlers with consistent label conventions
  (`Err_Handler`, `Fail_Num`).
- Input validation distinguishes domain errors (`#VALUE!`) from numeric-range
  failures (`#NUM!`), with validation messages that state the accepted domain
  through the optional `Status` channel.
- Guard constants are documented as representational bounds distinct from
  accuracy guarantees, and the one region where "accepted" did not imply
  "accurate" — extreme-df F — is now closed by measured rejection rather than
  left as a caveat.

**Deductions.** −1.0: standardization overflow in extreme inputs returns
`#NUM!` even where the sign of the limit is known and a mathematically exact
limiting value (0 or 1) exists. This is a deliberate, documented policy choice
rather than a bug, but it is a place where a user computing a CDF at an
absurd-but-well-defined input receives an error where a limit value would be
defensible.

### 3.4 Code quality & style consistency — 9.5/10 (weight 10%)

- `Option Explicit` in 7/7 modules (six source, one test). A rigorously
  uniform house style: banner and ruler comment conventions, structured
  function headers (PURPOSE / INPUTS / RETURNS / ERROR POLICY / DEPENDENCIES /
  UPDATED), aligned declarations, explicit `Double` literals (`0#`, `1#`),
  `vbNullString`, consistent naming tiers (`K_STATS_` public, `PROB_` kernels,
  `Try*` boolean contracts).
- Kernels are single-purpose and composable; cross-module dependencies flow in
  one direction (distributions → special functions → core), and the discrete
  module consumes the shared incomplete-beta/gamma kernels rather than keeping
  private duplicates.
- Numerical constants that must remain identical are held once at module
  scope; header `UPDATED` fields are current with the code they describe.
- No dead code or commented-out blocks were observed in the source modules.

**Deductions.** −0.5: the high-precision reference helper `_ibeta.py` is
duplicated verbatim in three benchmark study folders — a deliberate
self-containment tradeoff, but file-level duplication with drift risk of
exactly the kind the module-scope constants were introduced to prevent.

### 3.5 Testing — 9.0/10 (weight 10%)

- 94 test procedures, 599 assertions, wired into a single `RunAll` with
  per-module suites, pass/fail accounting, and a CI failure log whose entries
  embed the actual/expected/tolerance diagnostic in the test name.
- Tests include structural mathematical identities that validate kernels
  independently of any reference value (delta identities and composition laws,
  including in the both-subunit regime), regression points with
  high-precision literals for previously failing regimes, and *branch-boundary*
  regressions that pin both sides of internal numerical crossovers (e.g. the
  `v ≥ 709` switch inside the lognormal-moment reconstruction) so the
  overflow-avoiding path cannot be silently regressed.
- The F envelope is tested from below, at, and above the boundary, including
  a test asserting that the closed-form density is deliberately *not*
  restricted.
- The role split between test tiers is documented explicitly: the VBA suite is
  a fast deterministic regression and public-contract smoke test with broad
  tolerances *by design*; the external benchmark is the measured
  high-precision gate. A Python-side self-test guards the gate's own
  failure-mode behavior.

**Deductions.** −1.0: coverage is value-centric; there is no randomized or
property-based fuzzing of the input space, and boundary sweeps (denormals,
±0, argument permutations at guard edges) are present but not systematic. The
suite verifies the contracted surface excellently; it explores outside it
modestly.

### 3.6 API design & Excel integration — 9.0/10 (weight 10%)

- Function names state exactly what they compute; the naming discipline is
  strong enough that optional flags that would change semantics have been
  (correctly) rejected in favor of separate names.
- Parameterization deliberately matches Excel's built-ins argument-for-
  argument, including Excel's own inconsistencies (Exponential takes a rate,
  Gamma and Weibull a scale), which minimizes migration friction for the
  target user.
- The `Variant`-return / `CVErr` pattern is the correct Excel-native failure
  mode; survival and inverse-survival pairs give tail-accurate entry points
  Excel lacks; the discrete family adds log-PMF entry points for log-domain
  work.

**Deductions.** −1.0: the discrete surface covers three of the five canonical
distributions (negative binomial and hypergeometric are absent), and there are
no vectorized/array entry points — understandable in VBA but real.

### 3.7 Documentation — 8.5/10 (weight 10%)

- Root README (which documents the discrete family alongside the continuous
  ones), per-study READMEs, a benchmark README whose contract table is
  generated from the contract file between explicit markers, an
  honestly-scoped reference-integrity section tied to the regenerable SciPy
  cross-check artifact, CI documentation, contributing/security/community
  files, and committed summary artifacts for every validation campaign
  (holdout summaries, boundary-study summaries).
- Function headers are effectively reference documentation, including error
  policy and numerical method per function; module headers state
  regime-specific accuracy by *pointing at* the contract file rather than
  duplicating it.
- The documentation is candid: it tells regulated-use readers to re-validate
  against an independent oracle, and it documents its own limitations with
  measurement studies attached.

**Deductions.** −1.5: the survival-function headers emphasize far-tail
*stability* (true in absolute terms) without cross-referencing the registered
relative-accuracy limitation and its measured domain, so a header-only reader
gets a rosier picture than the contract file gives; some residual accuracy
prose in older headers and wiki pages sits outside the generated
single-source chain.

### 3.8 Repository, CI & process — 8.5/10 (weight 5%)

- Two workflows with correctly divided labor: a **public-runner accuracy
  gate** (`ubuntu-latest`, pure Python) that runs the strict gate on the
  committed grid *and* the gate's own degradation self-test on every benchmark
  change — meaning any outsider can watch the accuracy contracts being
  enforced — and the Excel/VBA regression harness on a self-hosted Windows
  runner that opens Excel, imports the modules and runs the suite.
- Repository hygiene is clean: no compiled artifacts in version control, a
  Python-aware `.gitignore`, issue/PR templates and community-health files
  complete.

**Deductions.** −1.5: the Excel half of CI still requires a self-hosted
runner no outsider can reproduce, and the two-phase benchmark retains its
manual Excel-export middle — both inherent to choosing VBA as the delivery
vehicle, and both documented, but both real constraints on independent
reproducibility of the *full* chain.

### 3.9 Completeness of scope — 7.5/10 (weight 5%)

Continuous coverage is strong and uniform: Normal/Lognormal family (23 UDFs),
t-family including Chi-square and F (12), the general continuous module
(Gamma, Beta, Exponential, Weibull, Uniform — 29), and a discrete module
(Binomial, Poisson, Geometric — 24 UDFs) each offering PMF/density, CDF,
survival, inverse, moments, and (discrete) log-PMF.

**Deductions.** −2.5: the discrete family stops at three distributions —
negative binomial and hypergeometric are absent, so common tasks like
over-dispersed count models and finite-population sampling remain out of
reach. Multivariate and less-common continuous families (logistic, Laplace,
Pareto) are also absent but are second-order next to completing the discrete
set.

### 3.10 Maintainability & duplication control — 8.5/10 (weight 5%)

- The benchmark chain is single-sourced (contract → grid → summary → README
  table) with drift cross-checks; shared numerical constants are held once at
  module scope; the reference helper's inverse-search brackets are derived
  from the target probability rather than hard-coded, so future extreme-tail
  grids will not be silently clipped.
- **Deductions.** −1.5: `_ibeta.py` triplication across study folders (see
  3.4); residual accuracy prose in headers/wiki outside the generated chain
  (see 3.7).

---

## 4. Findings register

| ID | Severity | Area | Finding | Recommendation |
|----|:--:|------|---------|----------------|
| F-01 | Medium | Scope | Discrete family incomplete: negative binomial and hypergeometric absent. | Highest-value next feature; the existing kernel/test discipline transfers directly (negative binomial reuses the incomplete-beta path already wired for binomial). |
| F-02 | Medium | Benchmark coverage | The discrete family is tested (structural identities, tail regressions) but carries no accuracy contracts or grid rows — the only family outside the contract regime. | Add discrete grid rows and contracts, especially for the large-n/large-mean incomplete-beta/gamma regimes, then holdout-validate and freeze like the rest. |
| F-03 | Low | Benchmark methodology | Five contracts remain `measured provisional` (density/moment helpers measured on the main grid only). | Fold them into the next holdout campaign and flip provenance. |
| F-04 | Low | Maintainability | `_ibeta.py` duplicated verbatim in three study folders. | Single-source it (one shared module imported by each study), accepting the loss of strict per-study self-containment. |
| F-05 | Low | CI | Excel/VBA suite requires a self-hosted runner; the benchmark's Excel-export middle is manual. | Inherent to the platform; keep the constraint documented prominently. The public accuracy gate already covers the reproducible half. |
| F-06 | Low | Documentation | Survival headers emphasize far-tail stability without cross-referencing the registered relative-accuracy limitation and its measured z-domain. | Add one pointer line per affected header to `SurvivalTailRel` / the boundary study. |
| F-07 | Info | Robustness (policy) | Overflowing standardization returns `#NUM!` where the limit value is known. | Documented policy; acceptable as-is. Revisit only if users report friction. |
| F-08 | Info | Documentation | Residual contract facts in wiki/prose outside the generated single-source chain. | Replace numeric claims with pointers opportunistically. |

No high-severity findings. No numerical-correctness defects were identified at
the reviewed commit: the strict accuracy gate is green with every active
contract evaluated, 71 of 76 contracts are frozen against independent holdout
evidence, both registered limitations carry their own measurement studies, and
spot-reading of the kernels found the algorithms matching their documented
methods.

---

## 5. Comparative context

Judged against its ecosystem: most published VBA statistical code has no
tests, no error contract, and no accuracy measurement whatsoever; this
repository has 599 assertions, a machine-readable regime-scoped accuracy
contract, holdout-validated frozen bounds, a release gate that refuses to pass
with an unevaluated contract, and a public CI run of that gate anyone can
inspect. Judged against mainstream numerical libraries (SciPy, Boost.Math),
the *verification process* here is comparable or stricter per function covered
— including practices (per-regime contracts with provenance, retroactive
out-of-sample validation, measured rejection envelopes, characterized-tail
domain restriction) that the mainstream libraries themselves do not uniformly
apply. The gap is breadth — families covered and platform reach — which is
inherent to the choice of VBA as the delivery vehicle, not to the engineering.

## 6. Conclusion

**9.1 / 10.** A small library executed to an unusually high standard, whose
distinguishing feature is that its epistemic honesty is now *load-bearing*:
accuracy claims are contracts with out-of-sample provenance, the gate
structurally cannot certify what it has not evaluated, boundaries between
"accurate", "accurate in a measured region", and "rejected by policy" are all
measured rather than asserted, and the two things the library cannot do to its
own tightest standard are registered limitations with attached studies rather
than prose hedging. The clear priorities are completing the discrete family
(F-01) and bringing it under the same contract regime as everything else
(F-02); everything else is polish.

*This review was produced by Claude Fable 5 against commit `ecf3e45` and
reflects the repository state at that commit only.*
