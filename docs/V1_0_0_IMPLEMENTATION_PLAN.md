# v1.0.0 implementation and release-readiness plan

Status date: 2026-08-30
Repository: danielep71/VBA-PROBABILITY-DISTRIBUTIONS  
Repository HEAD: 0dd748884599d4d0da815cb53eeceb13efd51f05
Numerical source baseline: bde92dd7037e4fde05e620745a1c54b0cbc3a261  
Milestone: v1.0.0

## Executive decision

v1.0.0 is not release-ready.

The project architecture and evidence framework are strong, but the release currently has six open P1 numerical issues plus the P1 release-certification tracker, stale source-bound accuracy observations, an unavailable self-hosted Excel result for the current source, an audit-baseline set of 36 unclaimed main-grid rows whose live count must thereafter be generated, unbound independent holdout observations, stale public assurance metrics, no changelog/tag/release, and a security policy that describes an unreleased tag as already stable.

The corrected v1.0.0 milestone contains 26 issues: 12 open and 14 closed. The implementation order below is dependency-driven and preserves the frozen numerical contracts. Two additional open hardening issues belong to v1.01: #7 and #32.

### Scope decision: do not defer #13/#14

A documentation-only descoping option was considered: register the positive-subnormal and hard-underflow regimes in `numerical_limitations.csv`, move #13/#14 to v1.1.0, and publish v1.0.0 without correcting them. It is rejected.

#11 includes accepted-domain calls that return a plausible silent wrong value, including `0` where the representable result is approximately `0.912`, without raising an error. This is the project's highest-severity defect class. Merely documenting that behavior would contradict the first stable release's measured-accuracy claim and create a credibility problem, not just a numerical limitation.

The only defensible fallback would be a source change that explicitly rejects the affected regime with the documented worksheet error and narrows the supported domain, accompanied by matching tests, contracts, and documentation. That is not the selected v1.0.0 scope. #13 and #14 remain release blockers, and #11 must close before publication.

### Scheduling decision: decouple #22's detector from its evidence study

The original plan made all of #22 a predecessor of #23 so a claim-completeness detector would exist before later value-changing commits. That coupling is too broad. #23 does not numerically depend on the new inverse contracts, median study, dyadic bridges, or large-shape Chi-square reference generator; making those prerequisites moves the longest and least-certain estimate onto the numerical critical path.

The revised schedule splits the work:

- Phase 0 lands a cheap canonical coverage-delta guard. It records the 36-row audit-baseline identity set, reports the current unresolved subset and count as generated `KNOWN COVERAGE DEBT`, and fails on any unauthorized new or changed unclaimed row. Strict zero-missing mode remains the release authority.
- #23 then proceeds as soon as #29's binding mechanism and a controlled Excel route are ready.
- #22's oracle, classifier, zero-threshold, and fitting/holdout preparation continues in parallel with the numerical chain.
- Final #22 measurement occurs on the release-candidate numerical source after #14, then the temporary debt fingerprint is deleted and strict mode must report zero missing rows.

This deliberately accepts a later inverse-evidence wave. Where practical it is combined with #31's mandatory release-candidate Excel export, so it does not add a separate certification cycle. The trade removes the new reference generator from the #23 -> #13 -> #14 critical path without weakening the final v1.0.0 gate.

## Audit basis

### Source and repository

- repository HEAD: 0dd748884599d4d0da815cb53eeceb13efd51f05
- latest numerical source baseline: bde92dd7037e4fde05e620745a1c54b0cbc3a261
- remote branches: main only
- Git tags: none
- GitHub Releases: none
- branch protection: main is unprotected
- required production modules: six tracked .bas files under src
- consolidated regression module: tests/M_STATS_PROBDIST_TEST.bas
- example workbook: examples/STATS-Distributions demo.xlsm
- present governance files: LICENSE, CONTRIBUTING.md, CODE_OF_CONDUCT.md, SECURITY.md, issue templates, PR template
- missing first-release file: CHANGELOG.md

### Verified numerical evidence

- latest verified real-Excel regression: 902 / 902 PASS
- environment: Excel 16.0 build 20131, 64-bit
- #23 adds seven assertions, so 909 is the next expected count; it is not yet verified
- registry rows: 166
- registry status: 165 active, 1 characterization-only
- probability grid: 2,088 rows
- main-grid rows: 1,732
- rows matching at least one contract key: 2,052
- unclaimed main-grid rows at the bde92dd audit baseline: 36; every later current-debt count is generated from the grid and disposition data
- independent holdout grid: 559 observation rows
- independent holdout summary: 80 contract verdicts, 80 PASS / 0 FAIL; the badge is numerically correct but ambiguous, and the observations are not source-bound

### Live CI state

- Accuracy Gate runs 166 through 171: failed only with the frozen two-file stale-evidence signature; every Phase 0 commit preserved it
- the #13 analyzer and its fitting/holdout claims passed before the provenance failure
- Excel VBA Regression run 95: cancelled because the self-hosted Excel runner was unavailable; it produced no evidence
- workflow logs warn that older action majors target deprecated Node 20 and are being forced onto Node 24

### Phase 0 known-red invariant

`main` is intentionally red from the #23 source merge until Phase 1 performs a fresh Excel export and rewrites the source binding. A red top line is therefore not an adequate Phase 0 regression signal. Under the workflow-equivalent Python 3.12 plus mpmath environment, every Phase 0 commit must retain this exact ordered signature:

| Order | Check | Required Phase 0 result | Required cause |
| ---: | --- | --- | --- |
| 1 | evaluator unit tests | PASS | — |
| 2 | manifest unit tests | PASS | — |
| 3 | strict accuracy gate | FAIL | `STALE EVIDENCE`, exactly 2 mismatches |
| 4 | gate blocks without references | FAIL | same stale-evidence cause and same 2 mismatches |
| 5 | source claims vs registries | PASS | — |
| 6 | holdout analyzer tests | PASS | — |
| 7 | independent holdout | PASS | numerical verdict only; provenance remains blocked by the strict gate |

The mismatch set is exactly:

- `src/M_STATS_PROBDIST_SPECIALFUNCS.bas`;
- `tests/M_STATS_PROBDIST_TEST.bas`.

Every Phase 0 commit compares the ordered tuple `(label, result, failure class, mismatch count, mismatch paths)` with this signature. Any difference is a regression despite the already-red top line. Runs 166 through 171 preserve the invariant.

#29 keeps the existing stale main-manifest check first during this window while exercising the holdout verifier through blocking fixtures. Phase 1 exports both observation sets and writes both truthful bindings atomically. The waiver expires immediately afterward and all seven checks must PASS.

## Backlog reconciliation performed

### Corrected or materially updated

| Issue | Correction |
| --- | --- |
| #11 | Replaced placeholder dependencies with #12, #13, and #14; added the incomplete-gamma prerequisites and exact parent closure rule. |
| #13 | Replaced superseded hypotheses with the measured per-family candidate policy, then added mandatory post-#23 crossover revalidation before any cutoff is wired. |
| #14 | Replaced placeholder links; made the log-quantile architecture, final-scale reconstruction, and prerequisites explicit. |
| #17 | Reopened after the closure audit, then narrowed to the v1.0.0 non-destructive/fail-closed generator safeguards. Full reconstruction moved to #32 under v1.01. |
| #22 | Recounted the audit-baseline grid; selected real contracts for its 36 unclaimed rows; confirmed that production F accepts df through 1E10, making the <=1E5 `PROB_F_ValidateEnvelope` enforcement text false on four registry rows; found eight mis-scoped F main rows, eight mis-scoped holdout rows at df 2E5/5E5, five additional exact-zero Student-t medians, and three actual duplicate F inputs; preregistered gap-free regimes, bridge evidence, and exact-zero tooling; split the cheap coverage guard from the release-candidate evidence wave; and accepted the df = 1E8 gmpy2/MPFR substitution after 69/69 quadrature and gmpy2 convergence with Rmpfr agreement on all 46 feasible points. |
| #23 | Retitled to the actual off-grid Stirling-prefactor root cause; recorded the merged implementation and evidence still required. |
| #24 | Specified direct lower-region Q from a pre-exponentiation LogP and made #23 a hard prerequisite. |
| #26 | Reframed from a presumed separate Lentz defect to a mandatory post-#23 revalidation gate. |
| #28 | Corrected v1.1.0 to v1.0.0, refreshed all metrics, and specified offline generated/fail-closed authorities. |

### Closed

| Issue | Reason |
| --- | --- |
| #20 | The exact binary64 Stirling boundary tests already exist in the consolidated harness and are included in the verified 902 / 902 Excel run. |

### Added

| Issue | Purpose |
| --- | --- |
| #29 | Split the Excel-free provenance mechanism from the first truthful binding, which is written and verified during the #23 holdout export. |
| #30 | Move workflows from deprecated Node 20 action majors to supported Node 24 majors. |
| #31 | Own repository preparation, final certification, tagging, and publication of v1.0.0. |
| #32 | Own complete post-v1.0.0 grid reconstruction, declared origins, reference/observation separation, and byte-stable regeneration under v1.01. |

### Deleted

No issue was deleted.

The parent #11 remains useful, and #26 must remain separate until the shared-prefactor fix is remeasured. Deleting or merging #26 now would erase a release-blocking uncertainty rather than resolve it.

## Release-readiness assessment

| Area | State | Evidence / blocker | Owning issue |
| --- | --- | --- | --- |
| Off-grid Gamma prefactor | Implementation landed; evidence incomplete | Current source has no real-Excel/export/holdout closure evidence | #23 |
| Lower incomplete-gamma survival | Blocked | Lower branch still forms Q = 1 - P | #24 |
| Positive subnormal forward paths | Blocked | Constants/classifier exist but have no production callers; pre-#23 cutoffs require revalidation | #13, #11 |
| Inverse representability | Blocked | Unit-scale quantile can underflow before final rescaling | #14, #11 |
| Continued-fraction Q | Unknown after shared fix | Old breach may move or disappear after #23 | #26 |
| Exact Stirling boundary regression | Complete | Permanent arithmetic-built Excel tests; 902 / 902 verified | #20 |
| Main-grid claim completeness | Blocked | 36 rows produced no verdict at the audit baseline; the current count must be generated; Student-t `all` overclaims the measured large-df rows; F regime metadata is stale and three numerical inputs are duplicated across regimes | #22 |
| Main-grid provenance | Correctly stale | Manifest rejects the changed current source | #23 then final export |
| Holdout provenance | Blocked | 559 rows producing 80 contract verdicts are not bound to current source | #29 |
| Grid-regeneration safety | Blocked | The documented generator can still overwrite the combined grid and blank observations by default | #17 |
| README assurance | Blocked | 835 / 161 / 1 905 / zero-known-defect claims are stale | #28 |
| CI action runtime | Cleanup required | Node 20 deprecation warnings | #30 |
| Release documentation | Blocked | No changelog; SECURITY.md prematurely says v1.0.0 is stable | #31 |
| Tag and GitHub Release | Not started | Neither exists | #31 |
| Core architecture | Preserve | No broad redesign required | all numerical issues |
| Governance baseline | Strong | License, conduct, contributing, security, templates present | #31 final audit |

Full grid reconstruction is deliberately not a v1.0.0 blocker. #32 owns that v1.01 programme after #17 makes the current tooling safe.

## Dependency graph

~~~mermaid
flowchart TD
    M29["#29 Provenance mechanism"] --> E23["#23 Excel evidence wave"]
    G22["#22 Coverage-delta guard"] --> E23
    P22["#22 Oracle and study preparation"] --> C22["#22 Final inverse evidence"]
    E23 --> C29["#29 Truthful binding"]
    C29 --> I23["#23 Closure"]
    I23 --> R13["Revalidate #13 cutoffs"]
    I23 --> I24["#24 Direct lower Q"]
    I23 --> I26["#26 Revalidate CF-Q"]
    R13 --> D13["#13 Density and CDF"]
    I24 --> S13["#13 Survival recovery"]
    D13 --> I13["#13 Integrated closure"]
    S13 --> I13
    I13 --> I14["#14 Log inverse reconstruction"]
    I14 --> I11["#11 Parent closure"]
    I11 --> C22
    C22 --> I28["#28 Generated README"]
    C29 --> I28
    I30["#30 Action upgrades"] --> I31["#31 Release certification"]
    I11 --> I31
    I26 --> I31
    I28 --> I31
~~~

## Governing implementation rules

1. Work directly on main with focused, atomic commits until v1.0.0, matching the maintainer's single-branch policy.
2. Never relax a frozen accuracy contract to make a defect pass.
3. Establish a minimal counterexample and exact source path before changing numerical source.
4. Commit predicted movement and untouched holdout design before the source edit.
5. Keep Gamma X / ScaleParam and Chi-square 0.5 * X dispatch policies separate.
6. Keep Loader/deviance near the distribution body; use log-domain arithmetic only in the measured low-significand/hard-underflow regime.
7. Keep distribution policy outside CORE. Add only general arithmetic primitives to CORE when independently justified.
8. Build subnormal test values arithmetically; do not use handwritten subnormal VBA literals.
9. Preserve Try-contract and worksheet-error semantics unless the issue demonstrates a specific contract defect.
10. Re-export evidence from real Excel after every value-changing .bas edit and before writing provenance.
11. Treat the pre-#23 positive-ratio cutoffs as candidates: rerun the frozen fitting arm and untouched holdout after #23, then retain or rederive the constants before production wiring.
12. Land #29's mechanism and #22's cheap coverage-delta guard in Phase 0. #23 may then proceed without waiting for #22's reference generator or median study. Build those in parallel, prove the Chi-square generator feasible before Phase 2, and create final dispositions for every row in the recorded audit-baseline debt set on the release-candidate source after the last inverse-affecting numerical change.
13. Close an issue only after fitting, untouched holdout, permanent VBA regression, full Excel run, strict gate, and unrelated-regression checks pass.
14. Use Refs #N in intermediate commits; use a closing keyword only in the final evidence-complete commit.

## Detailed implementation plan

### Phase 0 — Prepare evidence mechanisms and restore Excel execution

Objective: finish all work that does not require a fresh Excel observation, while keeping binding claims and active row dispositions truthful.

#### Track A — Excel-free release mechanisms

This track gates Phase 1 and can finish even if the runner remains unavailable.

**Status: complete at 0dd7488.**

| Commit | Completed work |
| --- | --- |
| `2dcfe04` | #29 fail-closed holdout-provenance writer, verifier, fixtures, documentation, and staged gate wiring |
| `136b66d` | #22 canonical row-disposition checker and frozen 36-row transition guard |
| `4be8cd0` | recursive holdout-exporter binding and added-module fixtures |

All three commits preserved the exact Phase 0 seven-line/two-file expected-red signature. No current-source holdout manifest was falsely created.

1. Keep bde92dd as the numerical source baseline for #23 validation; the later issue-template commits do not change numerical behavior.
2. Implement the Excel-free half of #29:
   - manifest writer and verifier;
   - fail-closed CI and `refresh_evidence.py` wiring;
   - `PROVENANCE.md`;
   - clean, changed-source, changed-grid, changed-registry, changed-schema, missing-manifest, and line-ending tests.
3. Do not bind the existing 559 holdout rows to current source. They were exported at 4553afa, so the verifier must report them stale or unbound until #23 produces a fresh export. #29 remains open.
4. Implement #22's cheap coverage checker before further value-changing commits:
   - strict mode requires every main-grid row to have a real contract or final explicit exemption and fails on any missing row;
   - transition mode compares canonical binary64 row identities with a frozen fingerprint of the 36-row bde92dd audit-baseline debt set;
   - transition mode permits only a subset of that debt, reports it as `KNOWN COVERAGE DEBT`, and fails on any new or changed unclaimed row;
   - the fingerprint records owner, issue, milestone expiry, source/grid identity, and its initial audit-baseline count, while the current remaining-debt count is always derived;
   - an intentional grid/regime edit may not land alone: it must be atomic with either an explicitly reviewed fingerprint transition or the governing contracts plus fingerprint deletion;
   - fixtures cover the exact baseline, one new row, changed arguments, changed regime, a resolved row, and a corrupted fingerprint.
5. Wire transition mode into development CI and the generated summary without describing completeness as PASS. Strict mode remains the release authority. The fingerprint is neither a contract nor a final exemption and must be deleted when #22 closes.
6. Confirm the F runtime authority explicitly:
   - `PROB_F_MAX_DF = 1E10`;
   - `K_STATS_F_InverseCumulative` passes that constant to `PROB_F_ValidateEnvelope`;
   - df = 1E6 is therefore accepted production input, and the source comment records `inverse_probe` success through df = 2E10 without refusals, including previously suspect unbalanced ratios;
   - the stale text is in all four F `validated` contract domains—CDF, survival, inverse quantile, and inverse tail—which wrongly says the <=1E5 core regime is enforced by the runtime validator.
7. Do not create either current-source provenance record until its observations have actually been exported from real Excel.
8. Confirm the only intended worktree additions are task-specific; never add `.deps`, local workbooks, or upload scratch.

#### Track A2 — #22 oracle and contract preparation in parallel

This work begins in Phase 0 but is not a Phase 1 entry gate. It may continue alongside #23, #24, #13, #14, and #26.

1. Preserve the final policy: all 36 baseline unclaimed rows receive real contracts; none is finally exempted or retired. Five core Student-t medians join the three large-df medians in an absolute regime, and three older claimed F rows are removed only because higher-precision duplicates already exist.
2. Preregister fitting and untouched holdout designs for F, Student-t, and Chi-square inverse surfaces. Holdout points stay hidden until the metric and thresholds are frozen.
3. Freeze the Student-t decision inputs:
   - fitting offsets `2^-53, 2^-43, 2^-33, 2^-23` on both sides of 0.5 at df = 1E6, 1E7, 1E8;
   - untouched holdout offsets `2^-48, 2^-38, 2^-28, 2^-24`;
   - bridge fitting probabilities 0.625, 0.75, 0.875 and complements;
   - disjoint bridge holdout probabilities 0.5625, 0.6875, 0.8125, 0.9375 and complements;
   - the seven-contract relative route only when every nonzero fitting quantile is nonzero, finite, and no worse than 5E-9;
   - otherwise the eight-contract absolute fallback with `Abs(q_reference) >= 100 * T_abs` applied before holdout inspection.
4. Retain the source-path prediction: exact input subtraction near 0.5 should make nonzero neighbours suitable for relative scoring, with any observed floor attributable to the seed/forward-rounding implementation rather than intrinsic conditioning.
5. Build the reference authority:
   - factor the converging lower-series / upper-Lentz-CF incomplete-gamma route for Chi-square and revalidate at materially higher precision;
   - use the robust incomplete-beta route and exact public transforms for Student-t and F;
   - cross-check fitting and holdout references independently with Rmpfr wherever its incomplete-gamma routine can execute;
   - at df = 1E8, use the accepted converging gmpy2/MPFR route plus algorithm-independent quadrature;
   - fail closed on non-convergence or failure to stabilize.
6. Satisfy a hard early feasibility checkpoint before Phase 2 begins, without making it a #23 closure prerequisite:
   - generate the frozen Chi-square reference set at df = 1E6, 1E7, and 1E8 with the converging series/continued-fraction route;
   - reproduce it at two materially separated working precisions and reject any unstabilized value;
   - cross-check all 69 points with algorithm-independent quadrature and gmpy2/MPFR arithmetic, plus Rmpfr on every feasible point;
   - record the agreement, compared measures, precision pair, convergence status, runtime, and the Rmpfr ceiling.

**Checkpoint complete and maintainer decision recorded.** Commits `4deeef4` and `e68cc82` freeze the 69-point set and correct the stored-reference truncation. Quadrature and gmpy2/MPFR converge 69/69 with minimum agreement of 109.87 and 115.11 significant digits respectively. Rmpfr agrees to at least 115.96 digits on all 46 points it can execute, but its `mpfr_gamma_inc` aborts above shape approximately 4.5E7; df = 1E8 requires shape 5E7. The df = 1E8 substitution is accepted: gmpy2/MPFR supplies the converging arithmetic route and quadrature supplies the algorithm-independent cross-check. Rmpfr remains authoritative where feasible. This decision freezes only the oracle feasibility, not a production accuracy threshold, and no #22 holdout has been inspected.
7. Add the machine-readable inverse-regime classifier and seam fixtures. Its disjoint priorities are: existing F tiny-unbalanced predicate; F core when both df <=1E5 and envelope otherwise through 1E10; Student-t exact median first, core below df 1E6 and selected large-df regime through 1E8; Chi-square core below df 1E6 and envelope through 1E8.
8. Preregister the F cleanup:
   - keep the higher-precision `inverse_probe` rows at p = 0.5, 0.9, 0.99 with df = (1E6, 3);
   - remove the three older duplicate `validated` rows;
   - reclassify the five remaining unique main-grid rows above 1E5 and the eight already-inspected holdout rows at `(df1, df2) = (1.5, 2E5)` and `(5E5, 4)`, four probabilities each;
   - treat those holdout rows as legacy support, not the new independent holdout;
   - remove the false runtime-enforcement wording from all four F `validated` domain strings;
   - retain the established 2E-10 F core threshold conservatively rather than retightening it opportunistically.
9. Extend the main and holdout evaluators for Chi-square/Student-t tail residuals, compensated `hi + lo` observations, exact-zero handling, wrong regimes, bridge/seam routing, clean/degraded cases, and zero-return mutants.
10. Add explicit exact-invariant threshold-zero derivation. Test parse, exact pass, nonzero fail, main/holdout verdicts, rejection of an empirical-zero shortcut, summary/README formatting, and disposition reconciliation.
11. Record the current diagnostic baseline without freezing it: F quantile 9.3115E-9; Student-t nonzero large-df quantile 1.8480E-10; Chi-square quantile 1.0890E-15 under `hi + lo`; preliminary Chi-square tail residual about 1.1207E-11 pending independent cross-check.
12. Do not observe the #22 fitting or holdout arms yet. Final measurement occurs after the last inverse-affecting numerical source change, no earlier than #14.

#### Track B — Excel execution route

This track gates entry into Phase 1 but does not block Track A.

1. Preferred: restore the existing self-hosted runner and execute run 95 or a fresh run on the same source.
2. Contingency: provision a replacement Windows/Excel 64-bit self-hosted runner using the same exact-source import and artifact procedure.
3. Interim: a controlled manual fresh-workbook Excel run may clear an individual numerical phase only if it records source hashes, Excel version/build/bitness, assertion totals, exported observations, and provenance in the same machine-readable format.
4. The interim route does not waive the final requirement for a retained self-hosted workflow artifact on the exact release commit.
5. If neither controlled Excel route is available, stop after Track A preregistration and do not accumulate further value-changing numerical commits.

Exit gate: #29's writer/verifier and #22's strict checker plus the transition guard are committed and green in their intended modes; the current debt count is generated rather than hard-coded; the repository-level Accuracy Gate remains intentionally red with exactly the frozen seven-line/two-file signature and no additional failure; the holdout remains honestly stale; the #22 Chi-square oracle checkpoint is accepted; and a controlled Excel route is available before Phase 1 begins.

### Phase 1 — Validate and close #23

Objective: prove the merged off-grid `PROB_StirlingError` recurrence removes the public Gamma/Chi-square CDF floor without waiting for #22's new inverse reference infrastructure.

1. Before exporting anything, reproduce the frozen Phase 0 signature exactly: five PASS lines, two stale-evidence FAIL lines, exactly two mismatches, and only the two frozen paths. In the same run, verify #29's Excel-free mechanism and #22's transition guard.
2. Run the expected 909 assertions on the exact tracked modules.
3. Export the preregistered #23 fitting grid and untouched #23 holdout only; do not execute or inspect #22's inverse fitting/holdout arms.
4. Export the current global holdout, write and verify the first truthful #29 binding immediately, and close #29 only after deliberate source/grid mismatches fail.
5. Compare #23 movement with `benchmark/gamma_series_small_shape/MOVEMENT_MANIFEST.md`.
6. Rerun the frozen `positive_ratio_subnormal` fitting arm and untouched holdout:
   - compare direct-path and log-recovery errors around the 48-bit and 40-bit candidate boundaries;
   - retain or rederive each exact power-of-two cutoff;
   - amend preregistration before any #13 production call site if a crossover moves.
7. Remeasure the X = 2 continued-fraction ladder for #26.
8. Promote representative #23 small-shape and ordinary off-grid rows under existing public contracts.
9. Re-export the complete current main grid and global holdout, write both truthful provenance bindings atomically, and generate main/holdout summaries.
10. Run every numerical strict gate and unrelated holdout. The expected-red waiver must retire with all seven checks PASS. Run #22 in transition mode and prove #23 introduced no new or changed unclaimed row.
11. Close #23 only if its unchanged frozen contracts pass and coverage debt did not grow.

Parallel feasibility gate before Phase 2: complete. The frozen df = 1E6/1E7/1E8 references stabilize at two precisions; quadrature and gmpy2/MPFR cover 69/69, while Rmpfr agrees on its 46 feasible points. The accepted df = 1E8 substitution and Rmpfr ceiling are recorded in #22.

Exit gate: #29 and #23 are closed; the frozen known-red signature has been replaced by seven PASS results; the 3E-15 Gamma cumulative contract is preserved; main-grid and holdout evidence are bound to the same current source; #13 cutoffs and #26 have post-#23 decision inputs; and #22 remains open on its release-candidate track with no new coverage debt.

### Phase 2A — Wire #13 density and cumulative arms

Objective: use the post-#23 validated cutoffs without waiting for lower-region survival reconstruction.

This phase may proceed in parallel with Phase 2B after Phase 1 establishes the final cutoffs. It is an implementation option, not a requirement to split #13 closure.

1. Reuse `PROB_MIN_NORMAL`, the enum, and classifier, but only with the post-#23 validated or rederived cutoffs.
2. Add shared log-domain cumulative and density arithmetic in SPECIALFUNCS.
3. Dispatch at the Gamma and Chi-square density/cumulative call sites while the original operands remain in scope.
4. Keep `X / ScaleParam` and `0.5 * X` policies separate.
5. Add exact seam-neighbor tests and separate family contracts.
6. Run fitting, the now source-bound untouched holdout, full Excel, provenance, and gates.

Exit gate: four density/cumulative surfaces satisfy their post-#23 regimes. #13 remains open pending survival integration.

### Phase 2B — Implement #24 direct lower-region Q

Objective: remove complement amplification without masking an inaccurate P.

1. Add a study-only lower-series LogP output and decompose prefactor plus series-sum error.
2. Preregister shape ladder, seam points around X = A + 1, movement, and untouched holdout.
3. Refactor shared series arithmetic so P uses Exp(LogP) and Q uses -PROB_Expm1(LogP).
4. Never recover LogP by taking Log of an already rounded P.
5. Add direct-kernel and equivalent Gamma/Chi-square VBA regressions.
6. Add/freeze machine-readable lower-region survival evidence.
7. Validate monotonicity, P/Q consistency, error codes, Poisson users, and the CF seam.
8. Run full Excel/export/provenance/gates.

Exit gate: lower Q never forms 1 - P; 5E-15 Gamma survival contract preserved; #24 closed.

### Phase 3 — Complete #13 survival and integrated closure

Objective: preserve mathematically positive standardized arguments through the six public forward surfaces.

1. Confirm Phase 1's post-#23 cutoff record; do not assume the pre-#23 `/ 16#` and `/ 4096#` constants survived.
2. Integrate Phase 2A density/cumulative work, or implement those four surfaces now if atomic delivery was preferred.
3. After #24, add shared survival hard-underflow reconstruction from the direct lower-region log result.
4. Dispatch at all six public-family call sites while the original operands are in scope.
5. Use no retained-bit survival dispatch unless new measurement explicitly proves one; use mandatory hard-underflow recovery on all six surfaces.
6. Probe exact seam neighbors produced by arithmetic.
7. Add/freeze separate Gamma and Chi-square regimes; do not pool policies.
8. Compare movement against the post-#23 preregistration and rerun the fitting arm and source-bound untouched holdout.
9. Run full Excel/export/provenance/gates.

Exit gate: all six parent forward counterexamples correct; no body-regime regression; #13 closed.

### Phase 4 — Implement #14 log-domain inverse reconstruction

Objective: materialize only the final-scale quantile.

1. Study and preregister the crossover between ordinary and log-quantile inversion.
2. Add PROB_TryGammaInvPLog or the smallest equivalent shared kernel.
3. Use PROB_TryLogGamma1p(Shape) in the small-shape asymptotic.
4. Select the log branch before any seed that can overflow or underflow.
5. Reconstruct Gamma with Log(ScaleParam) and Chi-square with Log(2#).
6. Exponentiate once under the existing final overflow/underflow contract.
7. Add quantile-relative and forward tail-residual contracts.
8. Add exact-binary64 seam, representable-final, true-final-underflow, and true-final-overflow VBA tests.
9. Run fitting, untouched holdout, full Excel/export/provenance/gates.

Exit gate: corrected forward/inverse round-trips satisfy frozen contracts; #14 and then #11 close.

### Phase 5 — Resolve #26 from evidence

Objective: determine whether a CF-specific defect remains after #23.

- If the unchanged X = 2 ladder passes, attribute the old breach to the shared prefactor, add permanent coverage, promote rows, and close with no CF source change.
- If it fails, decompose prefactor, h recurrence, final multiply, tolerance, and iteration budget in a study harness.
- Preregister any residual fix and untouched holdout before changing the recurrence.
- Validate seam, moderate/large shape, Poisson, monotonicity, and non-convergence behavior.

Exit gate: the continued-fraction route satisfies the unchanged 5E-15 contract and #26 closes.

### Phase 6 — Close #22 on the release-candidate numerical source

Objective: finish the heavy inverse-contract work after the numerical chain, so its longest-pole reference generator does not block #23/#13/#14 and its measurements do not need to be repeated after #14 changes inverse architecture.

1. Require #23, #24, #13, #14, #26, and parent #11 to have their final numerical source disposition. #22's reference/oracle preparation and Rmpfr cross-check must already be complete.
2. Verify the transitional guard still reports only an authorized subset of the recorded audit-baseline debt set and derives its current count; any unauthorized new or changed missing row blocks the wave.
3. Export the preregistered #22 fitting arms from the release-candidate source.
4. Apply the frozen Student-t relative-versus-absolute decision and derive thresholds without inspecting holdout observations.
5. Export and inspect the untouched #22 holdouts only after the route and thresholds are frozen.
6. Apply the complete F correction:
   - remove three older duplicate main rows;
   - reclassify five remaining main rows and the eight legacy holdout rows at `(df1, df2) = (1.5, 2E5)` and `(5E5, 4)`, four probabilities each;
   - remove the false `PROB_F_ValidateEnvelope` enforcement phrase from all four F `validated` domains;
   - preserve the runtime 1E10 cap as the distinct production envelope.
7. Atomically land the selected seven or eight new contracts, all eight Student-t median relabels, F/Student-t/Chi-square domain partitions, promotions, machine-readable classifier, evaluator fixtures, strict checker, and regenerated summaries.
8. The same final #22 closure commit owns the F row reclassification and duplicate removal. It deletes the transitional fingerprint rather than refreshing it to a larger unclaimed set, and strict mode must report zero missing dispositions. No intermediate direct-main commit may change those row identities without atomically updating the reviewed fingerprint and generated current-debt count.
9. Run the full real-Excel regression/export/provenance/holdout sequence. Where practical, use this as #31's release-candidate evidence wave rather than scheduling an additional standalone cycle.

Exit gate: #22 closes with zero missing dispositions, no domain gap/overlap, no transitional waiver, and source-bound fitting plus independent holdout evidence on the final numerical source.

### Phase 7 — Complete assurance infrastructure

#29 is already source-bound and #22 is now strict. Refresh their generated evidence after the final numerical source change and before publishing public assurance metrics.

#### #17 non-destructive grid-regeneration safeguards

- make `generate_reference_values.py` report-only or non-authoritative by default;
- require an explicit reviewed flag before any authoritative combined-grid write;
- preserve `observed_vba` outside the Excel exporter;
- fail hard on duplicate canonical keys and unexpected reference movement;
- require the existing explicit reasoned action for row retirement;
- add focused CI fixtures for each safeguard.

Complete origin reconstruction, reference/observation redesign, and byte-stable clean regeneration are deferred to #32 under v1.01 and are not v1.0.0 release criteria.

#### #28 README assurance generation

- Add a machine-readable real-Excel result record.
- Render badges and table from one offline model.
- Make stale main-grid or holdout provenance visible and blocking.
- Remove unsupported zero-known-defect wording unless backed by a committed readiness registry.
- Fail CI on README drift.

#### #30 GitHub Actions runtime

- Review breaking changes and upgrade checkout, setup-python, upload-artifact, and github-script consistently.
- Preserve permissions, runner labels, filters, artifact behavior, and strict failure semantics.
- Verify all workflows and remove Node 20 warnings.

Exit gate: the current generator is non-destructive and fail-closed; assurance is complete, current, reproducible, and fail-closed. #32 remains correctly outside v1.0.0.

### Phase 8 — Certify and publish under #31

1. Correct SECURITY.md pre-release language.
2. Add CHANGELOG.md with Unreleased and v1.0.0 entries.
3. Audit README, installation, wiki links, example workbook, templates, license, and limitation documentation.
4. Perform the final clean Excel import/compile/regression/export on the intended release commit.
5. Bind main grid and holdouts to that source; regenerate every summary and README metric.
6. Require zero FAIL, zero PENDING, zero missing main-grid dispositions, no transitional coverage-debt fingerprint, and no unexplained movement.
7. Rerun hosted and self-hosted workflows on the exact release commit.
8. Freeze release notes and artifact policy.
9. Create annotated tag v1.0.0 and the GitHub Release.
10. Verify tag, assets, hashes, public links, README, and SECURITY.md from GitHub.

Exit gate: #31 closes only after publication verification.

## Commit and evidence strategy

Suggested atomic commit sequence:

1. #29 Excel-free writer/verifier/CI/documentation/tests, with the holdout correctly reported stale;
2. #22 strict checker, fixtures, and temporary canonical fingerprint of the 36-row audit-baseline debt set; development CI derives the current debt count, fails on any unauthorized new debt, and never reports completeness PASS;
3. #23 real-Excel fitting, untouched holdout, current global-holdout export, and first truthful #29 binding;
4. #29 evidence-complete closure;
5. #23 closure evidence, including post-#23 #13 cutoff revalidation, #26 ladder, main-grid export, provenance, and the proof that #22 debt did not grow;
6. amended #13 cutoff preregistration if the crossover moved, or an evidence-only retention record if it did not;
7. optional #13 density/cumulative source-and-test commit;
8. #13 density/cumulative export/provenance evidence if step 7 was used;
9. preregistration commit for #24;
10. #24 source/test commit;
11. #24 export/provenance/contract closure commit;
12. #13 survival integration and six-surface test commit;
13. #13 export/provenance/contract closure commit;
14. preregistration commit for #14;
15. #14 source/test commit;
16. #14 export/provenance/contract closure commit and #11 parent closure;
17. #26 evidence-only closure or preregistration/fix/closure sequence;
18. parallel #22 classifier/zero-threshold preparation commits as they become ready, without observing the holdout; the accepted Chi-square df = 1E6/1E7/1E8 oracle checkpoint is already recorded at e68cc82;
19. #22 release-candidate fitting export, frozen Student-t route selection, then untouched holdout;
20. atomic #22 selected-contract/domain-string/row-cleanup/regime/classifier/strict-checker/summary closure commit; this commit owns the F row transition and deletes, rather than expands, the temporary debt fingerprint;
21. #17 non-destructive/fail-closed generator safeguards and focused CI fixtures;
22. #28 generated README metrics;
23. #30 workflow major upgrades;
24. #31 documentation and final release-evidence commit;
25. annotated v1.0.0 tag and GitHub Release.

Each .bas/grid/contract commit must include:

~~~text
Evidence: <exports/regenerated artifacts, or none>
Gate: FAIL=<n> PENDING=<n>
Tests: <pass>/<total>
~~~

Mandatory generation order after a real Excel export:

1. export observations from exact tracked source;
2. write main-grid and holdout provenance;
3. generate main and holdout summaries;
4. run row disposition in the phase-appropriate mode: transition mode may recognize only the frozen baseline debt and must fail on any new row; strict mode is mandatory from #22 closure onward;
5. generate root README assurance metrics;
6. run strict verification and diff checks;
7. commit source, observations, provenance, summaries, and generated documentation together.

## Evidence matrix

| Change | Minimal counterexample | Fitting | Independent holdout | VBA regression | Main-grid contract | Full Excel | Unrelated checks |
| --- | --- | --- | --- | --- | --- | --- | --- |
| #23 | Gamma/Chi CDF at A=.0001, X=.5 | off-grid shapes | frozen #23 design | 7 new assertions | promote representative rows | required | CF-Q, Poisson, large shape |
| #24 | Gamma/Chi survival at A=.0001, X=.5 | shape ladder + seam | untouched shape holdout | kernel + two public routes | lower-Q regime | required | P/Q, CF, Poisson |
| #13 | six parent forward calls | post-#23 retained-bit crossover rerun | source-bound frozen Gamma/Chi holdouts | six surfaces + seams | separate family regimes | required | body Loader/deviance and cutoff movement |
| #14 | scaled representable quantile | crossover + scale matrix | untouched log-inverse holdout | round-trip + true limits | quantile + tail residual | required | ordinary inverse regimes |
| #26 | Gamma/Chi Q at A=.0001, X=2 | CF ladder + seam | untouched CF holdout | two public routes | CF representative rows | required | series, Poisson, non-convergence |
| #22 | 36-row audit-baseline inverse debt set; four false F runtime-enforcement domain strings; eight mis-scoped F main rows; eight mis-scoped F holdout rows at df 2E5/5E5; three duplicate F inputs; eight exact-zero Student-t medians | release-candidate inverse fitting; Student-t `delta = 2^-53, 2^-43, 2^-33, 2^-23`; dyadic bridge p = .625/.75/.875 plus complements; accepted 69-point Chi-square df = 1E6/1E7/1E8 oracle set | disjoint offsets `2^-48, 2^-38, 2^-28, 2^-24`; bridge p = .5625/.6875/.8125/.9375 plus complements; Chi-square quadrature and gmpy2/MPFR 69/69, Rmpfr 46/69 where feasible; robust-beta references checked independently | inverse round-trip, exact-zero assertions, F/t/chi seams, p=.6 routing, zero-threshold and zero-return mutants | transition guard before numerical commits; final audit-baseline dispositions/domain cleanup and duplicate consolidation after #14 | release-candidate wave required | new-debt fingerprint, lost-contract/stale-disposition, duplicate/conflicting regime, zero-reference, cutoff, formatting, and unsupported-oracle fixtures |
| #29 | changed-source fixture | n/a | first truthful binding written on the Phase 1 global holdout export | n/a | n/a | export required for closure | manifest unit tests |
| #17 | destructive-default and duplicate/reference-change fixtures | n/a | n/a | observation-preservation checks | canonical-key safeguards | no Excel export required | safe-default, explicit-write, retirement, and failure-path fixtures |
| #28 | manual/stale metric fixture | n/a | reads #29 state | reads Excel record | reads #22 state | result record | README diff |
| #30 | workflow run | n/a | analyzer still runs | artifact upload | strict gate still runs | runner required | label sync |
| #31 | clean import | all resolved | all resolved | all pass | all current | release commit | public release verification |

## Open backlog

| Order | Issue | Priority | Dependency | Immediate next action | Closure artifact |
| ---: | --- | --- | --- | --- | --- |
| 1 | #29 | P2 | Excel-free mechanism complete; Excel-bound closure | write the first truthful main/holdout binding pair atomically during #23 export | source-bound current holdout and seven PASS checks |
| 2 | #22 | P2 | coverage guard and Chi-square oracle checkpoint complete; heavy fitting/holdout, eight-row zero-reference split, neighbour/bridge decision arm, deterministic regimes, and F cleanup remain | continue hidden-holdout classifier/contract preparation; run fitting/holdout and atomically close on the release-candidate source after #14 | final zero missing rows, no domain gap/overlap, no temporary fingerprint |
| 3 | #23 | P1 | #29 mechanism, #22 delta guard, controlled Excel route | first reproduce the exact Phase 0 expected-red signature, then run its evidence wave; close #29 then #23 | fresh source-bound numerical grid/holdout; seven PASS checks; coverage debt unchanged |
| 4 | #24 | P1 | #23 | preregister LogP/direct-Q study; #13 density/CDF may proceed in parallel | direct-Q contracts and tests |
| 5 | #13 | P1 | #23 cutoff gate; #24 only for survival | retain/rederive cutoffs, then wire per-family dispatch | six surface regimes |
| 6 | #14 | P1 | #13 | preregister log-inverse crossover | quantile + tail residual |
| 7 | #26 | P1 | #23 | act on the re-exported X=2 ladder | evidence-only close or isolated fix |
| 8 | #11 | P1 | #13, #14 | close parent after round-trip evidence | parent counterexamples |
| 9 | #17 | P2 | independent of numerical source | make default regeneration non-destructive and add fail-hard fixtures | safe generator path; #32 deferred |
| 10 | #28 | P2 | #22, #29 | implement one assurance renderer | generated root README |
| 11 | #30 | P3 | independent | review and upgrade action majors | warning-free workflows |
| 12 | #31 | P1 | all v1.0.0 blockers | maintain release checklist | tag and GitHub Release |

## Go / no-go checklist

GO requires every item:

- [ ] all milestone P1 issues closed with source-bound evidence;
- [ ] no accepted-domain silent-wrong path deferred through documentation-only limitation wording;
- [ ] full real-Excel regression passes on the release commit;
- [ ] a retained self-hosted Excel workflow artifact exists for the exact release commit;
- [ ] main grid and independent holdouts were exported from the same release source;
- [ ] zero FAIL and zero PENDING;
- [ ] zero missing main-grid dispositions;
- [ ] the temporary #22 coverage-debt fingerprint has been deleted and strict mode is authoritative;
- [ ] no frozen contract relaxed;
- [ ] README metrics generated and current;
- [ ] SECURITY.md no longer claims a nonexistent stable tag;
- [ ] CHANGELOG.md and release notes complete;
- [ ] workflows green without deprecated runtime warnings;
- [ ] clean import/example smoke test passes;
- [ ] annotated v1.0.0 tag and GitHub Release verified.

Any unchecked item means NO-GO.

## Complete v1.0.0 milestone issue register

This register contains every issue assigned to milestone v1.0.0 after reconciliation, including closed history. Titles, states, labels, and bodies are reproduced from the live GitHub issues. Comments are not included.

### #2 — CR-P1-02: incomplete gamma/beta prefactor cancellation at large shape

- State: closed (completed)
- Labels: bug, P1, special-functions, continuous, accuracy
- URL: https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/issues/2

#### Body

## Summary

The regularized incomplete gamma and beta kernels formed their prefactor by
literal subtraction:

    -X + A*Log(X) - LogGamma(A)             (incomplete gamma)
    A*Log(X) + B*Log(Y) - LogBeta(A, B)     (incomplete beta)

Each subtracts two quantities of size `A*Log(A)` to leave a modest logarithm.
`PROB_LogGamma` carries a *relative* error contract, so the absolute error of
the difference grows with shape. Every cumulative, survival and inverse
probability outside the normal family inherits that error, silently.

Class: silent wrong answer inside the accepted domain.

## Evidence (absolute error in the log-prefactor, worst case X = A / mode)

| shape | incomplete gamma | incomplete beta |
| --- | --- | --- |
| 1E6  | 6.8E-10 | 1.7E-11 |
| 1E10 | 7.2E-06 | 2.9E-07 |
| 1E12 | 2.0E-03 | 5.0E-05 |
| 1E16 | 4.7E+01 | 8.5E-01 |
| 1E20 | 5.2E+05 | 2.2E+01 |

Reachable: `Gamma_Cumulative`/`Survival`/`InverseCumulative` and the Beta
equivalents have **no envelope** (shapes accepted to the 1E100 magnitude
guard), while frozen contracts reached only shape 3 for Gamma and 1E5 for Beta.
T/Chi/F were shielded by their existing df envelopes.

## Root cause

Same defect class as CR-P1-01, in the *incomplete* functions rather than the
densities. The stable composite kernels built for CR-P1-01 already express both
prefactors exactly:

    Log(X^A e^-X / Gamma(A))       = GammaLogPdf(X; A, scale 1) + Log(X)
    Log(X^A Y^B / Beta(A, B))      = BetaLogPdf(X, Y, A, B) + Log(X) + Log(Y)

## Fix

- `PROB_TryGammaPrefactor` added; used by `PROB_TryGammaSeriesP` and
  `PROB_TryGammaContinuedFractionQ`
- `PROB_TryGammaInvP` Newton derivative routed through `PROB_TryGammaLogPdf`
- `PROB_TryBetaRegularized` routed through `PROB_TryBetaLogPdf`, behind a
  **measured** regime dispatch `PROB_IBETA_LOADER_MIN_SHAPE = 1000`: below the
  crossover the literal form is at least as accurate and is the form the frozen
  tiny/unbalanced contracts were validated against

Measured after the fix: series path improves ~1000-5000x at large shape; the
binding limit becomes the series stopping criterion, not cancellation.

## Status

- [x] Root cause identified and measured
- [x] Prefactors rerouted through the stable kernels
- [x] Regime dispatch added so tiny/unbalanced contracts are not regressed
- [x] Permanent VBA regressions (A = 1E4, 1E6) with measured tolerances
- [x] 148 observations re-exported; all 147 contracts green
- [x] `cdf_large_shape` study exported and analysed
- [x ] Contracts frozen for the repaired regime, holdout validated

## Closure criteria

1. Accuracy at large shape measured in Excel, not simulated
2. Contracts frozen from main-grid measurements, validated on an independent
   holdout subset
3. The validated-domain edge documented (the lower-tail series reaches
   `PROB_GAMMA_MAX_ITER` above ~1E8 and returns `#NUM!`, which is correct
   Try-contract behaviour, not a defect)
4. Evidence regenerated in order: export -> `write_manifest.py` ->
   `compute_errors.py --out accuracy_summary.md`

### #3 — PROB_NUM_EPS: series/CF stopping criterion limits large-shape accuracy

- State: closed (completed)
- Labels: enhancement, performance, P3, core, special-functions, accuracy
- URL: https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/issues/3

#### Body

## Summary

`PROB_NUM_EPS = 3E-14` (CORE) is the stopping criterion for both the ascending
series and the Lentz continued fractions. After CR-P1-02 removed the prefactor
cancellation, this constant — not cancellation — became the binding limit on
large-shape cumulative accuracy.

Class: accuracy opportunity, not a defect. Nothing is silently wrong.

## Evidence

Regularized lower incomplete gamma via the series, at X = A - sqrt(A):

| A | stop at 3E-14 (current) | stop at machine eps | terms (current -> eps) |
| --- | --- | --- | --- |
| 1E4 | 3.89E-13 | 2.10E-15 | 646 -> 711 |
| 1E6 | 4.14E-12 | 2.85E-14 | 6,097 -> 6,760 |

Roughly **145x more accuracy for about 10% more iterations**. The A = 1E6 figure
was confirmed against Excel: the measured 4.14E-12 reproduces exactly when the
3E-14 criterion is mirrored, which is what identified the cause.

## Why this is not a trivial change

1. **Blast radius.** `PROB_NUM_EPS` is shared by the incomplete gamma series,
   the incomplete gamma CF and the incomplete beta CF. Changing it moves
   observations for Gamma, Beta, ChiSquare, StudentT, F, Binomial, Poisson and
   NegativeBinomial simultaneously.
2. **Interaction with the iteration cap.** The lower-tail series needs
   O(sqrt(A)) terms and already reaches ~57,600 at A = 1E8 against
   `PROB_GAMMA_MAX_ITER = 100000`. A 10% increase moves the `#NUM!` boundary
   slightly *down*. That trade needs measuring, not assuming.
3. Frozen contract thresholds were derived under the current criterion. They
   should improve, but every one must be re-measured.

## Options

| | approach | trade |
| --- | --- | --- |
| A | Tighten `PROB_NUM_EPS` globally | Simplest; largest blast radius |
| B | Dedicated tighter epsilon for the incomplete gamma/beta kernels only | Contained; adds a second constant |
| C | Leave as is, document the measured limit | Zero risk; leaves ~145x on the table |

## Closure criteria

1. Decision recorded with reasoning
2. If changed: full re-export, every contract re-measured, iteration-cap
   boundary re-measured, `numerical_limitations.csv` updated
3. If not changed: the measured limit documented so it is not rediscovered

### #4 — Revisit StudentT/ChiSquare/F df envelopes after CR-P1-02

- State: closed (completed)
- Labels: enhancement, P3, t-family, accuracy
- URL: https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/issues/4

#### Body

## Summary

Current caps:

| function family | constant | cap |
| --- | --- | --- |
| StudentT CDF/survival/inverse | `PROB_T_MAX_DF` | 1E6 |
| ChiSquare CDF/survival/inverse | `PROB_CHI_MAX_DF` | 1E6 |
| F CDF/survival/inverse | `PROB_F_MAX_DF` | 1E5 |

These were set in P2-02 and earlier, when the incomplete gamma and beta kernels
still carried the CR-P1-02 prefactor cancellation. They may now be conservative.

## Hypothesis worth testing

The P2-02 measurement found ChiSquare survival wrong by ~7E2 at df = 1E16.
ChiSquare routes through the incomplete gamma with A = df/2 = 5E15, where the
old prefactor was wrong by roughly e^46. **The P2-02 divergence and CR-P1-02 are
plausibly the same defect.** If so, the chi cap was measuring the prefactor bug
rather than an intrinsic limit.

This is a hypothesis, not a finding. It has not been measured post-fix.

## Constraints

- Raising a cap is only defensible on fresh measurement. The house policy is
  strict validated domains, and a cap set by assumption is worse than a
  conservative one.
- The lower-tail series iteration cap (see #2) may bind before accuracy does,
  in which case `#NUM!` is the correct answer and the cap should stay.

## Closure criteria

1. A study measuring each family across its candidate range, post-CR-P1-02
2. Caps either raised to a measured boundary or explicitly reaffirmed
3. `numerical_limitations.csv` and the envelope validator headers updated to
   cite the new measurement

Related: #4 (the chi band that blocks this for ChiSquare)

### #5 — ChiSquare df 1E6-1E16 band is unmeasured (registered limitation)

- State: closed (completed)
- Labels: enhancement, P3, t-family, accuracy
- URL: https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/issues/5

#### Body

## Summary

`PROB_CHI_MAX_DF = 1E6` is **not** a measured degradation boundary. It is the
last df validated by an existing test (regression T1 validates
`ChiSquare_Cumulative(1E6, 1E6)`). The band from 1E6 to 1E16 was never measured
because it was reference-limited at the time; the kernel is known to diverge by
df = 1E16 (survival off by ~7E2).

Registered in `numerical_limitations.csv` as
`IncompleteGamma.ChiSquareLargeDF`, whose note already says
"pending a dedicated 1E6-1E16 study".

## Why it may now be tractable

The `cdf_large_shape` study (CR-P1-02, #1) solved the reference problem that
originally blocked this: mpmath's own `gammainc` fails above A ~ 1E8, but
computing by the converging route (series below A+1, Lentz CF above) and
revalidating at higher precision produces self-checked references to at least
A = 1E10. The same generator design should extend to this band.

## Closure criteria

1. References produced and self-checked across 1E6 to 1E16
2. The true degradation boundary located
3. `PROB_CHI_MAX_DF` set to it, or reaffirmed at 1E6 with the measurement cited
4. `numerical_limitations.csv` note updated from "pending" to the result

Blocks: #3 (for ChiSquare)

### #6 — StudentT_Density at large df is unenveloped and unstudied

- State: closed (completed)
- Labels: enhancement, P3, t-family, accuracy
- URL: https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/issues/6

#### Body

## Summary

`K_STATS_StudentT_Density` has **no** accuracy envelope and **no** large-df
study. CR-P1-01B added `PROB_DENSITY_SHAPE_MAX = 1E20` to the Gamma, Beta,
ChiSquare and F densities, but StudentT_Density was not included.

Its only large-df coverage is:

    AssertTrue "t density not enveloped (large df)", _
        (Not IsError(K_STATS_StudentT_Density(0#, 1E+18)))

which asserts that it returns *a* value — not that the value is correct.

Class: unmeasured domain. No known wrong answer; also no evidence of a right one.

## Context

The t density is a Beta-type expression in disguise and shares the LogBeta and
LogGamma machinery that CR-P1-01/CR-P1-02 stabilised elsewhere, so it is a
plausible carrier of the same cancellation. As df grows the t density converges
to the standard normal density, which gives an independent oracle for the far
end of the range.

## Closure criteria

1. Measured against mpmath across df to at least 1E20, at and around the mode
2. Either contract-frozen at the measured accuracy, or enveloped at the measured
   boundary consistently with the other four densities
3. The placeholder `Not IsError` assertion replaced with a value assertion

### #8 — Refresh holdout

- State: closed (completed)
- Labels: bug, testing, P2, accuracy
- URL: https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/issues/8

#### Body

## Summary

The contract registry was tightened after the gamma-series epsilon change
(#3), but the consolidated holdout observations were not re-exported. Running
`analyze_holdout.py` against the current registry today fails two contracts:

| contract | threshold | committed holdout worst | over |
| --- | --- | --- | --- |
| `Poisson_Cumulative.all.output_rel` | 2E-13 | 2.58E-12 | 12.9x |
| `Poisson_Survival.all.output_rel` | 2E-13 | 1.10E-10 | 550x |

`holdout_summary.md` still reports PASS because it embeds the previous
thresholds (2E-10 / 5E-10). Verified: exactly these two, no others.

This is an evidence inconsistency, not a demonstrated production defect. The
main grid measures the same functions at 2.56E-14 and 2.91E-14.

## Cause

Thresholds were derived from main-grid measurements alone. The holdout had
been the binding constraint three times previously in this project and was not
consulted.

The holdout is stale by **two** numerical improvements, not one: simulating the
old stopping criterion gives 1.22E-12 for the failing point, but the committed
value is 1.10E-10, so it also predates the CR-P1-02 prefactor repair.

## Expected outcome

Simulation of the failing point (k=100632, lambda=100000) under current source
predicts **7.45E-15**, a 27x margin under the 2E-13 threshold. If the refresh
lands near that, the thresholds stand and only the evidence was old.

## Closure criteria

1. Holdout re-exported from current source
2. `analyze_holdout.py` exits zero against the current registry
3. `holdout_summary.md` regenerated and committed
4. Any threshold change justified by holdout measurement, not widened to fit
5. CI regenerates and diff-checks the holdout summary (#9)

## Related finding

81 of 161 active contracts have no rows in the consolidated holdout grid,
including every regime added recently. Those were validated against held-back
points inside their own studies at freeze time, but the held-back points were
never promoted into `holdout_grid.csv`, so the ongoing independent check does
not cover them. Worth its own issue.

### #9 — Accuracy Gate does not run the holdout or the evaluator/manifest unit tests

- State: closed (completed)
- Labels: enhancement, testing, CI, P2, accuracy
- URL: https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/issues/9

#### Body

## Summary

The hosted Accuracy Gate validates the main grid thoroughly but never touches
the independent holdout, and does not run three of the repository's five Python
test suites.

Currently run:

| step | covers |
| --- | --- |
| `compute_errors.py` (strict) | main grid vs registry |
| regenerate + diff `accuracy_summary.md` | main summary freshness |
| `test_gate_degradation.py` | gate blocks when the reference helper is missing |
| `render_contract_table.py --write` + diff | benchmark README freshness |
| `check_source_thresholds.py` | source header drift (narrowly) |

Not run:

- `benchmark/test_contract_eval.py`
- `benchmark/test_manifest.py`
- `benchmark/holdout/test_analyze_holdout.py`
- `benchmark/holdout/analyze_holdout.py`
- diff of `benchmark/holdout/holdout_summary.md`

## This is not hypothetical

Two defects sat undetected in the repository until an external review found
them, both squarely inside the untested set:

1. **Stale holdout (#8).** The registry was tightened after #3 without
   re-exporting the holdout. `analyze_holdout.py` would have failed
   `Poisson_Cumulative` (12.9x over) and `Poisson_Survival` (550x over), while
   the committed summary reported PASS with the previous thresholds.

2. **Broken unit test.** `test_analyze_holdout.py` was failing. Its envelope
   fixtures used df 5E5 to mean "beyond the F cap", which stopped being true
   when `PROB_F_MAX_DF` rose to 1E10, silently inverting two assertions.

Neither was catchable by the existing gate: the manifest does not bind the
holdout grid, and CI does not execute either file. The repository could show a
green main accuracy state while committing an obsolete independent summary and
a failing test.

## Proposed steps

```yaml
      - name: Run evaluator and manifest unit tests
        working-directory: benchmark
        run: |
          python test_contract_eval.py
          python test_manifest.py

      - name: Test the holdout analyzer
        working-directory: benchmark/holdout
        run: python test_analyze_holdout.py

      - name: Verify the independent holdout against the current registry
        working-directory: benchmark/holdout
        run: |
          python analyze_holdout.py
          git diff --exit-code holdout_summary.md
```

All five suites and both analyzers pass at the current head, so these steps go
in green rather than requiring a fix first.

## Design note: what CI can and cannot catch here

The holdout diff catches a summary regenerated against changed thresholds. It
does **not** catch stale holdout *observations*, because nothing binds the
holdout grid to source the way `observation_manifest.json` binds the main grid.
A future change could still leave the holdout observations old while the
summary stays internally consistent.

Closing that properly needs the holdout grid brought under manifest
provenance - either a second manifest or one extended manifest binding both
grids, their exporters and their summaries. That is a larger change and belongs
in its own issue; the steps above are worth having regardless and would have
caught both of the defects above.

## Closure criteria

1. The five steps above run in the Accuracy Gate and block on failure
2. A deliberately stale `holdout_summary.md` fails the workflow (verify once)
3. Provenance for the holdout grid is either implemented or tracked separately

### #10 — Source headers and limitations registry contradict the code

- State: closed (completed)
- Labels: documentation, CI, P2, accuracy
- URL: https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/issues/10

#### Body

## Summary

An external review found five documentation claims that contradict the code
they sit beside. A sweep of the source found nine more of the same kind.

Every one was **correct when written** and made stale by a later measured
change. None affects behaviour, but a reader trusting the module header or the
limitations registry reaches a different conclusion than the code enforces -
and for a library whose value proposition is measured, honest accuracy, that is
the worst place to be wrong.

## The contradictions

| location | claim | reality |
| --- | --- | --- |
| `TFAMILY` header | `df above 1E5 are REJECTED` | `PROB_F_MAX_DF` is 1E10 |
| `TFAMILY` header | `validated to roughly 1E9` | the cap is 1E8 |
| `TFAMILY` header | `a larger df up to 1E100 is accepted and attempted` | the wrapper rejects past 1E8 |
| `TFAMILY` comments (x9) | `density is closed-form and unrestricted` | densities cap at 1E20 (`PROB_DENSITY_SHAPE_MAX`) |
| `CONTINUOUS` header | `balanced ... density and survival to 5E-15` | the frozen Beta density contract is 1E-14 |
| `CONTINUOUS` header | kernels `converge over ... roughly 1E9 and 1E7` | Beta now measures to 1E12 |
| `numerical_limitations.csv` | F rejects `above 1E5` | code accepts through 1E10 |
| `numerical_limitations.csv` | `F_Density is closed-form and unrestricted` | F density rejects above 1E20 |

Not a contradiction, and to be left alone: `StudentT_Density is unrestricted`
is correct - #6 concluded deliberately that it needs no envelope.

## Why the existing guard did not catch this

`check_source_thresholds.py` exists precisely to stop measured values being
restated in source. It recognises exactly one phrasing:

    <= <number> relative|absolute

None of the fourteen claims above uses it, so the checker reported clean while
the source contradicted the registry. **A checker that passes while the thing
it guards is broken is worse than no checker**, because it converts an open
question into a false assurance.

## Approach

Correcting the numbers would fix today's contradictions and guarantee
tomorrow's. The review's recommended policy is the actual fix:

1. `accuracy_contracts.csv` is the sole authority for measured thresholds
2. `numerical_limitations.csv` is the sole authority for public caps
3. Source headers describe the **algorithm** and point to those registries
   without repeating volatile numbers
4. The checker is widened to the phrasings that actually occurred

## Closure criteria

1. All fourteen contradictions resolved, by pointing at the registries rather
   than by restating corrected numbers
2. `check_source_thresholds.py` catches each of the original phrasings
3. No false positives on legitimate lines - in particular, a `Private Const`
   declaration must not be flagged, since it *is* the authoritative value
4. Gate green, tests green

### #11 — [Bug]: ICR-P1-01 — Positive intermediate underflow and subnormal precision loss in Gamma-family scale transformations

- State: open
- Labels: bug, P1, special-functions, t-family, continuous, accuracy
- URL: https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/issues/11

#### Body

## Description

Parent release blocker for the Gamma/Chi-square representability defects identified by independent review ICR-P1-01.

Two locally correct numerical layers compose incorrectly:

- a positive standardized argument can round to a low-significand subnormal or to zero before the incomplete-gamma kernel sees it;
- an inverse can underflow on the unit scale even though the final scaled quantile is representable.

The public API then returns a plausible support-boundary value, a spurious worksheet error, or zero for a representable quantile.

## Dependency status

- [x] #12 — PROB_TryLogGamma1p prerequisite kernel
- [ ] #13 — forward Gamma/Chi-square density, cumulative, and survival paths
- [ ] #14 — scaled inverse quantile reconstruction

Related incomplete-gamma blockers must also be resolved before final round-trip evidence is meaningful:

- [ ] #23 — validate the shared off-grid Stirling prefactor fix
- [ ] #24 — remove lower-region Q = 1 - P construction
- [ ] #26 — revalidate the continued-fraction Q breach after the shared prefactor fix

This parent closes only after #13 and #14 close and the final release evidence demonstrates their forward/inverse composition on the exact source.

## Confirmed counterexamples

For X = 1E-300, Shape = 1E-4, ScaleParam = 1E100, the mathematical standardized argument is 1E-400:

| quantity | current baseline | high-precision reference |
| --- | ---: | ---: |
| Gamma cumulative | 0 | 0.9120634760684959678789255 |
| Gamma survival | 1 | 0.0879365239315040321210745 |
| Gamma density | #VALUE! | 9.1206347606849596788E295 |
| Gamma inverse at the matching probability | 0 | 1E-300 |

For the smallest positive binary64 X and df = 0.0002:

| quantity | current baseline | high-precision reference |
| --- | ---: | ---: |
| Chi-square cumulative | 0 | 0.92824867943255041433481 |
| Chi-square survival | 1 | 0.07175132056744958566519004 |
| Chi-square density at df 1.99 | #VALUE! | 20.6892082696048338 |

Both input sets are inside the documented accepted domains.

## Root cause boundary

This is not a request to redesign CORE division semantics. PROB_TryDivide correctly returns the nearest binary64 result and valid underflow-to-zero is part of its contract. The defect is loss of mathematical-positivity and final-scale context across layers.

The repair belongs at the family call sites, supported by shared SPECIALFUNCS log-domain arithmetic:

- Gamma transforms X / ScaleParam;
- Chi-square transforms 0.5 * X;
- those transformations have different reachability and measured dispatch policies;
- the Loader/deviance path remains preferred near the distribution body.

## Closure

- [ ] #13 and #14 satisfy their frozen contracts and permanent regressions;
- [ ] all parent counterexamples return the correct value/error class;
- [ ] forward/inverse round-trips close in the corrected regime;
- [ ] final genuinely unrepresentable results retain the documented underflow/overflow policy;
- [ ] fitting-grid movement matches preregistration;
- [ ] independent holdouts pass;
- [ ] the full real-Excel regression and strict Accuracy Gate pass on the same source;
- [ ] no contract threshold is relaxed.

### #12 — [Bug]: ICR-P1-01 prerequisite — PROB_LogGamma(1 + X) destroys the increment below 2^-53

- State: closed (completed)
- Labels: bug, P1, special-functions, accuracy
- URL: https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/issues/12

#### Body

## 🐞 Description

`PROB_LogGamma(1# + X)` is unusable for small `X`. In binary64, `1# + X` rounds
to exactly 1 for every `X` below `2^-53` (1.11E-16), so the first-order term
`-EulerGamma * X` is lost. The result does not merely lose precision — it
returns a **positive constant** where the true value is negative and vanishing:

| X | `PROB_LogGamma(1 + X)` | true |
|---|---|---|
| 1E-13 | -5.3290705182008E-14 | -5.7721566490145E-14 |
| 1E-15 | **+3.5527136788005E-15** | -5.7721566490153E-16 |
| 1E-16 | **+3.5527136788005E-15** | -5.7721566490153E-17 |
| 1E-20 | **+3.5527136788005E-15** | -5.7721566490153E-21 |

The `+3.55E-15` floor is the rounding residue of `1 + X`.

This is **not** a subnormal-shape issue. Absorption begins at 2^-53, roughly 292
decimal orders of magnitude above the smallest normal Double, and degradation is
gradual well before that.

## Why it blocks ICR-P1-01

`LogGamma(Shape + 1)` appears in the forward CDF and survival log branches and,
critically, in the inverse:

    LogUnitQuantile = [LogProbability + LogGamma1p(Shape)] / Shape

The `/ Shape` amplifies the error. Measured quantile relative error using the
naive spelling, at probabilities that are ordinary Doubles:

| Shape | Probability | quantile relative error |
|---|---|---|
| 1E-06 | 0.999 | 1.78E-10 |
| 1E-10 | 0.9999999 | 3.88E-07 |
| 1E-14 | 0.99999999999 | **3.29E-02** |
| 1E-17 | 0.99999999999999 | **7.81E-01** |

`PROB_LogGammaDelta` requires `LargeArg >= 1` and `PROB_LogGammaHalfDiff` is a
half-integer shift, so neither covers this.

## Remediation

`PROB_TryLogGamma1p(X, Result, FailMsg)` in `M_STATS_PROBDIST_SPECIALFUNCS`,
evaluating the Maclaurin series in `X` so the increment never has to survive an
addition to one:

    Log(Gamma(1 + X)) = -EulerGamma*X + Sum(k=2..) (-1)^k Zeta(k) X^k / k

26-term fixed coefficient table, Horner, seam at `X = 0.25`. Term count and seam
chosen from measured error envelopes, not from the radius of convergence.

## Contract metric

Scaled absolute error, `Abs(observed - reference) / X`, because the scaled
inverse divides the result by `X`. Ordinary absolute error would be flattering
and would prove nothing about that caller.

- [ ] `LogGamma1p.small.scaled_abs` — measured worst **2.27E-16** over
      `X` in [3.9E-308, 0.25], worst at `X = 2.0E-16`
- [ ] seam continuity at `X = 0.25 ± 1 ulp` — measured jump 5.48E-14, bounded by
      `PROB_LogGamma`'s own 6.1E-14 contract and reducible only by improving it

## Limitation to document, not fix

Below `X ≈ 1E-308` the product `EulerGamma * X` is itself subnormal, so the
returned Double has no grid point near the answer. Scaled error then degrades as
`2^-1075 / X`: 1.3E-14 at 1E-310, 1.4E-04 at 1E-320, 4.2E-01 at the smallest
positive subnormal. No evaluation order can fix this.

- [ ] `LogGamma1p.SubnormalResultRepresentability` added to
      `benchmark/numerical_limitations.csv`
- [ ] `LogGamma1p.small.scaled_abs` scoped to `X >= 1E-308`

## Downstream reuse

- [ ] small-positive `PROB_LogGamma` via `LogGamma1p(Z) - Log(Z)`
- [ ] `PROB_StirlingError` small-`N` path, removing `(N + 1) / N`
- [ ] forward Gamma / Chi-square log branches
- [ ] scaled inverse quantile

Both `PROB_LogGamma` and `PROB_StirlingError` sit under Beta, F, Student t and
the discrete families. Land them one at a time with a compile and full `RunAll`
between, not as a single edit.

## Evidence

`benchmark/loggamma1p_study/`

### #13 — [Bug]: ICR-P1-01A — Gamma/Chi-square density, CDF and survival lose a positive standardized argument

- State: open
- Labels: bug, P1, special-functions, t-family, continuous, accuracy
- URL: https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/issues/13

#### Body

## Description

Child of #11. Six public surfaces lose the fact that their standardized argument was mathematically positive when the binary64 transform becomes a low-significand subnormal or zero:

- K_STATS_Gamma_Density
- K_STATS_Gamma_Cumulative
- K_STATS_Gamma_Survival
- K_STATS_ChiSquare_Density
- K_STATS_ChiSquare_Cumulative
- K_STATS_ChiSquare_Survival

The transformation is correctly rounded. The defect is downstream composition: the rounded standardized value is treated as an exact support coordinate.

## Current implementation state

Phase A measurement, independent holdouts, oracle cross-validation, and preregistration were completed against the pre-#23 source. They establish candidate policies, not final post-#23 cutoffs.

- [x] PROB_MIN_NORMAL exists in CORE, formed from a normal binary64 literal.
- [x] the 48-bit and 40-bit cutoffs are derived by exact power-of-two division; no handwritten subnormal literal is used.
- [x] PROB_PositiveRatioClass and PROB_ClassifyPositiveRatio exist in SPECIALFUNCS.
- [x] pre-#23 fitting and independent holdout evidence established the registered candidate cutoffs.
- [ ] the frozen `positive_ratio_subnormal` fitting arm and untouched holdout must be re-exported from post-#23 source before any cutoff is wired.
- [x] the predicted movement manifest is committed.
- [ ] the classifier and log-domain arithmetic are not yet wired into production family behavior.
- [ ] permanent public-surface regressions and machine-readable contracts are not yet complete.

The current scaffolding landed in two no-behavior-change commits: 9cac10ee7851faa654481089c42b1a0923134fdd adds `PROB_MIN_NORMAL` and the derived power-of-two cutoffs; db92bfc8e7b2447bff8c1cd5ba99244fe9e6a734 adds the three-state classifier. Neither commit changes returned values by itself.

## Registered pre-#23 candidate policy

| public surface | positive-subnormal dispatch | hard-underflow recovery |
| --- | --- | --- |
| Gamma density | StandardX < PROB_MIN_NORMAL / 16# | always |
| Gamma cumulative | StandardX < PROB_MIN_NORMAL / 4096# | always |
| Gamma survival | none | always |
| Chi-square density | StandardX < PROB_MIN_NORMAL / 16# | always |
| Chi-square cumulative | StandardX < PROB_MIN_NORMAL / 16# | always |
| Chi-square survival | none | always |

The strict less-than comparison was load-bearing in the pre-#23 study. PROB_MIN_NORMAL / 16# is 2^-1026, reproducing its 48-bit boundary; PROB_MIN_NORMAL / 4096# is 2^-1034, reproducing its 40-bit Gamma cumulative boundary.

These are candidate constants, not settled implementation constants. #23 changed the direct path through PROB_TryGammaPrefactor -> PROB_TryGammaLogPdf -> PROB_StirlingError, while the log-recovery path assembled with PROB_TryLogGamma1p is unchanged. At the old 48-bit boundary, argument rounding was about 8.9E-16 for Shape = 0.5 versus about 1E-14 from the repaired prefactor, so the direct-path error budget was dominated by the component #23 removed. At the old 40-bit boundary, argument rounding was about 2.3E-13 and the same prefactor component was much smaller. The predicted direction is therefore a lower crossover, especially for the 48-bit policies, but its magnitude must be measured.

Rerun the frozen fitting arm and untouched holdout on the exact post-#23 source. If a crossover moves, rederive the power-of-two cutoff and amend the preregistration before writing any production call site. If it does not move, record the post-#23 evidence that retains it.

Survival left the retained-bit dispatch problem after the study separated #23 and #24, but survival still needs hard-underflow recovery.

## Architecture

Dispatch at the six public-family call sites, where the original operands and mathematical positivity are still known. Keep the arithmetic in shared SPECIALFUNCS helpers.

Three cases must remain distinct:

~~~text
positive mathematical input, StandardX = 0  -> mandatory log-domain recovery
StandardX > 0 and below surface cutoff      -> measured precision dispatch
otherwise                                   -> existing direct Loader/deviance path
~~~

Family transforms must remain separate:

- Gamma: LogStandardX = Log(X) - Log(ScaleParam)
- Chi-square: LogStandardX = Log(X) - Log(2#)

Do not route Chi-square through a fictitious X / ScaleParam policy and do not move distribution logic into CORE.

The log-domain kernels must support:

- cumulative: assemble LogP with PROB_TryLogGamma1p(Shape), then exponentiate under the shared Try contract;
- survival hard-underflow: reconstruct Q stably from the direct lower-region log result after #24;
- density: assemble the log density from LogStandardX and the original scale/Jacobian;
- explicit failure if an internally assembled probability is materially outside its mathematical range.

## Dependency order

1. validate #23 with fresh real-Excel fitting and holdout evidence;
2. rerun the frozen `positive_ratio_subnormal` fitting arm and untouched holdout on the exact post-#23 source;
3. retain or rederive each cutoff and amend the preregistration before production wiring;
4. density and cumulative dispatch may then proceed because they do not depend on #24;
5. implement #24 before wiring survival hard-underflow recovery;
6. integrate and validate all six public surfaces before implementing #14 round-trip closure.

## Permanent evidence

Add/freeze the following six contract rows, without relaxing ordinary contracts:

- Gamma_Density.positive_ratio_subnormal
- Gamma_Cumulative.positive_ratio_subnormal
- Gamma_Survival.positive_ratio_hard_underflow
- ChiSquare_Density.positive_ratio_subnormal
- ChiSquare_Cumulative.positive_ratio_subnormal
- ChiSquare_Survival.positive_ratio_hard_underflow

If one contract schema cannot express both retained-bit and hard-underflow classes honestly, use separate named regimes rather than pooling them.

## Acceptance

- [ ] all six parent counterexamples pass;
- [ ] the post-#23 fitting arm and untouched holdout establish the final crossover for every affected surface;
- [ ] graded positive-subnormal cases route exactly at the post-#23 validated boundaries;
- [ ] hard-underflow cases distinguish mathematical positivity from the support boundary;
- [ ] no seam discontinuity exceeds its frozen regime contract within one ulp of a cutoff;
- [ ] Gamma and Chi-square policies remain independently validated;
- [ ] no new #NUM!, #VALUE!, monotonicity, or complement-consistency regression;
- [ ] observed movement matches the post-#23 preregistration, or every difference is explained and the manifest is amended before source wiring;
- [ ] fitting grid, frozen holdouts, full Excel regression, and strict Accuracy Gate pass;
- [ ] no frozen accuracy threshold is relaxed.

### #14 — [Bug]: ICR-P1-01B — Gamma/Chi-square inverse discards a representable quantile at the unit scale

- State: open
- Labels: bug, P1, special-functions, t-family, continuous, accuracy
- URL: https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/issues/14

#### Body

## Description

Child of #11. K_STATS_Gamma_InverseCumulative currently solves on the unit scale and multiplies by ScaleParam afterwards. A unit-scale quantile below binary64 range becomes zero before the wrapper can recover it, even when the final scaled quantile is representable.

For the parent counterexample, the unit quantile is 1E-400 and the final Gamma quantile is 1E-300. Returning zero is an implementation defect, not a representability limitation.

The equivalent Chi-square path has the same architectural requirement with a final scale factor of 2.

## Prerequisites

- [x] #12 — PROB_TryLogGamma1p exists for small-shape asymptotics.
- [ ] #13 — corrected forward paths are required before inverse tail-residual and round-trip evidence can close.
- [ ] #23/#24/#26 — incomplete-gamma forward errors must not contaminate inverse validation.

## Required architecture

Add a log-quantile kernel in SPECIALFUNCS, conceptually PROB_TryGammaInvPLog, which returns LogUnitQuantile without first materializing an unrepresentable unit-scale Double.

Required behavior:

- form a stable log probability from Probability and ComplementProbability;
- use Log(Probability) on the lower side and PROB_Log1p(-ComplementProbability) on the upper side;
- for the small-shape/small-quantile regime, use:

~~~text
LogUnitQuantile =
    (LogProbability + PROB_TryLogGamma1p(Shape)) / Shape
~~~

- select the log asymptotic before evaluating the Wilson-Hilferty seed;
- where the ordinary unit-scale solver is safe, reuse it and return Log(UnitQuantile);
- preregister the regime crossover and movement before changing production source;
- do not form Shape + 1 and call ordinary LogGamma in the small-shape branch.

Reconstruct exactly once at the public final scale:

~~~text
Gamma:     LogFinalQuantile = LogUnitQuantile + Log(ScaleParam)
Chi-square: LogFinalQuantile = LogUnitQuantile + Log(2#)
~~~

Exponentiate only the final log value under the shared overflow/underflow contract.

## Representability policy

- final quantile representable: return it;
- final quantile genuinely underflows: return zero;
- final quantile genuinely overflows: return #NUM!;
- never return zero only because the intermediate unit-scale value was unrepresentable.

Document final mathematical values outside binary64 separately as limitations:

- GammaInverse.FinalQuantileRepresentability
- ChiSquareInverse.FinalQuantileRepresentability

## Permanent contracts

Regime: scaled_quantile_from_underflowed_unit.

- [ ] Gamma_InverseCumulative quantile relative error
- [ ] Gamma_InverseCumulative tail-probability residual
- [ ] ChiSquare_InverseCumulative quantile relative error
- [ ] ChiSquare_InverseCumulative tail-probability residual

Both measures are required: quantile error tests final-scale reconstruction; tail residual tests composition with the corrected forward function.

## Acceptance

- [ ] the parent Gamma counterexample returns a representable nonzero quantile;
- [ ] equivalent Chi-square cases obey the same final-scale policy;
- [ ] forward/inverse round-trips satisfy frozen contracts throughout the corrected regime;
- [ ] one-ulp seam probes around the log/ordinary crossover remain monotone and within contract;
- [ ] permanent exact-binary64 VBA regressions are registered;
- [ ] final-underflow and final-overflow cases retain documented behavior;
- [ ] fitting grid and preregistered independent holdouts pass;
- [ ] the full real-Excel regression and strict Accuracy Gate pass;
- [ ] no existing inverse contract is relaxed.

### #15 — [Bug]: Replace the global PROB_LogGamma relative-error claim with regime-aware contracts

- State: closed (completed)
- Labels: bug, P1, special-functions, accuracy
- URL: https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/issues/15

#### Body

## 🐞 Description

`PROB_LogGamma` advertises "Relative error below 6.1E-14 across Z in
[1E-8, 1E+50]". That single global relative claim is wrong in two independent
ways. Pre-existing; found while validating `PROB_TryLogGamma1p` (#12), but not
part of ICR-P1-01 (#11).

## Subfinding A — genuine subnormal reflection defect

For small positive `Z` the reflection path forms `Sin(PROB_PI * Z)` after
`PROB_PI * Z` has entered the subnormal range and lost significand bits.
`Log(Sin(...))` is then off by that relative error *absolutely*, on a result of
magnitude ~700.

| Z | LogGamma(Z) | measured relative error |
|---|---|---|
| 1E-310 | 713.801379 | 1.20e-17 |
| 1E-315 | 725.314304 | 2.06e-13 |
| 1E-318 | 732.222061 | 7.81e-10 |
| 1E-320 | 736.827241 | 8.89e-08 |
| 4.94E-324 | 744.440072 | **6.19e-05** |

About 4.6E-02 of absolute error in log space at the smallest positive Double.
A real defect, and corrected by the `LogGamma1p(Z) - Log(Z)` branch in Phase 1.

## Subfinding B — relative metric is ill-conditioned near the zeros

`Log(Gamma(Z))` is zero at `Z = 1` and `Z = 2`, so a global *relative* contract
is ill-conditioned by construction. Measured 9.31E-14 relative at `Z = 1.75` —
about 7.9E-15 absolute on a value of magnitude 0.084. Not necessarily a kernel
defect; evidence that the metric is wrong.

## Replacement

Regime-aware contracts keyed on absolute error in the logarithm, since that is
what downstream callers propagate: the relative error of `Exp(v)` is
approximately the absolute error of `v`.

- [ ] `LogGamma.small_positive.log_abs` — small positive Z, including subnormal
- [ ] `LogGamma.near_zero.log_abs` — neighbourhoods of Z = 1 and Z = 2
- [ ] `LogGamma.general.output_rel` — where `Abs(LogGamma(Z))` is bounded away
      from zero, relative error remains useful

## Tasks

- [ ] withdraw the unqualified claim from the module header **before** #12's
      Phase 1 lands, so the subnormal improvement is not read as a regression
      caused by that edit
- [ ] grid: 1E-300, 1E-308, 1E-310, 1E-315, 1E-318, 1E-320, minimum positive
      subnormal; binary landmarks 2^-1022, 2^-1030, 2^-1040, 2^-1050, 2^-1060,
      2^-1070, 2^-1074; seam neighbours NextDown/NextUp of 0.25 and 0.5; zero
      neighbourhoods NextDown/NextUp of 1 and 2, plus 1.25, 1.5, 1.75
- [ ] `LogGamma(1)` and `LogGamma(2)` are exactly zero and may be evaluated for
      **absolute error only** — they cannot participate in a relative contract
- [ ] freeze thresholds only after the Phase 1 main grid and holdout are
      populated

Evidence: `benchmark/loggamma1p_study`.

### #17 — [CI]: Make accuracy-grid regeneration non-destructive and fail closed

- State: closed (completed)
- Labels: bug, CI, P2, accuracy
- URL: https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/issues/17

#### Body

> **Operative scope correction — 2026-08-30.** #17 is now the narrow P2 v1.0.0 safety gate: non-destructive default behavior, explicit authoritative writes and retirements, observation preservation, fail-hard duplicate canonical keys, fail-hard unexpected reference movement, documentation, and focused CI fixtures. Full origin reconstruction, generator/grid reconciliation, reference/observation redesign, and byte-stable clean rebuild moved to #32 under v1.01. Any broader historical checklist retained below explains the audit finding but is not a v1.0.0 closure criterion.

## Description

`generate_reference_values.py` documents itself as the constructor of
`probability_accuracy_grid.csv`. It is currently neither a reproducible
constructor nor safe to run, and separately the harness has no single
definition of what an input *is*.

Three confirmed problems. Found while auditing the reference-oracle migration
(#b4c5618), which needed to know whether the generator could be trusted as the
grid's source of truth.

> **Correction (superseded draft).** An earlier draft reported 563 grid-only and
> 238 generator-only rows, 236 stale references, and 254 rows of lost
> provenance. Those counts compared **decimal spellings**, not values. `0.85`
> and `0.84999999999999998` are the same Double and reach the library as the
> same Double, because the exporter parses with `Val`. All counts below are
> computed on IEEE-754 bit patterns.

> **Status update.** Section A is substantially complete and section B is
> largely resolved; section C is untouched. Counts below re-measured against the
> current grid (2091 rows, all carrying an observation), which has grown from
> 2030 since this issue was opened. See **Current state** at the foot.

## A. Row identity must be the Double, not the CSV text

**Implemented.** `reconcile_grid.py` keys on IEEE-754 bit patterns via
`struct.pack(">d", ...).hex()`. 222 matched rows differ in argument text while
being bit-identical Doubles; a string key would have counted every one of them
as coverage drift.

The canonical key is `function` + the four arguments as IEEE-754 bit patterns +
`regime` + `evidence_set`. Hex bit patterns rather than normalised decimal, so
formatting is removed from the identity question entirely.

Parity probe (Python `float`): `0.85`/`0.84999999999999998` →
`3feb333333333333`; `0.07`/`0.070000000000000007` → `3fb1eb851eb851ec`;
`0.2`/`0.20000000000000001` → `3fc999999999999a`;
`5E-324`/`4.9406564584124654E-324` → `0000000000000001`. The seam neighbours
`0.24999999999999994` / `0.25` / `0.25000000000000006` stay distinct at
`3fcffffffffffffe` / `3fd0000000000000` / `3fd0000000000001`, so the key
collapses spelling and still discriminates one ulp.

- [x] `reconcile_grid.py` keys on binary64 bits
- [ ] parser-parity test between Python `float` and VBA `Val` over: 0.85, 0.07,
      0.2 and their 17-digit round trips; min normal; representative
      subnormals; minimum positive subnormal; one-ulp boundaries; max finite
      Double. Parity is expected but should be proven once.

### A1. Residual coverage drift, on the corrected counts

Current reconciliation over 2110 keys:

| category | count |
|---|---|
| MATCH | 1335 |
| GRID_ONLY_EXPECTED | 356 |
| GRID_ONLY_UNEXPLAINED | 350 |
| METADATA_DIFFERENCE | 33 |
| GENERATOR_ONLY_UNEXPLAINED | 22 |
| REFERENCE_DIFFERENCE | 14 |

**22 generator-only rows**, all generator staleness — the grid was revised and
the script was not:

- 12 F rows at df `2.0e-16`, `1.98e-17`, `1.98e-13`, `2.0e-12`, far below any
  plausible domain; dropped when the F envelope was frozen (#3)
- 6 deep-tail inverse rows at probability `1.0e-12`, superseded by the
  `deep_tail` ladder `1e-15 … 1e-300` (`a710bb3`)
- 4 `DiscreteUniform_InverseCumulative` at a probability set the grid replaced

No evidence is missing in either case; the generator needs updating to match.

**350 grid-only unexplained rows**, all carrying an observation: 208 uniquely
attributable to a study folder by numeric argument match, 100 ambiguous across
two neighbouring studies (coincidence, not evidence), **42 unattributed**. By
function: `F_Density` 64, `Beta_Density` 46, `StudentT_Density` 30,
`F_InverseCumulative` 26, `ChiSquare_Density` 24, `Gamma_Density` 24,
`Beta_Cumulative` 15, `Beta_Survival` 15, `Gamma_Cumulative` 11,
`Gamma_Survival` 11, `F_Cumulative` 10, `F_Survival` 10.

Every NegativeBinomial row is accounted for. `f1077f6` introduced the 210 NB
grid rows and the 101 generator lines producing them in the same commit;
provenance was never lost, only invisible to a string key.

- [ ] the 22 generator-only rows are reconciled or the generator updated
- [ ] the 42 unattributed rows are explained
- [ ] `origin` is written by whatever promotes a row; retrospective matching is
      a lead, not a declaration, as the 100 ambiguous rows show
- [ ] the reconciliation column is named `probable_origin` while it is inferred

### A2. Row keys must be unique

**Still open.** 3 duplicate keys remain under the binary64 key, all
`LogChoose`, each pair byte-identical in every column. `build_rows()` iterates
`k = floor(n * frac)` over `frac in [0.0, 0.01, 0.5, 0.99, 1.0]`, and for small
`n` two fractions land on the same integer: `n=2` gives `0,0,1,1,2` and `n=10`
gives `0,0,5,9,10`. Three pairs, measured and weighted twice. Removal is
lossless.

`promote_grid_rows.py --retire FUNCTION REGIME COUNT --reason ...` now provides
the explicit deletion action this needs, so the blocker is gone.

- [ ] generator emits `sorted({int(mp.floor(n * frac)) for frac in ...})`
- [ ] the three rows are removed, byte-identity proven before the write
- [x] an explicit row-deletion action exists (`promote_grid_rows.py --retire`,
      requires function, regime, expected count and a reason)
- [ ] duplicate keys fail hard once uniqueness is established

## B. References must be evaluated at the Double the function receives

**Largely resolved.** `REFERENCE_DIFFERENCE` has fallen from **381 to 14**.
The reference-oracle migration and the canonicalisation work since this issue
was opened have re-evaluated the great majority of disagreeing rows at the
binary64 arguments the VBA functions actually receive.

The invariant stands: **every accuracy reference is evaluated at the exact
binary64 arguments the VBA function receives**, not at the decimal they were
written from. This matters most at contracts of 1E-15 to 1E-13, where input
rounding is the same order as the claimed output error. The VBA UDF never
receives the mathematical decimal `0.85`; it receives `Val("0.85")`.

Canonicalisation must happen *before* any derived quantity. Not

```python
pr = mp.mpf("0.85")
mean = r * (1 - pr) / pr      # derived from the exact decimal
```

but

```python
pr = mp.mpf(float("0.85"))    # the Double VBA will receive
mean = r * (1 - pr) / pr
```

so every generated evaluation point, bracket and moment reflects the same
Double-valued parameters. Applies to probabilities, shapes, scales, degrees of
freedom and evaluation points alike. Count inputs that are intentionally
integers stay exact where representable.

- [ ] shared `as_reference_arg` / `format_double_arg` / `double_key` helpers
- [ ] every public-function argument passes through canonicalisation before the
      reference function sees it
- [ ] canonicalisation precedes all derived quantities
- [ ] test comparing intended decimal, serialised decimal, binary64 bits and the
      promoted high-precision value
- [ ] the remaining **14** re-evaluated and classified REFERENCE_CORRECT /
      REFERENCE_STALE / ARGUMENT_TEXT_ONLY_DIFFERENCE / ORACLE_DIFFERENCE
- [ ] the 33 `METADATA_DIFFERENCE` rows classified — new category since this
      issue was opened, not yet triaged
- [ ] **no reference is patched before this is settled** — an earlier draft
      claimed regenerating would improve `NegativeBinomial_LogPMF` 120x; that
      may be entirely artificial and is withdrawn

## C. Full regeneration is destructive

**Untouched.** `main()` still writes a new CSV from `build_rows()` with
`observed_vba = ""`, and `--out` still defaults to
`probability_accuracy_grid.csv`. Running the documented command **blanks all
2091 committed observations**, independently of A and B.

The architecture should separate reference specification from observed evidence:

    generate_reference_values.py  ->  reference_grid.csv
    Excel execution               ->  observation_grid.csv
    deterministic merge           ->  probability_accuracy_grid.csv

The Python generator owns function, kernel, claim/metric, arguments, reference,
regime, evidence_set, expected_error. The Excel exporter owns key and
`observed_vba`, and never reads the reference. The merge tool owns row
correspondence and the committed grid.

### Non-destructive path

The `--patch-existing` flag proposed in the original draft was **not** built.
The need was met instead by `promote_grid_rows.py`, which is now the only
sanctioned route for adding or retiring a grid row and already enforces most of
the safety properties listed below. What remains is that
`generate_reference_values.py` itself is still unsafe to run as documented.

- [x] never delete an existing row except under an explicit, reasoned action
- [x] never alter `observed_vba` outside an export
- [x] fail when a requested generated row cannot be matched
- [ ] `generate_reference_values.py` is report-only by default; overwriting the
      committed grid requires an explicit flag
- [ ] its docstring stops describing a destructive command as the normal usage
- [ ] fail on duplicate keys
- [ ] report generator-only and grid-only rows

## Current state

Re-measured at `main`:

- grid 2091 rows, all carrying an observation (was 2030 when opened)
- 166 contracts, gate `FAIL 0 / KNOWN LIMITATION 0 / CHARACTERIZATION ONLY 1 /
  PENDING 0`
- `reconcile_grid.py`, `promote_grid_rows.py`, `derive_freeze.py`,
  `audit_references.py`, `audit_study_references.py`, `check_promotion.py` all
  exist; `refresh_evidence.py` regenerates the four artifacts and runs all
  seven checks
- #12 and #15 are closed, so the short-term path is no longer blocking them

**Remaining work is A2 (3 duplicates), the A1 reconciliation (22 + 42 rows), the
B residue (14 + 33 rows), and all of C.**

## Closure

- [ ] every committed main-grid row has a declared origin, reproducible from
      committed code
- [ ] full reconstruction produces exactly the committed key set under the
      binary64 key
- [ ] references regenerable independently of observations
- [ ] no reconstruction command blanks or overwrites observations accidentally
- [ ] the 22 generator-only and 42 unattributed rows are explained
- [ ] duplicate keys fail hard; row deletion requires an explicit action
- [ ] unexpected reference changes fail hard
- [ ] a clean rebuild is byte-stable except for documented normalisation
- [ ] CI verifies reconstruction/reconciliation


### #18 — [Bug]: PROB_StirlingError small-N recurrence overflows when 1/N exceeds Double range

- State: closed (completed)
- Labels: bug, P2, special-functions, t-family, continuous
- URL: https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/issues/18

#### Body

# [Bug]: PROB_StirlingError small-N recurrence overflows when 1/N exceeds Double range

## Summary

`PROB_StirlingError` handles `0 < N < 0.5` by forming the explicit reciprocal
`(N + 1#) / N`. Once `1 / N` exceeds the Double range the quotient is not
representable and the expression faults **before** `Log()` is called. The
parameter domain admits such N, so the fault is reachable from the public
surface.

This is **not** the `1 + N` precision hypothesis that originally motivated
looking at this function. That hypothesis was tested and rejected — see the
evidence section. The defect is the reciprocal.

## Boundary

`1 / DoubleMax = 2^-1024 * (1 + 2^-53)`. The excess over `2^-1024` is `2^-1077`,
an eighth of a subnormal ULP and below the representable grid, so `1 / DoubleMax`
rounds to exactly `2^-1024`. The decisive pair is therefore two adjacent Doubles:

| N | value | `(N + 1#) / N` |
| --- | --- | --- |
| `2^-1024 + 2^-1074` | 5.5626846462680084E-309 | DoubleMax exactly — last safe |
| `2^-1024` | 5.5626846462680035E-309 | not representable — first failing |

## Reachability

`PROB_CN_ValidateXShapeScale` delegates to
`PROB_IsPositiveWithinSupportedMagnitude`, which tests `X > 0` and an **upper**
magnitude bound of 1E100. There is no lower cutoff, normal or subnormal. A
subnormal shape validates, flows to `PROB_TryGammaLogPdf`, and reaches
`PROB_StirlingError` unchanged.

## Reproduction

```
K_STATS_Gamma_Density(1, 5.5626846462680035E-309, 1)   ->  #VALUE!
K_STATS_Gamma_Density(1, 5.5626846462680084E-309, 1)   ->  2.0464E-309
```

The true density at `X = 1, Scale = 1` is `Exp(-1) / Gamma(N) ~ N / e`: small,
finite and representable throughout. It is 3.68E-321 at `N = 1E-320`. Only at
`N = 2^-1074` does the correct answer underflow to zero.

## Observed effect

Runtime Error 6 `Overflow` propagates to `Err_Handler` and surfaces as
`CVErr 2015` (`#VALUE!`). **The failure is visible, not silent** — the caller
receives a refusal rather than a plausible wrong density. Severity is reduced
accordingly: this is a domain defect, not a wrong-answer defect.

## Affected surface

Reachable only where a shape below `2^-1024` can be supplied — the Gamma and
Beta log-pdf paths and what derives from them (Gamma, ChiSquare, Beta, F,
Student-t density/CDF/SF/inverse). The discrete Loader arrangements are **not**
affected: Binomial, Poisson, and `PROB_LogChoose` pass integer counts, and both
mass functions short-circuit at zero, so no argument in `(0, 0.5)` can arise.
Geometric does not call `PROB_StirlingError` at all.

## Evidence

`benchmark/stirling_overflow_probe/` — 14 points measured on real VBA at
`dfdbbbc`, extreme values constructed by halving rather than parsed from
literals.

Rejected precision hypothesis: 4,920 points against mpmath at 60 digits, worst
absolute error 4.828E-14 for both the current spelling and the `Log1p`
rearrangement, bit-identical at 82 per cent of the sample. Replacing the
spelling for precision reasons is not justified.

## Proposed fix

```vba
PROB_StirlingError = _
    PROB_StirlingError(n + 1#) + _
    (n + 0.5) * (PROB_Log1p(n) - Log(n)) - 1#
```

`PROB_Log1p` returns `X` unchanged once `1 + X` rounds to one, so the
replacement needs neither the increment to survive nor `1 / N` to be
representable. Measured finite through `2^-1074`, and **bit-identical to the
current result at every point where the current implementation works** except
`N = 0.25`, where it differs by 5.55E-17.

The direct formulation
`LogGamma1p(N) - (N + 0.5) * Log(N) + N - HALF_LOG_TWO_PI` is bit-identical
below 0.25 and better at 0.25 (2.14E-18), but requires a second branch seam at
`PROB_LG1P_SERIES_MAX` for an improvement to an error already three orders
inside contract. Documented, not adopted.

## Contract impact

None expected. `StirlingError.all.output` stays at 1E-13 absolute; worst
measured for the replacement is 4.828E-14, a 2.07x margin. Predicted grid
movement outside `N < 0.5` is zero. Any breach is to be investigated as a
regression, not accommodated by loosening the threshold — threshold changes
belong to #7.

## Margin note

`N = 1E-244` measures 4.828E-14, worse than the frozen holdout worst of
3.57E-14. Threshold holds; true margin is 2.07x rather than 2.8x. For #7.


### #20 — [Test]: Permanent exact-binary64 Stirling subnormal regression

- State: closed (completed)
- Labels: enhancement, testing, P2, special-functions, accuracy
- URL: https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/issues/20

#### Body

## Description

Permanent regression coverage for the Stirling reciprocal-overflow defect fixed in #18 is now implemented and verified.

The consolidated Excel harness constructs the relevant binary64 values arithmetically and exercises the exact production path. No subnormal VBA literal or Val-based parse is used.

## Implemented coverage

In Test_CN_GammaDensity:

- [x] construct the minimum positive Double by repeated stored halving;
- [x] derive 2^-1024 by exact repeated doubling;
- [x] verify K_STATS_Gamma_Density(1, 2^-1024, 1) returns the representable density;
- [x] verify the one-ulp-above shape remains finite;
- [x] verify PROB_StirlingError(2^-1074) returns the finite expected value;
- [x] fail if construction produces zero or any public call returns CVErr/runtime failure.

The comments explicitly preserve the VBA extended-expression pitfall: each halved result is assigned to a Double before comparison.

## Closure evidence

- production fix: #18 / 17ebb1c6c50119acd95b1f153ad9d5bd5112d48c
- permanent tests: tests/M_STATS_PROBDIST_TEST.bas
- latest verified exact-source Excel run containing these tests: 902 / 902 PASS
- environment: Excel 16.0 build 20131, 64-bit
- the stronger main-grid completeness mechanism is tracked by #22; it is not required to keep this exact boundary permanently covered because the real-Excel workflow now executes it on every source/test change.

All acceptance criteria for this test issue are complete.

### #21 — [Bug]: PROB_Expm1 returns values below -1 when Exp(X) is subnormal

- State: closed (completed)
- Labels: bug, P1, core, continuous, accuracy
- URL: https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/issues/21

#### Body

## 🐞 Description

`PROB_Expm1(X)` returns values **strictly below -1** when `Exp(X)` is subnormal,
roughly `-745.14 < X < -709.1`. Since `Exp(X) - 1` lies in `(-1, 0)` for every
finite `X`, the kernel returns a result outside its own mathematical range.

Two public UDFs propagate this directly as a **cumulative probability greater
than one**. `K_STATS_Exponential_Cumulative(745, 1)` returns
`1.0007521466129217`.

Fixed in `f7e71f0`. Filed after the fact so the defect, its reachable region and
its evidence are on the record; the fix commit references this issue.

## 🚦 Defect class

- [x] **Silent wrong value** — a plausible number is returned and it is wrong
- [ ] **Spurious error**
- [ ] **Wrong error code**
- [ ] **Unhandled runtime error**
- [ ] **Accuracy shortfall**
- [ ] **Performance or non-convergence**

More serious than #18 on this axis. #18 refused a valid input with `#VALUE!`,
which announces itself. This returns a plausible probability that violates the
range invariant, and nothing downstream can detect it: anything taking `1 - p`,
feeding an inverse, or accumulating the value gets nonsense with no error to
trace.

## 🔖 Version and source state

```text
Release tag:      N/A
Commit SHA:       a1de93f0dfb2a9a3d0c7c5b143db768d2c12b1dd  (defect present)
Fixed in:         f7e71f0                                    (see below)
Source obtained:  official repository, main
```

## 🔢 Exact call and result

```text
Function: K_STATS_Exponential_Cumulative(745.13, 1)
Returned: 1.000926774504277
Expected: 1
Status:   no error raised; a probability above one is returned silently
```

```text
Function: PROB_Expm1(-745.13)
Returned: -1.000926774504277
Expected: -1
Status:   outside the range of Exp(X) - 1, which is (-1, 0)
```

The mathematical CDF at `Lambda * X = 745.13` is still infinitesimally below
one; the correctly rounded binary64 result is exactly `1`.

### Argument bit-exactness

Not required. Every argument here is an ordinary normal-range decimal; the
subnormal quantity is the intermediate `Exp(X)`, not an input.

### Smallest runnable example

```vba
Option Explicit

Public Sub ReproduceIssue()

    Debug.Print PROB_Expm1(-745.13)                          '-1.000926774504277
    Debug.Print K_STATS_Exponential_Cumulative(745.13, 1#)   ' 1.000926774504277
    Debug.Print K_STATS_Weibull_Cumulative(745.13, 1#, 1#)   ' 1.0009267745042765

End Sub
```

## 🎯 Boundary localisation

The window opens where `Exp(X)` becomes subnormal and closes where it underflows
hard to zero, at which point the old `U = 0#` branch took over and returned the
correct `-1`. That is why the pre-existing assertions at `-746` and `-1E+300`
passed throughout and never surfaced this.

```text
Last good input above:  -709      returns -1.0, correct
First bad input:        -716      returns -1.0000000000000002
Peak error:             -745.13   returns -1.000926774504277
First good input below: -745.14   returns -1.0, correct
```

Error growth, measured on real VBA:

| X | `PROB_Expm1(X)` | relative error |
| --- | --- | --- |
| -709 | -1.0 | 0 |
| -716 | -1.0000000000000002 | 2.22E-16 |
| -724.1 | -1.0000000000000595 | 5.95E-14 |
| -731.2 | -0.9999999997568725 | 2.43E-10 |
| -738.1 | -1.0000003893543672 | 3.89E-07 |
| -742.6 | -0.999934945718797 | 6.51E-05 |
| **-745.13** | **-1.000926774504277** | **9.268E-04** |
| -745.14 | -1.0 | 0 |

## 🔬 Independent reference

```text
Reference system:   mpmath
Reference function: expm1(x)
Version:            1.4.1
Precision:          60 decimal digits
Expected result:    expm1(-745.13) = -1.0 to within 5E-324
```

Verified over 7,301 points across the window at 0.005 spacing.

## 🔁 Steps to reproduce

1. Check out `a1de93f`, import the six production modules, compile.
2. Run the three lines above in the Immediate Window.
3. Or import `benchmark/expm1_saturation/M_STATS_PROBDIST_EXPM1PROBE.bas` and
   run `Probe_Expm1Saturation`; compare against the committed
   `expm1_saturation_probe_prefix.csv`.

## 🧪 Environment

```text
Excel version:    16.0 build 20131
Office bitness:   64-bit
Operating system: Windows
Locale:           Italian
Use context:      worksheet formula, and VBA call
Workbook type:    .xlsm
```

## ✅ Regression-harness result

```text
Test_STATS_PROBDIST_RunAll            → pass
Test_STATS_PROBDIST_RunCore           → pass
Test_STATS_PROBDIST_RunNormalFamily   → pass
Test_STATS_PROBDIST_RunTFamily        → pass
Test_STATS_PROBDIST_RunContinuous     → pass
Test_STATS_PROBDIST_RunDiscrete       → pass
```

**Does the harness detect this defect?**

- [ ] Yes — a suite fails
- [x] No — every suite passes and the defect is still present

The CH3 characterization block asserted saturation at `-700`, `-746` and
`-1E+300` — all outside the window, on both sides of it. The defect sat in the
gap between them.

## 📋 Contract coverage

```text
Governing contracts:  Exponential_Cumulative.all.output  relative <= 5E-15
                      Weibull_Cumulative.all.output      relative <= 5E-15
Failing region covered by an active contract?  no
Grid rows in the failing region:               none
```

Both contracts exist and are frozen, but each carries **three grid rows**, the
deepest Exponential row at `Lambda * X = 1.5`. The failing region begins around
709 — roughly 500x beyond the deepest evidence point. **This is a coverage gap,
not a missing threshold**, and it is the reason the defect coexisted with a
green Accuracy Gate. Promoting far-tail rows into the two existing contracts is
tracked separately.

## 📐 Numerical region

- [ ] Central distribution body
- [ ] Lower tail
- [x] Upper tail
- [ ] Probability close to `0`
- [x] Probability close to `1`
- [ ] Support boundary
- [ ] Very small parameter
- [x] Very large parameter
- [x] Subnormal or denormal argument
- [x] Overflow or underflow region
- [ ] Parameter-validation boundary
- [x] Cancellation-prone region
- [ ] Inverse round-trip
- [ ] Moment calculation
- [ ] Error-code classification
- [ ] Performance or apparent non-convergence

## 🌐 Affected surface

```text
Reachable from:      K_STATS_Exponential_Cumulative  (-PROB_Expm1(-LambdaX), no clamp)
                     K_STATS_Weibull_Cumulative      (-PROB_Expm1(-PowerValue), no clamp)
Protected by clamp:  PROB_DS_TryGeometricCDF         (clamps to [0, 1] afterwards)
Provably unaffected: PROB_CN_TryWeibullLogVarianceFactor  (guarded 0 < Delta < 0.5)
                     PROB_LogExpm1                        (defined only for X > 0)
```

All five call sites were traced, not assumed. The geometric family was protected
by an incidental clamp rather than by design.

## 🛠️ Proposed fix

```vba
U = Exp(X)
V = U - 1#

If U = 1# Then          PROB_Expm1 = X
ElseIf V = -1# Then     PROB_Expm1 = -1#
Else                    PROB_Expm1 = V * X / Log(U)
End If
```

```text
Measured worst error:  0 across the window — correctly rounded, not merely improved
Bit-identical to current where current is correct?  yes
Expected contract impact:  none
```

Wherever `V` has rounded to `-1`, `-1` is already the correctly rounded value of
`Exp(X) - 1`, so the compensation is not merely inaccurate there but
unnecessary. The branch subsumes the old `U = 0#` case, introduces no cutoff
constant, and repairs the kernel rather than papering over two callers.

**A clamp on the callers would not have been a fix.** At `X = -731.2` and
`-742.6` the pre-fix result was *below* one while still wrong by 2.43E-10 and
6.51E-05, so a `<= 1` guard would have hidden the visible violation and left
those two silently wrong.

`V` must be a **stored** Double. VBA evaluates a Double expression in wider
precision and rounds only on assignment, so testing `U - 1#` inline would not
ask the question the branch depends on.

## 📎 Additional context

**Evidence.** `benchmark/expm1_saturation/` holds both real-VBA runs, before and
after, 16 points each. Post-fix every window point returns exactly `-1` with
zero error and both CDFs return exactly `1`. The six benign points from `-0.025`
to `-400` and the three past-window points are **bit-identical** between the two
runs — the repaired branch fires from `X <= -37.43`, far wider than the defect,
so proving it changes nothing outside the window was the point of including
them. A Python mirror over 8,582 points outside the window agreed: zero
differences.

**Grid movement.** One row of 2091 moved: `Geometric_Cumulative` at
`p = 0.001, k = 998`, by 1 ULP, from 6.878E-17 to 2.445E-16 relative against a
`rel <= 1E-9` contract — a margin of 4.1 million. That row is nowhere near the
saturation window; the shift is the stored intermediate `V` rounding where the
inline expression did not. It was **not** predicted: `PROB_DS_TryGeometricCDF`
had been set aside as "protected by its clamp" and dropped from the movement
analysis, but protected from the defect is not the same as unaffected by the
change.

**Permanent regressions added in the fix commit.** The CH3 `KNOWN DEFECT`
comment is replaced by the range invariant it blocked; a 13-point
negative-domain ladder asserts `-1 <= Expm1(X) < 0` from `-1E-8` to `-1E+300`;
and both public CDFs gain exact far-tail assertions. A ladder like that one
would have caught this on day one and costs nothing to run.


### #22 — [Bug]: Accuracy gate silently ignores main-grid rows without a matching contract

- State: open
- Labels: bug, testing, CI, P2, accuracy
- URL: https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/issues/22

#### Body

## Description

A row can exist in `benchmark/probability_accuracy_grid.csv` with a reference and a real-VBA observation while no registry contract matches its function/regime. Such a row contributes to no verdict and produces no warning.

The Accuracy Gate is strict inside each active contract, but it does not currently prove that every main-grid row is claimed. This is a fail-open coverage gap.

## Measured audit baseline

Recounted at main `bde92dd7037e4fde05e620745a1c54b0cbc3a261`:

- 2,088 total grid rows;
- 1,732 rows with `evidence_set = main grid`;
- 2,052 total rows match at least one registry function/regime;
- 36 main-grid rows match no contract at this audit baseline;
- 0 registry contracts lack matching grid rows.

The 36 unclaimed rows are:

| function | regime | rows |
| --- | --- | ---: |
| F_InverseCumulative | envelope_domain | 18 |
| ChiSquare_InverseCumulative | envelope_domain | 9 |
| StudentT_InverseCumulative | envelope_domain | 9 |

All 36 baseline rows carry `claim = measured`, a reference, an observation, and `evidence_set = main grid`. Their neighboring cumulative and survival surfaces already have `envelope_domain` contracts; the three inverse surfaces do not.

The previous body reported 2,091 / 2,055 from an older grid. The live counts above are authoritative for this issue baseline.

The number 36 identifies this immutable bde92dd audit snapshot; it is not a permanent live counter. After the transition guard lands, every current missing/known-debt count must be derived from the checked grid, registry, dispositions, and canonical fingerprint rather than copied into prose.

## Required design

Add a dedicated grid-claim completeness check. Do not infer deliberate exemption from the absence of a contract.

Each main-grid row must have exactly one explicit disposition:

1. claimed by one or more active/characterization contracts; or
2. covered by a named machine-readable exemption with a reason, owner, and review policy.

Missing disposition must fail CI.

Required outputs:

- [ ] fail on any main-grid function/regime with no contract or explicit exemption;
- [ ] fail on unknown exemption statuses, duplicate/conflicting exemptions, or exemptions that match no rows;
- [ ] report total rows, main-grid rows, claimed main-grid rows, explicitly exempt rows, and missing rows;
- [ ] name every missing function/regime and row count;
- [ ] run in `refresh_evidence.py` and the hosted Accuracy Gate;
- [ ] unit-test a clean grid, a lost contract, an explicit exemption, a stale exemption, and a non-main study row;
- [ ] include the same counts in `accuracy_summary.md` for human audit.

## Pre-freeze contract and regime audit

The gap is broader than 36 unclaimed rows. Exact `(function, regime)` matching currently hides stale or overlapping public domain claims.

| function | existing claim | measured/structural conflict | required correction |
| --- | --- | --- | --- |
| `StudentT_InverseCumulative` | `all` relative at 3.0E-12 over “full tested range” | the nine large-df rows reach 1.8480E-10; five already-claimed `all` rows also have exact reference zero at p = 0.5 | narrow `all` to the nonmedian core regime; move every exact median to the selected absolute median regime; add a distinct large-df envelope claim |
| `ChiSquare_InverseCumulative` | `all` relative at 4.7E-12 over “full tested range” | the large-df rows are accurate, but a separate `envelope_domain` regime makes the broad prose claim structurally ambiguous | narrow `all` to the core df regime and make the large-df contract explicitly additive |
| `F_InverseCumulative` | paired `validated` contracts at 2E-10, documented as both df <= 1E5 and “enforced by PROB_F_ValidateEnvelope” | production accepts through 1E10; eight main-grid `validated` rows exceed 1E5, and eight current holdout rows use `(df1, df2) = (1.5, 2E5)` / `(5E5, 4)` | retain <=1E5 only as the registry’s core-regime boundary, remove the false runtime-enforcement claim, move the mis-scoped rows to the large-df envelope, and validate that envelope independently |

The two `F_InverseCumulative.validated` registry rows are not duplicates. They are the intentional `quantile_error/relative` and `tail_probability_residual/relative` pair, have distinct contract IDs and measures, and both count among the current 166 contracts.

Runtime authority was checked directly at `b00fa0e`: `PROB_F_MAX_DF = 1E10`, and `K_STATS_F_InverseCumulative` passes that constant to `PROB_F_ValidateEnvelope`. The source comment also records that `benchmark/inverse_probe` measured the inverse kernel clean through df = 2E10 with no refusals, including the unbalanced ratios previously assumed to fail. The public function therefore accepts df = 1E6 and rejects only when either df exceeds 1E10. The eight main-grid rows are valid production inputs. The inaccurate statement is the registry text: `PROB_F_ValidateEnvelope` enforces the overall 1E10 public cap, not the <=1E5 `validated` core regime. Remove “enforced by PROB_F_ValidateEnvelope” from all four affected F `validated` domain strings—CDF, survival, inverse quantile, and inverse tail—and describe <=1E5 as contract-regime ownership instead.

There is a real grid duplication nearby: the numerical inputs `p = 0.5, 0.9, 0.99; df1 = 1E6; df2 = 3` each appear once as `validated` and once as `envelope_domain`. Keep the higher-precision `inverse_probe` envelope rows, remove the three older duplicate `validated` rows, and reclassify the five remaining unique main-grid rows above 1E5. Reclassify the eight already-inspected holdout rows at `(df1, df2) = (1.5, 2E5)` and `(5E5, 4)`, four probabilities each, as legacy supporting envelope evidence; they do not count as the untouched independent holdout for the new contracts.

The established 2E-10 F core thresholds are not retightened merely because their old binding rows move. Their notes and domain prose must be corrected, while any future tightening requires its own fitting/holdout decision.

### Deterministic regime ownership

Domain prose is not a router. Add one machine-readable inverse-regime classifier and use it in generation/promotion checks and CI to prove that each stored regime agrees with its arguments. The classifier must implement these disjoint priorities:

1. `F_InverseCumulative.tiny_unbalanced_representable` takes precedence under its existing representability predicate.
2. Otherwise F is `validated` when both df are <= 1E5, and `envelope_domain` when both are within the public 1E10 cap and at least one exceeds 1E5.
3. Student-t exact medians use the selected absolute median regime for every accepted df.
4. Student-t nonmedian rows are `all` below df = 1E6 and enter the selected large-df regime from df = 1E6 through the public 1E8 cap.
5. Chi-square rows are `all` below df = 1E6 and `envelope_domain` from df = 1E6 through the public 1E8 cap.

Under the expected relative route, `StudentT_InverseCumulative.envelope_domain` owns every accepted large-df probability except exactly p = 0.5. Under the fallback, the frozen median-neighborhood band owns its retained nonzero points and `envelope_domain` owns every accepted large-df point outside it. There must be neither a gap nor an overlap.

Lock the routing with boundary and mutation fixtures: p = 0.5, its immediate binary64 neighbours, p = 0.6, the selected median-band edge and its adjacent doubles, df = 1E6, and the F 1E5/1E10 seams. A wrong stored regime must fail CI even if some contract happens to share its function name.

### Transitional coverage-delta guard

The cheap detector must land before further value-changing commits, but the heavy reference study is not a prerequisite for #23.

Until #22 closes, CI may run a clearly labelled development transition mode against a canonical binary64 fingerprint of the 36-row bde92dd audit-baseline debt set. That mode:

- reports all remaining rows as `KNOWN COVERAGE DEBT`, never PASS or claimed;
- permits only a subset of the frozen baseline debt, so resolving rows is allowed;
- fails on any new unclaimed row, changed identity, new function/regime pair, or edited baseline without an explicit reviewed baseline change;
- exposes the current remaining count in the generated summary; no live count is copied from the audit-baseline prose;
- has an owner, milestone expiry, and a hard prohibition on surviving the #22 closure/release commit.

Strict mode continues to fail while even one row lacks a real contract or final explicit exemption. The transition fingerprint is neither an accuracy contract nor the final exemption path; it is an anti-regression fence that decouples schedule without pretending the recorded debt set is dispositioned.

The Phase 0 guard commit owns the initial audit-baseline fingerprint. No later direct-main commit may remove duplicates, relabel regimes, or otherwise change a fingerprinted row identity by itself. The intended final #22 closure commit owns the F row transition: it lands the governing contracts and row cleanup atomically and deletes the fingerprint instead of refreshing it to a larger unclaimed set. If an earlier reviewed row migration becomes unavoidable, that one commit must update the canonical fingerprint and generated current-debt count atomically and document the exact old/new identity mapping.

## Disposition decision and candidate contracts

All 36 baseline unclaimed rows will be claimed by real active contracts. None will be exempted or retired. The three deleted F rows are older claimed duplicates, not part of the 36.

Eight Student-t rows have the exact mathematical reference zero: five existing `all` rows at df = 1, 2, 5, 30, 1000 and the three `envelope_domain` rows at df = 1E6, 1E7, 1E8. All observe `0E+000;0E+000`. Following the #15 `LogGamma(1)` / `LogGamma(2)` precedent, none may participate in a relative quantile contract.

The regime-specific measured errors on the 36 unclaimed rows are:

| function | rows | current main-grid worst relative quantile error |
| --- | ---: | ---: |
| F_InverseCumulative | 18 | 9.3115E-9 |
| StudentT_InverseCumulative | 9 | 1.8480E-10 |
| ChiSquare_InverseCumulative | 9 | 1.0890E-15 |

The observation authority is the compensated two-part export `hi;lo`, parsed as `hi + lo` by `benchmark/_contract_eval.py::parse_observed`. Scoring `hi` alone produces the false Chi-square worst 5.5455E-15; the gate-correct nine-row worst is 1.0890E-15 at p = 0.9, df = 1E7.

For a requested probability `p` and recovered forward probability `p_hat`, tail-relative residual is `Abs(p_hat - p) / Min(p, 1 - p)`, using the exact binary64 input probability and summed quantile.

F and Chi-square each receive separate `envelope_domain.quantile_rel` and `envelope_domain.tail_rel` contracts.

### Student-t zero-reference decision study

The existing zero-reference evaluator rule—zero observed against zero reference scores zero, while nonzero observed against zero reference scores infinity—prevents a crash, but it does not make relative error well-conditioned at the exact median.

#### Code-path prediction before measurement

The current source predicts meaningful relative scoring for nonzero neighbours:

- the public function returns zero explicitly only when p = 0.5;
- for p > 0.5, it forms `1 - p`, exact in this interval by Sterbenz;
- `PROB_NormalInvCDFRaw` forms `Probability - 0.5`, also exact here;
- its central seed has a permanent published relative bound of 1.15E-9;
- the forward tail can round the Newton residual to zero before refinement improves the seed.

Predicted outcome: finite, roughly seed-level relative error on nonzero neighbours, not an intrinsic conditioning failure.

#### Frozen fitting and holdout arms

Let one upper-side ulp at 0.5 be `U = 2^-53`.

The median-neighbour fitting arm uses exact offsets `2^-53, 2^-43, 2^-33, 2^-23`; the untouched holdout uses disjoint interleaved offsets `2^-48, 2^-38, 2^-28, 2^-24`. Both sides of 0.5 and df = 1E6, 1E7, 1E8 are included, with references generated at the actual binary64 probabilities.

To prevent a semantic hole between the largest median offset and the existing p = 0.9 rows, add exact dyadic bridge probes. Fitting uses p = 0.625, 0.75, 0.875 and their complements; the untouched holdout uses p = 0.5625, 0.6875, 0.8125, 0.9375 and their complements, again across the preregistered large-df envelope. These bridge points belong to the large-df threshold evidence and are frozen before observation. The p = 0.6 fixture separately proves routing for a future promoted row.

#### Preregistered decision rule

**Relative route — expected, seven new contracts total**

Select this route only if every nonzero fitting point returns a nonzero quantile and every relative error is finite and no greater than 5E-9. Freeze the final threshold as the tightest strict 1-2-5 value, not above 5E-9, that satisfies the ordinary fitting-headroom rule. Trends are reported but cannot override the numeric rule.

Then:

- every nonmedian large-df Student-t row uses `envelope_domain.quantile_rel` and `envelope_domain.tail_rel`;
- all eight exact medians move to `StudentT_InverseCumulative.median_exact.quantile_abs` at threshold zero;
- the existing `all` contract retains only nonmedian core rows;
- no median tail contract is needed because a zero-threshold quantile contract is stronger.

Together with the two F and two Chi-square contracts, this is seven new contracts.

**Absolute fallback — only if relative scoring fails, eight new contracts total**

If the relative gate fails, fitting derives a candidate `T_abs`. Retain only offsets satisfying `Abs(q_reference) >= 100 * T_abs`; apply the same cutoff to the still-uninspected holdout before revealing observations. This guarantees a zero-return mutant fails by at least two orders of magnitude.

Then:

- every exact median and retained large-df neighbour uses `StudentT_InverseCumulative.median_neighborhood.quantile_abs`;
- the retained nonzero neighbours also use `median_neighborhood.tail_rel`;
- nonmedian points outside the frozen band use `envelope_domain.quantile_rel` and `envelope_domain.tail_rel`;
- permanent exact-zero VBA assertions enforce p = 0.5 -> 0 exactly;
- the note states that the absolute band records a measured implementation property, not a mathematical limit.

Together with the two F and two Chi-square contracts, this is eight new contracts.

### Exact-zero threshold handling

A threshold of zero is schema-representable and the current Decimal comparison can evaluate it, but it is unprecedented in the registry. The generic 1-2-5 derivation is meaningless at zero: the current helper would start at 0.1 rather than expressing exactness.

Add an explicit exact-invariant derivation mode. It may emit threshold zero only when every governed fitting reference and observation is exactly zero and the preregistration names an exact mathematical invariant. Ordinary empirical contracts must continue through nonzero 1-2-5/headroom derivation.

Required tests cover:

- parsing threshold `0`;
- exact zero observed/reference at threshold zero -> PASS;
- nonzero observed against zero reference at threshold zero -> FAIL;
- main-grid and independent-holdout verdicts;
- exact-invariant freeze derivation and rejection of an empirical zero shortcut;
- generated summary and README contract-table formatting preserving `0`;
- regime/disposition reconciliation with a zero-threshold contract.

## Reference and evaluator work

The current independent holdout contains zero `envelope_domain` rows for any function. Phase 0 must preregister new fitting and untouched holdout evidence for all three inverse surfaces.

For Chi-square:

- reuse or factor the converging-route incomplete-gamma generator from `benchmark/cdf_large_shape`;
- use the lower series below `A + 1` and Lentz continued fraction above;
- recompute at materially higher precision and reject unstabilized references;
- cross-check fitting and holdout references independently with Rmpfr;
- do not use stock `mpmath.gammainc` as authority at df = 1E8.

This generator has a hard early feasibility checkpoint before the numerical plan enters Phase 2, although it is not a #23 closure prerequisite. It must produce the frozen Chi-square reference set at df = 1E6, 1E7, and 1E8; reproduce it at two materially separated working precisions; pass an independent Rmpfr cross-check; and record in #22 the maximum mpmath-route-versus-Rmpfr agreement, compared measures, precision pair, convergence status, and runtime. Failure to satisfy that checkpoint blocks Phase 2 while there is still schedule room to change the reference strategy.

## Maintainer decision — df = 1E8 reference substitution accepted

Accepted on 2026-08-30 from commit e68cc82545ead9777b155751e51976e63f3f4243. Quadrature and the converging gmpy2/MPFR route cover 69/69 points with minimum agreement of 109.87 and 115.11 significant digits respectively. Rmpfr agrees to at least 115.96 digits on all 46 feasible points but aborts above shape approximately 4.5E7; df = 1E8 requires shape 5E7.

For df = 1E8, the accepted substitute is the converging series/CF route in gmpy2 MPFR arithmetic, cross-checked against algorithm-independent quadrature. Rmpfr remains the independent third-party check where it can execute. This closes oracle feasibility only; it does not freeze a production threshold or authorize holdout inspection.

For Student-t and F, use the repository's robust incomplete-beta route with exact public transformations, higher-precision stability checks, and independent Rmpfr cross-checks.

Extend both the main-grid evaluator and `benchmark/holdout/analyze_holdout.py` so `tail_probability_residual` is scoreable for Chi-square and Student-t inverses, not only Beta/F. Add clean and deliberately degraded fixtures, compensated-observation fixtures, zero-reference/zero-threshold cases, wrong-regime cases, bridge and seam routing, and unavailable/unconverged oracle failures.

Holdout points must be non-overlapping and frozen before observation. Thresholds are frozen from fresh current-source fitting before the untouched holdout is inspected; holdout failure triggers investigation, not widening.

## Scope boundary

This issue guarantees that evidence already called main grid cannot be ignored and that the three affected inverse domains are partitioned consistently. It does not prove that every numerically important region has rows. Exact Stirling boundary coverage is completed in #20; new numerical regimes remain the responsibility of their owning issues.

## Sequencing and atomicity

1. Phase 0 lands the cheap checker, strict mode, fixtures, and the temporary fingerprint of the 36-row audit-baseline debt set. Development CI derives the current debt count, fails on any unauthorized coverage regression, and still reports the known debt visibly.
2. #23 proceeds once #29’s Excel-free binding mechanism and a controlled Excel route are ready. #23 does not wait for the #22 gamma reference generator, Student-t fitting decision, or new inverse contracts.
3. In parallel with #23, preregister the neighbour and bridge arms; retain the accepted Chi-square quadrature/gmpy2/Rmpfr evidence; add the deterministic regime classifier; and extend both evaluators. Holdout observations remain hidden. The df = 1E6/1E7/1E8 feasibility checkpoint is already satisfied at e68cc82.
4. After the last inverse-affecting numerical source change—no earlier than #14—run the #22 fitting arm on the release-candidate source, select the seven- or eight-contract route exactly as preregistered, freeze thresholds, and only then inspect the untouched holdout.
5. Atomically land the selected contracts; corrected F domain strings; all eight Student-t median relabels; five F main-row reclassifications; removal of three older duplicate F rows; the eight F holdout reclassifications at df 2E5/5E5; promotions; classifier; strict checker; generated summary counts; negative fixtures; and CI wiring. This final #22 closure commit owns the F row transition and deletes the temporary debt fingerprint rather than refreshing it to a larger unclaimed set.
6. Close #22 before #28 generates public assurance metrics and before #31’s final certification. Its evidence wave should be combined with the final release-candidate Excel export where practical, rather than inserted into #23’s critical path.

If no controlled Excel route is available, Excel-free work may continue but no evidence-dependent closure or accumulated value-changing numerical sequence proceeds beyond the plan’s Excel contingency.

## Acceptance

- [ ] all 36 baseline unclaimed rows are claimed by the selected seven- or eight-contract design;
- [ ] the two F `validated` registry rows remain recognized as the intentional quantile/tail pair, not duplicates;
- [ ] the three duplicated F numerical inputs are consolidated, all eight mis-scoped main-grid F rows and the eight mis-scoped holdout rows at `(1.5, 2E5)` / `(5E5, 4)` are correctly handled, and the core/envelope domains are disjoint;
- [ ] all four F `validated` domain strings stop claiming that <=1E5 is enforced by `PROB_F_ValidateEnvelope`; production’s actual 1E10 runtime cap remains documented separately;
- [ ] Student-t `all` and Chi-square `all` domain strings are narrowed atomically with the new regimes;
- [ ] all eight exact Student-t medians are excluded from relative quantile scoring and have permanent exact-zero VBA assertions;
- [ ] the relative route is selected only under the frozen 5E-9/nonzero/finite rule; fallback uses the frozen `Abs(q_reference) >= 100 * T_abs` cutoff;
- [ ] the median-neighbour and dyadic bridge fitting/holdout arms remain disjoint and uninspected until their preregistered stage;
- [ ] the machine-readable classifier has no gap or overlap and routes p = 0.6 plus every seam fixture correctly;
- [ ] observations use `hi + lo`, and tail residuals normalize by `Min(p, 1 - p)`;
- [x] the frozen Chi-square gamma references stabilize at two precisions; quadrature and gmpy2/MPFR converge 69/69, and Rmpfr agrees on all 46 feasible points; Student-t/F beta reference work remains subject to its preregistered fitting stage;
- [x] before Phase 2, the Chi-square generator produces df = 1E6/1E7/1E8 references stable across two working precisions; the agreement, convergence status, precision, runtime, and Rmpfr ceiling are recorded;
- [ ] main-grid and holdout evaluators score the new measures and fail closed on unsupported or unconverged references;
- [ ] threshold-zero pass/fail, derivation, formatting, and reconciliation tests pass;
- [ ] before closure, the transitional guard derives and reports the exact current known debt and fails on any unauthorized new or changed unclaimed row;
- [ ] the checker, contracts, row changes, dispositions, summaries, fixtures, and CI wiring land atomically;
- [ ] the final #22 closure commit owns the F row transition and deletes the fingerprint with governing contracts present; no intermediate direct-main commit leaves a refreshed larger unclaimed set;
- [ ] the temporary coverage-debt fingerprint is deleted and strict mode reports zero main-grid rows with a missing disposition;
- [ ] the summary records the 2,088 / 1,732 baseline separately from the post-consolidation/promotion totals;
- [ ] deleting a governing contract or corrupting a regime/exemption makes CI fail;
- [ ] no frozen threshold is relaxed merely to close the coverage gap.

### #23 — [Bug]: Off-grid PROB_StirlingError error floor breaches Gamma/Chi-square cumulative contracts

- State: open
- Labels: bug, P1, special-functions, t-family, continuous, accuracy
- URL: https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/issues/23

#### Body

## Description

The original real-VBA counterexample exposed a 3E-15 Gamma cumulative contract breach in PROB_TryGammaSeriesP at small shape. Decomposition showed that the converged series sum was not the error source.

The broader root cause is the off-grid PROB_StirlingError path used by PROB_TryGammaPrefactor. Its former LogGamma identity contributed roughly 1E-14 absolute error to the Stirling correction. That is acceptable under StirlingError's own 1E-13 absolute contract but becomes roughly the same relative error in an incomplete-gamma prefactor, breaching the tighter public CDF contract.

## Minimal counterexample

~~~text
K_STATS_Gamma_Cumulative(0.5, 0.0001, 1)
reference: 0.9999440197070425982869658
baseline relative error: about 9.43E-15
contract: 3E-15 relative
~~~

Equivalent Chi-square route:

~~~text
K_STATS_ChiSquare_Cumulative(1, 0.0002)
    -> P(A = 0.0001, X = 0.5)
~~~

The scale transform is a no-op in the Gamma case, so #13 is not involved.

## Exact source path

~~~text
PROB_TryGammaSeriesP
  -> PROB_TryGammaPrefactor
     -> PROB_TryGammaLogPdf
        -> PROB_StirlingError
~~~

Measured decomposition:

- ascending series contribution: about 2E-17 to 3E-16 relative;
- prefactor contribution: about 8E-15 to 1.3E-14;
- public P follows the prefactor error almost one-for-one.

Preregistered scope probes established that the problem is not limited to shape below 0.5; ordinary off-grid shapes 0.6, 5.3, and 10.3 were predicted to expose the same composition floor.

## Implementation status

The smallest root fix is merged on main at bde92dd7037e4fde05e620745a1c54b0cbc3a261 through PR #27.

PROB_StirlingError now:

- retains the exact half-integer table at N <= 15;
- recurs off-grid N <= 15 upward to the existing asymptotic region;
- retains the existing N > 15 asymptotic path;
- preserves the #18 subnormal-safe recurrence.

Seven preregistered VBA assertions cover the kernel, equivalent Gamma/Chi-square public points, and broader off-grid shapes. No contract threshold changed.

## Why the issue remains open

Implementation is on main, but release evidence for that implementation is not complete.

Current state:

- [x] minimal counterexample and root cause established;
- [x] exact source path decomposed;
- [x] predicted movement and independent holdout design committed before source change;
- [x] smallest source fix implemented;
- [x] permanent VBA regressions added;
- [x] existing #13 analyzer and evidence-consistency checks still pass;
- [ ] the frozen `positive_ratio_subnormal` fitting arm and untouched holdout have not yet been re-exported from post-#23 source, so #13's 48-bit and 40-bit cutoffs are not yet revalidated;
- [ ] real Excel has not yet verified the expected 909 / 909 assertions;
- [ ] main fitting observations have not been re-exported from the changed source;
- [ ] the frozen independent #23 holdout has not been executed on the implementation;
- [ ] representative rows have not been promoted to the main grid;
- [ ] observation_manifest.json and accuracy_summary.md are stale by design;
- [ ] strict hosted Accuracy Gate is therefore red;
- [ ] #26 has not been remeasured even though it shares the repaired prefactor.

Current main is 0dd748884599d4d0da815cb53eeceb13efd51f05; bde92dd7037e4fde05e620745a1c54b0cbc3a261 remains the numerical source baseline. Accuracy Gate runs 166 through 171 preserve the intended Phase 0 failure mode: only SPECIALFUNCS and the consolidated test module mismatch the last source binding. Excel VBA Regression run 95 was cancelled because the self-hosted runner was unavailable; no Excel result was produced.

## Closure sequence

1. before exporting, reproduce the exact frozen seven-line/two-file Phase 0 signature; verify #29's Excel-free manifest mechanism and #22's transition guard;
2. run the full real-Excel regression on the exact current modules;
3. execute only the preregistered #23 fitting grid and untouched #23 holdout, then export the current global holdout;
4. write and verify the first truthful #29 source binding immediately; close #29 only after deliberate source/grid mismatches fail;
5. compare #23 movement with MOVEMENT_MANIFEST.md and add/freeze representative #23 main-grid rows under the existing public contracts;
6. remeasure the X = 2 continued-fraction ladder for #26;
7. rerun the frozen `positive_ratio_subnormal` fitting arm and untouched holdout, compare direct versus log-recovery errors, and retain or rederive #13's cutoffs before any production wiring;
8. re-export the complete current accuracy grid;
9. write fresh main-grid and holdout provenance atomically, then regenerate summaries;
10. require all seven former expected-red checks to PASS, plus #22's transition mode with no new or changed debt;
11. close #23 only if all frozen numerical contracts pass without relaxation and no new unclaimed row was introduced.

The #22 reference generator, median study, bridge points, and contract creation run on their own parallel track. They are not #23 prerequisites and close later on the release-candidate source.

## Acceptance

- [ ] minimal Gamma and equivalent Chi-square calls satisfy the existing contracts;
- [ ] all preregistered off-grid scope probes satisfy the existing contracts;
- [ ] post-#23 #13 cutoff revalidation is recorded, including any amended preregistration;
- [ ] 909 / 909 real-Excel assertions pass, or any different count is fully explained;
- [ ] fitting and untouched independent holdout pass;
- [ ] main-grid evidence permanently reaches the former gap;
- [ ] no series, continued-fraction, inverse, Poisson, large-shape, or unrelated contract regresses;
- [ ] #29 has demonstrated a truthful current-source holdout binding and is closed;
- [x] every completed Phase 0 commit through 0dd7488 preserved the exact known-red signature;
- [ ] the Phase 1 evidence commit retires the expected-red waiver with seven PASS results;
- [ ] #22's transitional guard confirms that #23 introduced no new or changed unclaimed row; #22's final inverse contracts and strict zero-missing closure are explicitly deferred to its release-candidate evidence wave;
- [ ] source-bound evidence and hosted CI are green;
- [ ] Gamma_Cumulative.all.output remains 3E-15 relative.

### #24 — [Bug]: PROB_TryGammaRegularizedQ amplifies lower-series error by forming Q = 1 - P

- State: open
- Labels: bug, P1, special-functions, t-family, continuous, accuracy
- URL: https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/issues/24

#### Body

## Description

In the lower incomplete-gamma region, PROB_TryGammaRegularizedQ currently obtains survival by computing P and then forming:

~~~text
Q = 1 - P
~~~

That construction contradicts the direct-survival contract. Any absolute error in P becomes the same absolute error in Q, so relative error is amplified by P / Q. At shape 1E-6 and X = 0.5 the measured amplification is about 1.79E6.

## Minimal counterexample

~~~text
K_STATS_Gamma_Survival(0.5, 0.0001, 1)
reference: 5.598029295740171303421043E-05
baseline relative error: 1.68E-10
contract: 5E-15 relative
breach: about 33,679x
~~~

Equivalent Chi-square route:

~~~text
K_STATS_ChiSquare_Survival(1, 0.0002)
    -> Q(A = 0.0001, X = 0.5)
~~~

ScaleParam = 1 makes StandardX = X exactly. This is not the positive-ratio transform defect in #13.

## Mechanism

For an approximate lower probability P-hat = P + delta:

~~~text
Q-hat = 1 - P-hat = Q - delta

relative_error(Q) / relative_error(P) = P / Q
~~~

The committed real-VBA ladder measures a 1.00 predicted/observed ratio across all tested shapes. The complement fully accounts for the lower-region survival amplification.

## Dependency on #23

The direct-Q reconstruction must start from an accurate log P. #23 repairs the shared prefactor error beneath the lower series and must be validated first.

For #13, this issue blocks only the survival hard-underflow arm. Gamma/Chi-square density and cumulative dispatch do not call the lower-region Q reconstruction and may proceed after the post-#23 cutoff revalidation gate. #13 remains open until #24 is complete and all six forward surfaces are integrated and validated.

Do not close this issue by hiding the visible survival error while an inaccurate lower P remains. Do not relax Gamma_Survival.all.output or the corresponding Chi-square contract.

## Required implementation

Expose or add a lower-series log result assembled before final exponentiation, conceptually:

~~~text
LogP = LogPrefactor + Log(SeriesSum)
P    = Exp(LogP)
Q    = -PROB_Expm1(LogP)
~~~

Requirements:

- never compute Q as 1 - P in the lower branch;
- do not compute LogP by taking Log of an already rounded P near one;
- share the same converged series and prefactor arithmetic between P and Q;
- use the repaired PROB_Expm1 from #21;
- treat a materially positive LogP as an internal numerical failure;
- allow only a narrowly justified roundoff repair at the mathematical probability boundary;
- preserve the continued-fraction branch for X >= A + 1;
- preserve direct survival through the Gamma and Chi-square wrappers;
- reuse the lower-region log result for #13 hard-underflow recovery where appropriate.

## Preregistered evidence

Before changing source:

- [ ] freeze the fitting ladder and untouched shape-major holdout;
- [ ] record expected movement for P, Q, and complement consistency;
- [ ] include seam probes immediately below/at/above X = A + 1;
- [ ] identify representative rows for permanent main-grid promotion.

Permanent VBA regressions must include:

- the minimal Gamma call;
- the equivalent Chi-square call;
- a very small shape with large P/Q amplification;
- a moderate shape near the branch seam;
- direct kernel success and public error-code behavior.

## Acceptance

- [ ] the lower branch never forms 1 - P;
- [ ] minimal Gamma and Chi-square counterexamples satisfy frozen contracts;
- [ ] the full preregistered shape ladder and independent holdout pass;
- [ ] P and Q remain monotone and complementary within their explicit contracts;
- [ ] no discontinuity beyond contract at X = A + 1;
- [ ] #13 hard-underflow survival can reuse the stable construction without a separate formula;
- [ ] continued-fraction behavior is unchanged except for shared-prefactor movement measured under #26;
- [ ] full real-Excel regression, main grid, unrelated holdouts, and strict Accuracy Gate pass;
- [ ] no frozen threshold is relaxed.

### #26 — [Bug]: Revalidate the continued-fraction-Q survival breach after the shared prefactor fix

- State: open
- Labels: bug, P1, special-functions, t-family, continuous, accuracy
- URL: https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/issues/26

#### Body

## Description

Before the #23 source fix, the continued-fraction branch at Q(A, X = 2) showed a roughly flat 4.35E-14 to 5.68E-14 relative error across shapes 1E-6 through 0.5. That breached the 5E-15 Gamma survival contract by about 8.7x to 11.4x.

The branch does not form Q = 1 - P, so the observation is distinct from #24. However, it is no longer proven to be a separate continued-fraction algorithm defect.

## Why revalidation is mandatory

The measured path is:

~~~text
PROB_TryGammaContinuedFractionQ
  -> modified Lentz recurrence
  -> PROB_TryGammaPrefactor
  -> Result = Factor * h
~~~

The fix merged for #23 changes PROB_StirlingError beneath the shared PROB_TryGammaPrefactor. Therefore the old X = 2 error may be entirely or partly removed without changing the Lentz recurrence or its termination condition.

The previous body named Abs(Del - 1#) <= PROB_NUM_EPS as the leading hypothesis. That remains a possible residual cause, not an established root cause. Production source must not be changed on that hypothesis until the post-#23 decomposition is measured in real VBA.

## Baseline counterexample

~~~text
K_STATS_Gamma_Survival(2, 0.0001, 1)
kernel: Q(A = 0.0001, X = 2)
baseline observed: 4.8908149539870163E-06
reference:         4.8908149539872476E-06
baseline relative error: about 4.73E-14
contract: 5E-15 relative
~~~

Equivalent Chi-square route:

~~~text
K_STATS_ChiSquare_Survival(4, 0.0002)
    -> Q(A = 0.0001, X = 2)
~~~

Evidence baseline: benchmark/positive_ratio_subnormal/gamma_q_complement.csv.

## Decision gate after #23

First, re-export the exact same preregistered X = 2 ladder from the #23 implementation.

If every row now satisfies the frozen contract:

- [ ] record that the shared prefactor was the root cause;
- [ ] add representative permanent Gamma and Chi-square regressions;
- [ ] promote representative main-grid evidence;
- [ ] close this issue without changing the continued-fraction recurrence.

If a residual breach remains:

- [ ] decompose log/prefactor error, recurrence h error, multiplication error, and stopping error independently;
- [ ] vary iteration tolerance and budget only in the study harness;
- [ ] establish the minimal residual counterexample;
- [ ] preregister fitting, seam, and independent-holdout movement;
- [ ] implement the smallest kernel change supported by that decomposition;
- [ ] retain the old path as a study comparator.

Do not delete or merge this issue before the remeasurement; it is a release-blocking uncertainty until the branch is demonstrated clean.

## Regression scope

Validate:

- the full small-shape X = 2 ladder;
- points immediately around X = A + 1;
- moderate and large A continued-fraction paths;
- equivalent Gamma and Chi-square public calls;
- Poisson paths sharing incomplete-gamma kernels;
- P/Q complement consistency;
- non-convergence and iteration-budget behavior.

## Acceptance

- [ ] every preregistered fitting and untouched holdout row satisfies the existing survival contracts;
- [ ] the X = 2 minimal Gamma and equivalent Chi-square calls pass;
- [ ] the result identifies prefactor, recurrence, termination, or final arithmetic as the actual cause;
- [ ] permanent VBA and main-grid regressions cover the resolved path;
- [ ] no series-P, inverse, large-shape, Poisson, monotonicity, or error-code regression;
- [ ] full real-Excel regression and strict Accuracy Gate pass;
- [ ] Gamma_Survival.all.output remains 5E-15 relative.

### #28 — [CI]: Generate and fail-closed-check root README assurance metrics

- State: open
- Labels: documentation, testing, CI, release, P2, accuracy
- URL: https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/issues/28

#### Body

## Description

The root README assurance badges and “Assurance at a glance” table are hand-maintained and contradict the repository's authoritative evidence.

At the v1.0.0 planning baseline:

| claim | root README | current repository state |
| --- | ---: | ---: |
| Excel regression assertions | 835 | 902 in the latest verified exact-source Excel run; 909 expected after #23, not yet verified |
| measured accuracy contracts | 161 | 166 registry rows: 165 active and 1 characterization-only |
| observation rows | 1,905 | 2,088 rows in probability_accuracy_grid.csv |
| known silent-wrong paths | 0 | multiple open P1 release blockers |
| freshness | says regenerated/fail-closed | current main evidence is stale and Accuracy Gate run 166 correctly fails |

The README says every number is regenerated by CI and fails the build if it drifts. No root-README assurance generator currently enforces that statement.

## Required distinction

The public summary must not collapse these independent facts:

1. what the last real-Excel regression executed and whether it passed;
2. how many registry and observation rows exist;
3. whether observations are bound to the checked-out source;
4. how many contract verdicts are PASS, FAIL, PENDING, or characterization-only;
5. whether independent holdout observations are current and source-bound;
6. whether the release is certified despite known blockers outside scored regions.

A green historical summary is not current release evidence after source changes.

## Authoritative inputs

Create a deterministic offline renderer. At minimum it must read:

- a committed machine-readable Excel regression record containing source commit, total/passed/failed counts, Excel version/build/bitness, and execution time;
- accuracy_contracts.csv for active and characterization counts;
- probability_accuracy_grid.csv plus observation_manifest.json for row count and freshness;
- the #22 completeness result for claimed/exempt/missing row counts;
- holdout summary plus #29 source provenance for holdout count and freshness;
- accuracy_summary.md or its machine-readable verdict data for PASS/FAIL/PENDING;
- a committed release-readiness state if the README makes a release-certification claim.

Do not scrape live GitHub issues at render time. The build must remain reproducible offline. Either maintain known release blockers in a committed machine-readable readiness registry, or replace the unsupported “0 known silent-wrong paths” claim with narrower language the committed artifacts can prove.

## Renderer behavior

- [ ] update only clearly delimited generated regions in root README.md;
- [ ] generate numeric badges and the assurance table from the same data model;
- [ ] fail on missing, malformed, contradictory, or stale required inputs;
- [ ] render stale Excel/main-grid/holdout evidence as visibly stale or blocking;
- [ ] never fall back to the previous number or zero;
- [ ] use exact formatting consistently, including 2,088 rather than mixed 1 905 / 1,905 styles;
- [ ] state the authority for every metric in documentation;
- [ ] add unit tests for current, stale source, missing input, malformed input, contradictory totals, and manual README drift.

## CI integration

- [ ] invoke the renderer from refresh_evidence.py after all authoritative summaries/provenance are generated;
- [ ] run the renderer in the hosted Accuracy Gate and fail on any README diff;
- [ ] ensure the check also runs when README, renderer, Excel result record, contracts, grid, manifests, holdout evidence, or readiness registry changes;
- [ ] keep the two badge values and table values impossible to diverge;
- [ ] verify a deliberate manual edit fails CI.

## Final v1.0.0 content

The final release block must display only values verified on the release source. Expected but unverified counts such as 909 must never be rendered as PASS.

At release, the block should show:

- exact Excel assertions and environment;
- registry status counts;
- total/main-grid/claimed/exempt/missing observation counts;
- source freshness;
- main contract verdict tally;
- source-bound independent holdout tally;
- explicit release certification state.

## Acceptance

- [ ] 835, 161, and 1 905 are removed as stale literals;
- [ ] the README does not claim zero known defects while P1 blockers remain;
- [ ] all displayed metrics are generated from named authorities;
- [ ] stale evidence produces a blocking visible state and CI failure;
- [ ] final v1.0.0 Excel and accuracy evidence regenerate the block cleanly;
- [ ] changing any authority without regenerating README fails CI.

### #29 — [CI]: Bind independent holdout observations to source provenance

- State: open
- Labels: testing, CI, release, P2, accuracy
- URL: https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/issues/29

#### Body

## Description

The independent holdout is rerun by the hosted Accuracy Gate, but its observations are not bound to the VBA source that produced them.

The main grid has benchmark/observation_manifest.json, which binds the observation file, registry, schema, exporter, production modules, and test module by hash. The holdout has no equivalent source binding. As a result, analyze_holdout.py can recompute a green summary against the current contract registry while reading observations exported from older VBA source.

This is a provenance gap, not a claim that the 80 current contract verdicts across 559 holdout-grid rows are numerically wrong. The existing badge is numerically correct but ambiguous about whether 80 counts rows or verdicts.

## Measured repository state

At the v1.0.0 planning baseline:

- current repository main: 0dd748884599d4d0da815cb53eeceb13efd51f05
- latest numerical source baseline: bde92dd7037e4fde05e620745a1c54b0cbc3a261
- holdout_grid.csv last changed at 4553afa6310e07858a97e412722554822a92b674
- holdout_summary.md last changed at the same older source state
- holdout_grid.csv contains 559 observation rows
- the summary reports 80 contract verdicts: 80 PASS and 0 FAIL
- the current Accuracy Gate reruns the analyzer and diffs the summary, but cannot prove those observations came from current source

Issue #9 explicitly identified this remaining gap when the analyzer and summary checks were added to CI; no dedicated follow-up issue was created.

## Required design

Extend the existing provenance architecture rather than inventing an unrelated one.

- [ ] Bind holdout_grid.csv to the exact production modules and holdout exporter that produced it.
- [ ] Bind the contract registry and holdout schema used to interpret the observations.
- [ ] Record Excel version, build, bitness, export timestamp, and source commit.
- [ ] Verify normalized content hashes with the same line-ending policy as the main manifest.
- [ ] Make missing, malformed, stale, added, removed, or changed source bindings fail closed.
- [ ] Make analyze_holdout.py or its caller refuse to publish a release verdict from stale observations.
- [ ] Add unit tests for clean, changed-source, changed-grid, changed-registry, changed-schema, missing-manifest, and line-ending cases.
- [ ] Run the verification in refresh_evidence.py and the hosted Accuracy Gate before the holdout analyzer.
- [ ] Document the relationship between main-grid and holdout provenance in benchmark/PROVENANCE.md.

A single extended manifest or a dedicated holdout manifest is acceptable if it preserves the same fail-closed semantics and keeps the authority unambiguous.

## Sequencing requirement

Split mechanism from the first truthful binding:

1. Phase 0 lands the Excel-free mechanism: manifest writer, verifier, fail-closed CI wiring, PROVENANCE.md, and clean/changed-source/changed-grid/changed-registry/changed-schema/missing-manifest/line-ending tests.
2. The existing 559 rows must never be bound to current source: they were exported at 4553afa. Until a fresh export exists, the new verifier must report the holdout as stale or unbound and the gate remains red.
3. The #23 Excel evidence wave produces the first truthful current-source binding while exporting the current global holdout. #22's new inverse neighbour/bridge evidence is deliberately not a prerequisite and is added in a later release-candidate wave. The manifest always records the exact row count and content hash of the file actually exported; neither 559 nor any later count is assumed permanent.
4. #29 closes only after that manifest verifies, the analyzer passes, and a deliberate source/grid mismatch fails CI.
5. Every later numerical phase refreshes the binding at export time; #31 owns the final refresh on the exact release source.

## v1.0.0 evidence requirement

After the last numerical source change:

1. import the exact tracked VBA modules into fresh Excel;
2. export the independent holdout observations;
3. write the holdout source binding immediately;
4. rerun the holdout analyzer against the frozen registry;
5. commit observations, provenance, and summary together;
6. verify CI rejects a deliberate source or grid mismatch.

## Acceptance

- [x] the Excel-free writer, verifier, documentation, CI wiring, and negative tests landed in 2dcfe04 and 4be8cd0 before the next holdout export;
- [x] every Phase 0 provenance commit through 0dd7488 preserves the frozen seven-line/two-file expected-red signature while its verifier fixtures pass;
- [ ] a fresh current-source holdout export has a valid source binding and passes the analyzer;
- [x] stale or missing holdout provenance cannot publish a current release verdict; public README display of the historical 80-verdict result is owned by #28;
- [ ] a source change without a new holdout export fails CI;
- [ ] Phase 1 writes the main-grid and holdout bindings atomically; no commit presents fresh main evidence with an absent or stale holdout binding;
- [ ] #31 requires the final v1.0.0 holdout to be rebound to the exact release source and records Excel 16.0 build 20131, 64-bit, or the replacement environment;
- [ ] no contract threshold is relaxed to close this issue.

### #30 — [CI]: Upgrade GitHub Actions to supported Node 24 majors

- State: open
- Labels: testing, CI, release, P3
- URL: https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/issues/30

#### Body

## Description

The v1.0.0 workflows still use GitHub Actions majors that target the deprecated Node 20 runtime. Accuracy Gate run 166 completed by forcing those actions onto Node 24, but emitted deprecation warnings for actions/checkout@v4 and actions/setup-python@v5.

A first stable release should not depend on GitHub's compatibility override when supported Node 24 action majors are available.

## Current workflow references

- actions/checkout@v4 in Accuracy Gate and Excel VBA Regression
- actions/setup-python@v5 in Accuracy Gate
- actions/upload-artifact@v4 in Excel VBA Regression
- actions/github-script@v7 in Label Taxonomy Sync

Latest upstream releases verified on 2026-08-28:

- actions/checkout v7.0.1
- actions/setup-python v7.0.0
- actions/upload-artifact v7.0.1
- actions/github-script v9.0.0

Use the supported major tags after reviewing their breaking-change notes; do not copy version numbers blindly into only one workflow.

## Required work

- [ ] Upgrade every occurrence consistently across .github/workflows.
- [ ] Preserve current permissions, path filters, concurrency, runner labels, artifact retention, and fail-closed semantics.
- [ ] Confirm the Accuracy Gate runs on Python 3.12 exactly as before.
- [ ] Confirm Excel result artifacts still upload on success and failure.
- [ ] Dispatch Label Taxonomy Sync and verify no labels are unintentionally pruned or changed.
- [ ] Verify the resulting workflow logs contain no Node 20 deprecation warning.
- [ ] Record the upstream major-version review in the commit message or release evidence.

## Acceptance

- [ ] all three workflows parse and run;
- [ ] hosted Accuracy Gate passes once numerical evidence is current;
- [ ] Excel VBA Regression reaches the self-hosted runner and publishes its result artifact;
- [ ] Label Taxonomy Sync reports the canonical label set unchanged;
- [ ] no deprecated Node 20 action runtime remains.

### #31 — [Release]: Certify and publish v1.0.0

- State: open
- Labels: documentation, testing, CI, release, P1, accuracy
- URL: https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/issues/31

#### Body

## Purpose

Own the final repository-readiness, certification, tagging, and publication work for the first stable release.

This issue closes only after every v1.0.0 release blocker is complete and the exact tagged source has reproducible Excel and hosted-gate evidence.

## Current readiness baseline

Repository-readiness state measured at main 0dd748884599d4d0da815cb53eeceb13efd51f05; latest numerical source baseline remains bde92dd7037e4fde05e620745a1c54b0cbc3a261:

- no Git tag exists;
- no GitHub Release exists;
- the repository has LICENSE, CONTRIBUTING.md, CODE_OF_CONDUCT.md, SECURITY.md, issue templates, PR template, source modules, consolidated tests, example workbook, and numerical-evidence documentation;
- CHANGELOG.md does not exist;
- SECURITY.md incorrectly calls v1.0.0 the latest tagged stable version before the tag exists;
- README assurance metrics are stale and hand-maintained (#28);
- Accuracy Gate runs 166 through 171 fail only with the frozen two-file stale-evidence signature; Phase 0 introduced no additional top-level failure;
- Excel VBA Regression run 95 was cancelled because the self-hosted runner was unavailable; it produced no evidence;
- main is intentionally the only remote branch and is currently unprotected;
- the project has 112 worksheet-facing functions, 166 registry rows, 2,088 total grid observations, 1,732 main-grid rows, a 36-row unclaimed audit baseline at bde92dd whose current count must thereafter be generated, and 559 independent-holdout rows producing 80 contract verdicts; the latest verified Excel result is 902/902, while the #23 tests make 909 the next expected count, not yet a verified result.

No release may be presented as stable while those states remain true.

## Blocking issue graph

Numerical correctness and representability:

- [ ] #23 — validate and close the shared off-grid Stirling prefactor fix
- [ ] #24 — construct lower-region Q directly from log P
- [ ] #13 — wire the measured positive-ratio/subnormal dispatch
- [ ] #14 — reconstruct representable inverse quantiles at final scale
- [ ] #26 — remeasure the continued-fraction Q path after #23 and fix only if residual error remains
- [ ] #11 — close the parent only after #13 and #14 are complete

Assurance and repository readiness:

- [ ] #29 — the Excel-free mechanism is complete; create the first truthful main-grid/holdout binding pair atomically and close it during the #23 holdout export
- [ ] #17 — make the current grid generator non-destructive by default and fail hard on duplicates, unexpected reference movement, observation overwrite, or implicit deletion
- [ ] #22 — Phase 0 coverage guard is complete and the df = 1E6/1E7/1E8 generator is accepted: quadrature and gmpy2/MPFR converge 69/69, while Rmpfr agrees on its 46 feasible points and cannot execute df = 1E8 because of its shape ceiling. Complete the remaining classifier/median/bridge evidence after the last inverse-affecting change, then atomically land the selected contracts, domain corrections, F cleanup, strict checker, and delete the temporary fingerprint.
- [x] #20 — permanent exact-binary64 Stirling boundary regression
- [ ] #28 — generated, current, fail-closed root README assurance metrics
- [ ] #30 — supported Node 24 GitHub Action majors

#32 is explicitly post-v1.0.0 under milestone v1.01. It owns full origin reconstruction, reference/observation separation, and byte-stable regeneration and is not a release blocker.

## Scope decision: do not defer #13/#14 from v1.0.0

A documentation-only descoping option was considered: register the positive-subnormal and hard-underflow regimes in numerical_limitations.csv and move #13/#14 to v1.1.0. It is rejected.

#11 contains accepted-domain calls that return a plausible silent wrong value, including 0 where the representable result is about 0.912, without raising an error. This is the project's highest-severity defect class. Shipping the first stable release while merely documenting that behavior would contradict the public measured-accuracy claim and weaken release credibility.

The only defensible fallback would be a source change that explicitly rejects the affected inputs with the documented worksheet error and narrows the supported domain, with matching tests, contracts, and documentation. That is not the selected v1.0.0 scope. The release remains blocked until #13 and #14 correct the accepted-domain behavior and #11 closes.

## Excel execution contingency

The existing self-hosted runner is the primary route, but it is a named schedule risk rather than an implicit dependency.

1. Restore the current runner first.
2. If it remains unavailable, provision a replacement Windows/Excel 64-bit self-hosted runner using the same exact-source import and artifact procedure.
3. A controlled manual fresh-workbook Excel run may keep an individual numerical phase moving only when it records source hashes, Excel version/build/bitness, assertion totals, exported observations, and provenance in the same machine-readable format.
4. Manual evidence does not waive the final requirement for a retained self-hosted workflow artifact on the exact release commit.
5. If neither controlled Excel route is available, stop after preregistration; do not accumulate further value-changing numerical commits without real-Excel evidence.

## Repository preparation

- [ ] Add CHANGELOG.md with an Unreleased section and a complete v1.0.0 entry.
- [ ] Treat tag v1.0.0 and its GitHub Release as the canonical version authority; do not introduce a second version file unless production code needs it.
- [ ] Correct SECURITY.md immediately to describe main as unreleased and no stable tag as published; switch it to supported v1.0.0 only in the release commit/tag.
- [ ] Audit README, benchmark documentation, issue templates, wiki links, example-workbook link, license references, and installation instructions against the release contents.
- [ ] Decide and document release artifacts: GitHub source archives are mandatory; any attached demo workbook or source bundle must be generated from or bit-identical to the tagged tree and accompanied by SHA-256.
- [ ] Remove or explicitly accept every CI warning; #30 owns the current Node runtime warnings.
- [ ] Confirm no untracked scratch, local workbook, credentials, generated cache, or editor files are part of the tag.
- [ ] Record why main remains unprotected for this solo-maintainer, single-branch pre-release workflow, or enable an appropriate ruleset after the direct-main implementation phase.

## Final certification sequence

Run this only after the final numerical source change.

1. Check out the intended release commit in a clean working tree.
2. Import the exact six tracked production modules and consolidated test module into a fresh .xlsm.
3. Compile the VBA project in Excel 16.0 64-bit.
4. Run Test_STATS_PROBDIST_RunAll and require every assertion to pass.
5. Export the full accuracy grid and every affected fitting study from that exact workbook/source.
6. Export all preregistered independent holdouts without changing their frozen design.
7. Write main-grid and holdout provenance before generating summaries.
8. Run refresh_evidence.py and every hosted Accuracy Gate check in release order.
9. Require zero FAIL, zero PENDING, zero unclaimed main-grid rows, no remaining transitional coverage-debt fingerprint, and no unexplained movement outside preregistered manifests.
10. Rerun the self-hosted Excel workflow from the exact release commit and retain its artifact.
11. Perform a clean import smoke test of the public modules and example workbook.
12. Freeze changelog and release notes, create an annotated v1.0.0 tag, push it, and create the GitHub Release.
13. Verify the tag, release assets, hashes, README badges, SECURITY supported-version table, and public links from GitHub after publication.

## Release notes minimum content

- first stable release scope and installation;
- supported Excel/VBA environment and host-independence statement;
- public API/family summary;
- numerical architecture and direct-tail behavior;
- regression assertion count and exact Excel provenance;
- active contract, grid, and independent-holdout counts;
- material numerical fixes, including #11/#13/#14/#23/#24/#26;
- documented representability and supported-domain limitations;
- verification commands and source commit/tag;
- upgrade notes for users of untagged main snapshots.

## Go / no-go rule

GO only when:

- every P1 in milestone v1.0.0 is closed with evidence;
- no accepted-domain silent-wrong path is deferred through documentation-only limitation wording;
- #22, #29, and #28 make public assurance complete and fail closed;
- final Excel, main grid, holdouts, and hosted workflows are green on the same source;
- documentation no longer describes an unreleased tag as stable;
- the release commit is clean, reproducible, and tagged exactly v1.0.0.

Otherwise the decision is NO-GO, regardless of whether an older grid or individual workflow is green.
