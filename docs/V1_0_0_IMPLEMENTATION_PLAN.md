# v1.0.0 implementation and release-readiness plan

Status date: 2026-11-11
Repository: danielep71/VBA-PROBABILITY-DISTRIBUTIONS
Base: `release/1.0.0` at `a9812c7`
Milestone: v1.0.0 (due 2026-12-18)

**Current evidence state.** The Excel run, grid bindings and export certification, coverage, contract and holdout verdicts, and registered release blockers are generated from committed records in the [README assurance block](../README.md#assurance-at-a-glance) and checked by the Accuracy Gate on every change. They are not restated here.

This plan covers only the work that remains. The milestone's
[GitHub issues](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/milestone/1)
and [`benchmark/release_readiness.json`](../benchmark/release_readiness.json)
are authoritative; where this page disagrees with them, they win. Earlier
versions of this page, including the 2026-09-30 audit, the phase-by-phase
plan and the issue register copied on 2026-08-29, are in Git history (last
full version at `a9812c7`).

## Status summary

v1.0.0 is **not release-ready** (NO-GO).

- **Registered release blockers.** `release_readiness.json` (reviewed
  2026-11-11, `certified: false`) lists 8 open blockers: #11, #13, #14, #22,
  #35, #163, #173 and #213. #23, #24, #26, #29, #34, #47, #72 and #307 are
  recorded there as resolved, with their evidence.
- **Milestone.** The P3 issues of v1.0.0 moved to v1.0.2 on 2026-11-11; the
  [milestone](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/milestone/1)
  lists what remains.
- **Balanced Beta.** #163 and #173 were added to the registry on 2026-11-07
  as open silent-wrong blockers, and the README renderer now requires both
  entries, as it does #34 and #35.
- **Retained Excel evidence.** Candidate `d6d952a`, 5,987/5,987 regression
  assertions, Excel 16.0 build 20430, 64-bit
  ([`excel_regression_record.json`](../benchmark/excel_regression_record.json);
  expected count in
  [`.github/excel-evidence-policy.json`](../.github/excel-evidence-policy.json)).
  It exported and certifies the 2,139-row main grid and the 659-row holdout.
  It does not certify the release commit.
- **Coverage debt.** 27 main-grid rows have no disposition:
  `F_InverseCumulative` 18 and `ChiSquare_InverseCumulative` 9, both
  `envelope_domain` ([`coverage_summary.md`](../benchmark/coverage_summary.md)).
- **No tag or GitHub Release exists.** `VERSION` is `0.0.0`.

## Remaining work

Every item that changes a `.bas` file needs a committed preregistration before
the source edit, then an Excel wave: regression plus fresh exports from the
same candidate, applied in one evidence commit (see
[Process](#process-that-remains-valid)).

### A. Gamma-family defects: #11, #13, #14

These issues share the incomplete-gamma kernels. #11 is the parent and closes
only after #13 and #14. #23, #24 and #26, which came first in this order, are
resolved (see [Completed](#completed-since-the-original-plan)); the remaining
order is #13 → #14 → #11.

- **#13: positive standardized argument.** `PROB_MIN_NORMAL`, the 48-bit and
  40-bit cutoffs and `PROB_ClassifyPositiveRatio` exist but have no
  production callers.
  - The cutoffs are pre-#23 candidates. Retain or rederive them from the
    post-#23 rerun before wiring.
  - Then dispatch at the six Gamma/Chi-square density, cumulative and
    survival call sites.
  - Survival hard-underflow recovery needed #24, now resolved.
- **#14: log-domain inverse reconstruction.**
  - Add a log-quantile kernel, conceptually `PROB_TryGammaInvPLog`.
  - Reconstruct the final scale with `Log(ScaleParam)` for Gamma and
    `Log(2#)` for Chi-square, exponentiating once.
  - Add quantile-relative and tail-residual contracts.
  - It needs the corrected forward path from #13 (#24 is resolved).
- **#11: parent.** Closes on forward/inverse round-trip evidence for its
  counterexamples on the final source.

Scope decision, still in force: #13 and #14 are not deferred by documenting
the regime as a limitation. #11 includes accepted-domain calls that return a
plausible wrong value, for example 0 where about 0.912 is correct. The only
alternative is a source change that rejects the regime with a documented
worksheet error.

### B. F inverse and coverage debt: #35 and #22

- **#35: unbalanced F inverse.**
  - `PROB_TryBetaInvRegularized` forms the unsolved pair member as `1# - U`.
    When that member is tiny, its relative error dominates F's
    `BetaX / BetaY`.
  - The isolation study (`e68c7ad`) recorded `NONE` against its
    preregistered criteria. The fix comparison was preregistered at `9fba175`
    (`benchmark/f_inverse_large_df_study/FIX_COMPARISON_PREREGISTRATION.md`),
    but its faithful binary64 model and candidate results do not exist yet.
  - Any fix also changes Beta inverse and Student-t's small-df beta-inverse
    path, so the Excel wave must cover all three.
  - Alternatively, a rigorous characterization can show that the 2E-10 F
    threshold is the achievable accuracy.
- **#22: zero coverage debt.** Its remaining scope, checked against the
  current grids and registry:
  - give the 18 F and 9 Chi-square `envelope_domain` inverse rows real
    contracts. The Student-t rows of the original 36-row debt were settled
    by #34;
  - remove the false "enforced by `PROB_F_ValidateEnvelope`" wording from the
    four F `validated` contracts in `accuracy_contracts.csv`. The runtime cap
    is `PROB_F_MAX_DF = 1E10`;
  - reclassify the eight main-grid `validated` F inverse rows with a df above
    1E5: (1E6, 3) and (2.5, 1E6), at four probabilities each;
  - consolidate the three duplicate (1E6, 3) identities at p = 0.5, 0.9 and
    0.99;
  - reclassify the eight legacy holdout rows at (1.5, 2E5) and (5E5, 4);
  - add the inverse-regime classifier and the threshold-zero tooling;
  - preregister fitting and untouched holdout arms for each threshold.

  The F work waits for #35, and the final measurement runs on the
  release-candidate source after #14. The closing commit deletes
  `benchmark/coverage_debt_v1_0_0.json`, and strict mode must then report
  zero missing rows.

### C. Student-t inverse near |q| = 8: #213

The #210 seam study (run 42281403598) measured 1.29E-12 relative error at
df = 1E7, for t = -8 and t = -8 + 1 ulp. That is inside the domain of
`StudentT_InverseCumulative.envelope_domain.quantile_rel` (5E-15,
`|q| <= 8`). The continued fraction serves `|t| > 8` and overestimates the
survival there, and the polish does not run above 8.

Choose one:

- fix the inverse by bracketing at 8 and polishing, with its own
  preregistration and Excel wave;
- narrow the contract domain and register the sliver as a characterized
  limitation.

Either way, add a regression at df = 1E7 for t = -8 and t = -8 + 1 ulp.

### D. Balanced Beta contracts: #163, #173

Decision of 2026-11-05: for v1.0.0, recalibrate the balanced
`Beta_Survival` (currently 5E-15) and `Beta_Cumulative` (currently 2E-14)
contracts to what the route delivers. The route repair moves to #175
(v1.0.1).

- The proposed threshold is about 2E-13, for ratio ≥ 0.1 with both shapes in
  [0.5, 100), from the `benchmark/beta_balanced_survival` model (PR #174).
- Before adoption, a preregistered Excel wave must confirm it:
  - at the modelled worst fitting points (0.99, 14, 19), (0.7, 19, 14) and
    (0.41, 50, 100);
  - on a disjoint holdout frozen before it is observed.
- Resolve the shape-100 point explicitly: [0.5, 100) excludes it. Either
  validate an inclusive bound or keep the point as an outside-domain control.
- The threshold change is a reviewed recalibration with recorded evidence,
  never a silent relaxation. No VBA change is planned, but the wave needs the
  export tooling.

### E. Documentation and test records: #154, #155, #4

- **#154.** Regenerate `benchmark/cross_check_scipy.md` from the current
  grids (the holdout is now 659 rows). Record the environment and the grid
  SHA-256 values, and reconcile the README coverage statement. The mapping
  fix (PR #160) is done.
- **#155.** Make `benchmark/README.md` describe the survival-reference route
  generator by generator. The references themselves were already rebased in
  #159.
- **#4.** Make the Student-t validator header in
  `src/M_STATS_PROBDIST_TFAMILY.bas` distinguish:
  - the df ≥ 1E5, |t| ≤ 8 large-df route;
  - the retained far-tail limitation;
  - the 5E-15 inverse contract now adopted.

  The Chi-square header is already fixed. A `.bas` comment change still
  alters the source binding, so land it inside another VBA Excel wave.

### F. Release readiness and certification: #28, #31

- **#28.**
  - Keep `release_readiness.json` in step with GitHub. The #34 and #72
    entries are now resolved, by PR #205 and PR #217.
  - #163 and #173 are registered as open silent-wrong blockers
    (2026-11-07).
  - After the last numerical change, regenerate the README assurance block.
    The renderer (`benchmark/render_readme_assurance.py`, PRs #60/#62) and
    its Accuracy Gate drift check already exist.
- **#31.** Final certification and publication:
  - complete the v1.0.0 `CHANGELOG.md` entry and reconcile `VERSION`;
  - freeze the release notes and artifact policy;
  - certify the exact release commit in Excel, which the release workflow
    enforces since #76 (PR #133);
  - tag `v1.0.0`, publish the GitHub Release and verify it after
    publication, following [RELEASING.md](../RELEASING.md).

## Completed since the original plan

| Issue | Resolution | Merged |
| --- | --- | --- |
| #17 grid-regeneration safeguards | commit `d2e1592` (before pull requests were mandatory) | closed 2026-09-30 |
| #30 Node 24 action majors | PR #36 | 2026-09-06 |
| #28 renderer (issue still open, see F) | PRs #60, #62 | 2026-10-01, 2026-10-09 |
| #47 export digests match committed bytes | PR #67 | 2026-10-12 |
| #76 exact-tag Excel certification | PR #133 | 2026-11-03 |
| #181 unattended Excel grid export; #29 holdout certification | PR #182 (`ci/Export-ExcelGrids.ps1`, `benchmark/apply_excel_export.py`), integrated by PR #183 | 2026-11-06 |
| #72 normal upper-tail kernel; #159 survival references at the 25-digit floor; #170 `holdout_older` re-export | PR #81, integrated by PR #185 | 2026-11-06 |
| #34 Student-t large-df route, inverse polish and 5E-15 `envelope_domain` contracts | PRs #189, #198, #200, #205, integrated by PRs #199, #211 | 2026-11-06 |
| #210 Student-t seam study (milestone v1.0.1) | PR #214, integrated by PR #215; found #213 | 2026-11-06 |
| Stale #72 blocker entry (#28) | PR #217 | 2026-11-06 |
| #23 off-grid Stirling prefactor evidence | fitting, holdout and promoted rows (`benchmark/gamma_series_small_shape/COMPLETION_RESULT.md`) | closed 2026-11-07 |
| #26 continued-fraction Q survival; public Gamma P/Q through the scaled double-double kernel | PRs #242, #289, #290 | closed 2026-11-09 |
| #24 direct lower-region Q | PR #304, integrated by PR #305 | closed 2026-11-09 |
| #298 per-procedure timing; #300 study batteries apart from the regression | PRs #308, #310 | closed 2026-11-10 |
| #307 Gamma kernel cost: exact LogGamma reuse; routing studies retained, no routing rule | PRs #311–#313, #317–#319, integrated by PRs #315, #320 | closed 2026-11-11 by maintainer decision |

Other closed v1.0.0 milestone issues, mostly CI, repository and evidence
hardening, are listed in the milestone: #38, #40, #42, #49, #52, #77–#80,
#86–#90, #99–#101, #103–#109, #111, #112, #129, #137, #142, #148, #164, #167,
#171, #187, #201–#203 and #206–#208.

## Recorded decisions

- **Chi-square df = 1E8 reference substitution (#22).** This is item 6 of the
  original plan's track A2, cited by `benchmark/chisq_reference_study`.
  - It was accepted on 2026-08-30 from `e68cc82`, which followed the
    69-point freeze in `4deeef4`.
  - Quadrature and the converging gmpy2/MPFR series/continued-fraction route
    cover 69/69 points, with minimum agreement of 109.87 and 115.11
    significant digits.
  - Rmpfr agrees to at least 115.96 digits on its 46 feasible points. Its
    `mpfr_gamma_inc` aborts above shape about 4.5E7, and df = 1E8 needs
    shape 5E7.
  - For df = 1E8, gmpy2/MPFR is the arithmetic route and quadrature the
    algorithm-independent cross-check.
  - This settles oracle feasibility only. It sets no threshold and does not
    authorize inspecting the holdout.
- **Deferred to v1.0.1.** #7 (contract-margin audit), #32 (byte-stable grid
  reconstruction), #175 (balanced Beta route repair). None blocks v1.0.0.

## Process that remains valid

- **Preregister before VBA.** Commit predictions, the fitting arm and an
  untouched holdout before the source edit. Never relax a frozen contract to
  make a defect pass. Close an issue only after its fitting, holdout,
  permanent VBA regression, full Excel run and strict gate pass. See
  [CONTRIBUTING.md](../CONTRIBUTING.md#validation-and-evidence).
- **Excel wave.** Dispatch the Excel VBA Regression workflow with the grid
  exports, then apply the artifact with `benchmark/apply_excel_export.py`. The
  grid, its manifest and its export record change together in one evidence
  commit. See
  [EXCEL_CERTIFICATION.md](EXCEL_CERTIFICATION.md#automated-fresh-export).
- **Branches.** Task pull requests target `release/1.0.0`. An integration PR
  merges it into `main` with `Refs #`, and a sync PR merges `main` back.
  Issues closed through the release branch are closed by hand with a
  Resolution section. See
  [RELEASING.md](../RELEASING.md#merge-convention).

## Go / no-go checklist

GO requires every item. Items marked done describe lasting state; the
others must hold on the exact release commit.

- [ ] Every blocker in `release_readiness.json` is resolved with evidence
  (11 open).
- [ ] Every v1.0.0 milestone issue is closed or explicitly moved (16 open).
- [ ] No accepted-domain wrong result is left documented as a limitation
  instead of fixed (#11, #13, #14, #163, #173, #213).
- [x] Exact-SHA Excel certification, unattended grid export and exact-tag
  certification exist (#38, #181, #76).
- [x] Export records validate against committed grid bytes, and the holdout
  export is certified (#47, #29). Repeat on the release commit.
- [ ] A full real-Excel regression passes on the release commit, and its
  workflow artifact is retained. The current record is `9cbe5d9`.
- [ ] The main grid and holdout are exported from the release source.
- [ ] Zero FAIL and zero PENDING, on the main grid and the holdout.
- [ ] Zero missing main-grid dispositions, and the #22 debt fingerprint is
  deleted (27 rows remain).
- [ ] No contract is relaxed. Any recalibration (#163/#173) has recorded,
  reviewed evidence.
- [x] The README assurance block is generated and checked by the Accuracy
  Gate (#28 renderer). It must be regenerated on the final evidence.
- [x] Rulesets are active:
  - `Protect main`: pull requests only, merge commits only, conversation
    resolution required, and up-to-date `Repository integrity` and
    `Strict accuracy gate` checks;
  - `Protect release branches`;
  - `Prevent release branch deletion`;
  - `Protect release tags`;
  - `Protect traffic history`.

  See [SETTINGS.md](repo_governance/SETTINGS.md).
- [x] `SECURITY.md` names no stable release before one exists. Recheck it
  against the tag.
- [ ] The `CHANGELOG.md` v1.0.0 entry and release notes are complete, and
  `VERSION` is reconciled.
- [x] Workflows run on supported Node 24 action majors (#30).
- [ ] A clean import and an example-workbook smoke test pass.
- [ ] The annotated `v1.0.0` tag and its GitHub Release are published and
  verified.

Any unchecked item means NO-GO. Resolve or explicitly disposition every
actionable review finding; the ruleset also requires resolved conversations
before merge.
