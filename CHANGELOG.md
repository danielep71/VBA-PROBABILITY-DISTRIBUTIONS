<div align="center">

# 📜 Changelog

### Release history for tail-aware numerical probability functions

[![Format](https://img.shields.io/badge/Format-Keep_a_Changelog-0969da?style=flat-square)](https://keepachangelog.com/en/1.1.0/)
[![Versioning](https://img.shields.io/badge/Versioning-SemVer-6f42c1?style=flat-square)](https://semver.org/spec/v2.0.0.html)
[![Dates](https://img.shields.io/badge/Dates-YYYY--MM--DD-217346?style=flat-square)](#date-and-version-rules)
[![Staging](https://img.shields.io/badge/Staging-Unreleased_first-d97706?style=flat-square)](#unreleased)
[![Contributing](https://img.shields.io/badge/Changes-Contribution_guide-2ea44f?style=flat-square)](CONTRIBUTING.md)

<br>

**User-visible history · Explicit compatibility · Reproducible evidence · Immutable releases**

</div>

---

All notable changes to **VBA Probability Distributions** are documented here.

This changelog follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/)
and [Semantic Versioning](https://semver.org/spec/v2.0.0.html). It records
released behavior and material unreleased changes; it is not a commit log, issue
tracker, or substitute for release evidence.

Versioning covers public worksheet functions, distribution parameterization, numerical and tail behavior, accuracy claims, convergence and error contracts, diagnostics, and supported Excel environments.

---

## 🧭 Maintenance policy

- Add material changes under **Unreleased** in the same pull request as the
  behavior or documentation they describe.
- Write from the user's perspective: describe the observable result, contract,
  compatibility impact, and migration need.
- Link the owning issue or pull request when it contains useful engineering
  detail.
- Keep entries concise; do not duplicate implementation notes already preserved
  in source, issues, or technical documentation.
- Record only validation actually performed. State skipped environments and
  known limitations plainly.
- Move Unreleased entries into a dated version section during release.
- Do not edit a published release entry except to correct a demonstrable factual
  or link error; annotate material corrections instead of rewriting history.
- Never claim that a tag, binary, workbook, hash, test run, or environment was
  certified unless the evidence binds it to the released source.

See [CONTRIBUTING.md](CONTRIBUTING.md) for change and evidence requirements and
[SECURITY.md](SECURITY.md) for private vulnerability reporting.

<a id="date-and-version-rules"></a>

### Date and version rules

| Rule | Standard |
|---|---|
| Version | `MAJOR.MINOR.PATCH`, without the leading `v` in headings |
| Release heading | `## [X.Y.Z] - YYYY-MM-DD` |
| Date | Gregorian calendar date in ISO `YYYY-MM-DD` format |
| Ordering | Unreleased first; released versions newest to oldest |
| Comparison | Unreleased → latest tag; each release → preceding tag |
| Patch | Backward-compatible correction or hardening |
| Minor | Backward-compatible capability |
| Major | Incompatible public-contract change |
| Pre-release | State maturity and compatibility boundaries explicitly |

A repository may remain below `1.0.0` while its supported surface is still
forming. Pre-release status does not excuse undocumented breaking changes.

<details>
<summary><strong>Entry categories</strong></summary>

<br>

| Category | Use for |
|---|---|
| **Added** | New supported capabilities, APIs, files, or tests |
| **Changed** | Changes to existing behavior, contracts, tooling, or documentation |
| **Deprecated** | Supported behavior scheduled for removal |
| **Removed** | Removed capabilities or compatibility |
| **Fixed** | Corrected defects |
| **Security** | Safely disclosed security corrections |
| **Documentation** | Material documentation-only changes |
| **Validation** | Evidence actually produced |
| **Compatibility** | Upgrade or migration effects |
| **Known limitations** | Deliberate, unresolved boundaries |

Use only the categories needed by a release.

</details>

---

<a id="unreleased"></a>

## [Unreleased]

### Added

- Added a release workflow (#31). Pushing an annotated `vX.Y.Z` tag verifies the release: the tag must equal `VERSION` and be on `main`, release semantics must validate, the retained Excel record must validate at the tagged commit, the README assurance block must be current, and `benchmark/release_readiness.json` must declare the version certified. The workflow then builds a reproducible source bundle (modules with CRLF line endings, licence, documentation and a member manifest) with `SHA256SUMS`, attests both with GitHub build provenance, and drafts a GitHub Release with the changelog section and an asset table for the maintainer to publish. A manual run rehearses the checks and the build without attesting or releasing. The bundle builder's self-test runs in Static checks.
- Extended weekly Dependabot updates to the benchmark and lint Python requirements, with maintainer assignment and an initial priority label. mpmath 1.4 is held back: no SymPy release accepts it yet, and a change of the reference library is a deliberate evidence change, not a routine update.

- Added exact-SHA Excel certification records binding the candidate commit, imported VBA source hashes, Excel environment, regression totals, cleanup status, and retained log digest.
- Added a machine-readable verification-depth inventory of the v1.0.0 release-blocking assurance controls, with explicit live command, negative proof, and current expected evidence state.
- Added mutation controls proving that incomplete-gamma dispatch drift, Student-t coefficient corruption, stale generated contract tables, and duplicated source-threshold claims are rejected.
- Added a generated 112-function `K_STATS_*` public-API manifest and a declaration-aware drift gate that blocks accidental additions, removals, renames, module moves, or compatibility-significant signature changes.
- Added a standardized installation and maintainer release documentation set with project-specific deployment, certification, provenance, recovery, and post-publication controls.
- Added a **Static checks** workflow for committed whitespace, procedure-scoped VBA jumps, release semantics (SemVer, `VERSION`/`CHANGELOG.md` agreement, dated releases, comparison links) and workflow validation with a version- and hash-pinned actionlint. The checkers are imported verbatim from the portfolio template into `tools/`, with their source revision and SHA-256 recorded in `tools/README.md`.
- Added Dependabot for GitHub Actions, proposing weekly reviewed pull requests that move each pinned action SHA and its version comment together.
- Added CodeQL security analysis (`security-extended`) of the Python tooling, the JavaScript workflow scripts and the GitHub Actions workflows. Pushes and a weekly schedule publish to code scanning; pull requests are analyzed read-only, without write permission. CodeQL does not analyze VBA.
- Added OpenSSF Scorecard for the default branch, published to the public Scorecard service and to code scanning.
- Added a daily, read-only label drift check that reports any difference between `.github/labels.json` and the live repository labels, using the same comparison rules as label sync but without write permission.
- Added a generated root README assurance block (#28). `benchmark/render_readme_assurance.py` renders the three evidence badges, the "Assurance at a glance" table and a new evidence-state list from one data model built from committed records: the public API manifest, the Excel regression record and policy, the contract registry, the accuracy and holdout summaries, both manifests, the coverage evaluation and a new committed `benchmark/release_readiness.json` registry of release blockers. Inputs are cross-checked; a missing, malformed or contradictory input writes nothing. Stale Excel, main-grid or holdout evidence is rendered as STALE and blocks. Incomplete export certification (#47, #29) and coverage debt (#22) stay visible and must be owned by an open blocker, and #34 and #35 cannot leave the registry without recorded evidence. The Accuracy Gate and `refresh_evidence.py` fail on any difference from a fresh render. The hand-verified figures and the hand-written evidence snapshot are replaced.
- Added a daily traffic export. GitHub keeps traffic data for 14 days only; the export appends views, clones, stars, forks, top referrers and popular paths to CSV files on the data-only `traffic-history` branch, writes the shields.io endpoints behind new README clones, visitors (14-day) and visits badges, and opens an alert issue on a traffic spike, a new star or fork, or a new referrer. Its read token lives in a dedicated `analytics` environment and the job runs no repository code; `SECURITY.md` documents the credential, its 90-day rotation and what the public branch contains.
- Added a correctness-only Python lint (Ruff, syntax errors and pyflakes rules via `ruff.toml`) to the **Static checks** workflow, installed from a hash-locked requirements file.
- Replaced the Markdown bug and feature templates with GitHub issue forms, and added a documentation form. Defect class, source identity (tag or full commit SHA), exact call, runnable example, independent reference, environment and whether the regression harness detects the defect are now required fields; all existing guidance, including bit-exact `hi;lo` arguments, boundary localisation and contract coverage, is retained.

- Added this changelog and the portfolio-standard release-history policy for
  VBA Probability Distributions.
- Added a root `VERSION` marker at `0.0.0`. This is a neutral pre-release
  baseline and does not claim that a functional release has been published.

### Changed

- Argument identity for the accuracy grids now has one shared definition (#32, first part). `benchmark/_binary64.py` provides `double_key`, `as_reference_arg` and `format_double_arg`; the reference generator and the five grid tools that kept byte-identical copies of the key code now import them. Generator output (1,404 reference rows) and the reports of all five tools are byte-identical before and after. `double_key` now marks non-finite tokens such as `1e999` or `inf` as unparsable instead of keying them as infinity, because VBA `Val()` never yields them; no committed argument is affected. `test_binary64.py` proves the Python side of parser parity over 34 edge values (ordinary decimals, 17-digit round trips, the normal/subnormal boundary, subnormals, the minimum subnormal, one-ulp seams, the maximum finite Double and negatives), checks every committed main and holdout argument, and pins every tool to the shared helpers. The note on double rounding above 2^52 is corrected: Doubles there are the integers, so a `.5` token is a tie that rounds half to even. VBA `Val()` parity and the rest of #32 remain open.
- Clarified PR-head versus synthetic-merge validation, required post-merge release artifact rebuild/retest/hash generation, and recorded the active branch/tag protections after PR #51 review.

- **Licence changed from MIT to the Mozilla Public License 2.0 (MPL-2.0).** `LICENSE` now holds the unmodified MPL 2.0 text. MPL 2.0 is a file-level copyleft: the library can still be used in proprietary and commercial workbooks and add-ins, but distributed modifications of files from this repository must stay under MPL 2.0 with their source available. The Exhibit A notice is given once in the README licence section rather than in each VBA module, so the certified source bytes are unchanged. Earlier revisions remain available under the MIT terms they were published with. README, INSTALLATION, CONTRIBUTING and RELEASING references and badges are updated.
- Aligned the issue labels with the portfolio template. The template's 20 core labels are adopted with its exact names, colors and descriptions, adding `behavior-change`, `security`, `repository`, `question`, `duplicate`, `invalid` and `wontfix`. `CI` becomes `ci`, and `testing` is renamed `tests` in place, so it stays on existing issues and pull requests. The nine distribution-domain labels are kept; `core` is recolored so that it no longer shares the template's `repository` color. The label manifest now meets the template's format rules (uppercase colors, sorted names, single-line descriptions of at most 100 characters), which the drift check enforces. Label sync supports declared renames and stops before any write if both the old and the new name exist.
- Hash-locked the existing mpmath/SymPy benchmark wheels and require binary, hash-verified installs in the Accuracy Gate.

- Moved all verification-depth proofs into the pre-gate evidence-tool shim so an intentionally red strict numerical gate cannot skip tests of the assurance machinery itself.
- Defined the product API explicitly as worksheet-facing `K_STATS_*` declarations; project-scoped `PROB_*` helpers and benchmark/test exporters remain outside the compatibility manifest.
- Standardized the pull-request review contract around exact-candidate evidence, compatibility, risk and recovery, security and provenance, and project-specific validation gates.
- `promote_grid_rows.py` now requires `--accept-reference-changes COUNT --reason` for any change to an existing reference, with COUNT matching the reviewed list exactly, and refuses to run on a grid containing any duplicate canonical key (#17).
- Pinned every GitHub Action to a full commit SHA with its release as a trailing comment, resolved from each action's own repository; each pin is the commit its floating major tag already referenced, so no executed code changed.
- Release guide steps 3 and 5 now name the concrete release-semantics command and require a green **Static checks** run on the exact candidate.
- Fixed 41 lint findings across the benchmark tooling: unused or redundant imports, unused variables (including two vestigial claim constants in the reference generator) and placeholder-free f-string prefixes. The generator's 1,404 rows are unchanged.
- `.editorconfig` now covers JavaScript workflow scripts (2-space indentation) and the Windows script types that `.gitattributes` already checks out with CRLF.
- Adopted a pull-request-only merge convention: every change, including the maintainer's own, reaches `main` through a squash-merged pull request. Evidence changes stay in one commit, certification binds the final merged commit, and a history-preserving merge needs a recorded reason. This is policy until the `Protect main` ruleset requires pull requests and status checks (#31).

- Future material changes must be staged here before release and describe
  observable behavior, compatibility, evidence, and known limitations.

### Fixed

- The retained Excel certification record now validates against the committed main grid (#47). Excel exports CSV with CRLF line endings and Git stores grids as LF, so the record's grid digest never matched the committed bytes. An exported grid entry now carries two exact digests: `sha256` of the committed LF bytes and `raw_sha256` of the bytes Excel wrote. The validator requires the first to equal the committed blob and the second to equal an exact line-ending form of it. The finalizer records both and refuses grids Git would not store as LF. The retained record was reconciled from its own evidence, with a `corrections` entry the validator restricts to re-expressing a recorded digest; no run, export or observation was added. The Accuracy Gate and `refresh_evidence.py` now validate the retained record itself on every change, and the README shows the main-grid export as certified. The dated evidence-snapshot paragraphs in five documents are replaced by a pointer to the generated README block.
- Clarified that the MPL exemption for independently written files does not
  extend to new files containing copied or adapted MPL-covered code (#63).

- Evidence refresh checks its Git prerequisite before any writes and reports
  an actionable error when Git cannot run. Git on PATH and full local history
  are now explicit evidence-tool prerequisites, including for Desktop users.

- README assurance rendering now validates the retained Excel candidate and
  exact source blobs with the canonical certification validator, and rejects
  all-PASS records without an observed Excel version, build or Office bitness.
  Negative fixtures cover both review findings from #60; local Git history is
  required, with no network access during rendering.

- Fixed `generate_reference_values.py` overwriting the committed grid by default. Run as previously documented, it replaced the 2,088-row `probability_accuracy_grid.csv` with 1,404 generator rows, deleting 706 rows and blanking every Excel observation. It is now report-only by default; `--out` writes reference rows only to a non-authoritative file and refuses the committed main or holdout grid and any file that already carries observations (#17).
- Fixed the `migrate_references.py` write check, which compared only the number of filled observations and so could not detect one observation replaced by another. It now asserts every observation value, row key and row order (#17).

### Documentation

- Ran the preregistered derivation of the #72 normal-tail kernel (`benchmark/normal_tail_study/`). The study fails one acceptance condition: A2 requires all coefficients to be positive, and the fitted numerator's `z^7` coefficient is −2.54E-10. That term changes the cancellation factor on `[0, 3]` by only 1.5E-7, so it has no numerical effect. Every other condition passes. The candidate's worst relative error is 5.3E-16, against 2.9E-9 for today's kernel, and the #34 route reaches 3.7E-16. The holdout was not evaluated. Per-row Excel-wave predictions are recorded for 159 Normal, Lognormal and Standard Normal rows. `test_kernel_model.py` pins the record. A decision between an amendment and a new preregistration is pending.
- Preregistered the normal upper-tail kernel repair (#72) before any coefficient is derived: `benchmark/normal_tail_study/PREREGISTRATION.md`. The current `PROB_NormalUpperTailPositive` is accurate in absolute terms only: up to 2.9E-9 relative error below its 7.07 split, and up to 5.7E-14 in the far tail because it rounds `z²` before the exponential (as does `PROB_NormalPDF`). The design is a degree-(7, 7) rational on `[0, 3)`, a 60-level continued fraction above, an exponential factor split so that its main argument is exact, and constants built from 15-digit literal pairs. Its acceptance bound is `2^-50` relative on fresh validation and holdout sets, and on the #34 route. It discloses the scratch exploration that chose the design, and freezes falsifiable predictions for every affected Normal, Lognormal and Standard Normal row in the next Excel wave. #72 is registered as an open release blocker.
- Recorded finding F1 for #34 before any VBA change (`benchmark/student_t_large_df_study/PRODUCTION_FEASIBILITY.md`). A binary64 model of the planned large-df Student-t tail route shows that its series is sound (2^-50 at every fitting point with an exact normal tail), but the existing VBA normal-tail kernel limits it to the kernel's own accuracy, up to 5.3E-10 at t = 6. With that kernel the route would be worse than the current code at 7 of 36 fitting points, against the study's preregistered improvement rule. The model reproduces every committed Excel normal-tail observation within 2 ulps (9 of 12 exactly), and `test_binary64_route_model.py` pins the finding. The production design now needs a decision and a new preregistration.
- `INSTALLATION.md` now explains that a GitHub source archive omits `.github/`, `.gitattributes`, `.gitignore`, `.editorconfig` and Git history, so maintainer validation needs a `git clone`.
- Reconciled release-readiness, provenance and certification instructions with the current evidence; corrected comparison totals, qualified numerical claims, and repaired the disabled Discussions navigation.

### Validation

- Added negative controls for Excel candidate/source/workflow SHA drift, assertion-count drift, cleanup failure, and false grid export target/hash/row-count claims.
- Added fail-closed meta-validation that rejects missing controls, missing proof scripts, unsupported states, or negative proofs not wired into the hosted pre-gate shim.
- Registered `committed-whitespace`, `vba-procedure-jumps` and `release-semantics` as release-blocking controls, each proved by its own self-test fixtures before the strict accuracy gate.
- Added the `grid-regeneration-safety` release-blocking control. `check_grid_keys.py` rejects duplicate canonical keys and non-canonical schemas in the committed main and holdout grids; `test_grid_regeneration.py` proves the generator, promotion and migration safeguards against temporary grids, plus one real default generator run verified to leave the committed grid byte-identical (#17).
- Added negative controls proving that parameter type/order, `ByVal`/`ByRef`, `Optional` defaults, return types, additive API, renames, and module moves are detected while implementation-only edits do not create API drift.
- Verified the changelog structure, policy links, section ordering, and
  repository-specific versioning scope.

### Known limitations

- Exact-SHA Excel certification (#38) is merged and inventoried. Release evidence still requires resolution of the main-grid digest mismatch [#47](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/issues/47), holdout export certification [#29](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/issues/29), coverage debt #22, and the numerical blockers tracked by #31. A green Accuracy Gate alone does not certify a release.
- Earlier project history has not been reconstructed. Existing commits, tags,
  releases, and repository documentation remain the authoritative record for
  changes made before this changelog was introduced.

---
