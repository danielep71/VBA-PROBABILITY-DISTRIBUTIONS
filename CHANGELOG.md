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

- Added exact-SHA Excel certification records binding the candidate commit, imported VBA source hashes, Excel environment, regression totals, cleanup status, and retained log digest.
- Added a machine-readable verification-depth inventory of the v1.0.0 release-blocking assurance controls, with explicit live command, negative proof, and current expected evidence state.
- Added mutation controls proving that incomplete-gamma dispatch drift, Student-t coefficient corruption, stale generated contract tables, and duplicated source-threshold claims are rejected.
- Added a generated 112-function `K_STATS_*` public-API manifest and a declaration-aware drift gate that blocks accidental additions, removals, renames, module moves, or compatibility-significant signature changes.
- Added a standardized installation and maintainer release documentation set with project-specific deployment, certification, provenance, recovery, and post-publication controls.
- Added a **Static checks** workflow for committed whitespace, procedure-scoped VBA jumps, release semantics (SemVer, `VERSION`/`CHANGELOG.md` agreement, dated releases, comparison links) and workflow validation with a version- and hash-pinned actionlint. The checkers are imported verbatim from the portfolio template into `tools/`, with their source revision and SHA-256 recorded in `tools/README.md`.
- Added Dependabot for GitHub Actions, proposing weekly reviewed pull requests that move each pinned action SHA and its version comment together.

- Added this changelog and the portfolio-standard release-history policy for
  VBA Probability Distributions.
- Added a root `VERSION` marker at `0.0.0`. This is a neutral pre-release
  baseline and does not claim that a functional release has been published.

### Changed

- Moved all verification-depth proofs into the pre-gate evidence-tool shim so an intentionally red strict numerical gate cannot skip tests of the assurance machinery itself.
- Defined the product API explicitly as worksheet-facing `K_STATS_*` declarations; project-scoped `PROB_*` helpers and benchmark/test exporters remain outside the compatibility manifest.
- Standardized the pull-request review contract around exact-candidate evidence, compatibility, risk and recovery, security and provenance, and project-specific validation gates.
- Made `generate_reference_values.py` report-only by default. It no longer writes `probability_accuracy_grid.csv`: run as previously documented, it replaced the 2,088-row grid with 1,404 generator rows, deleting 706 rows and blanking every Excel observation. `--out` writes reference rows only to a non-authoritative file and refuses the committed main or holdout grid and any file that already carries observations (#17).
- `promote_grid_rows.py` now requires `--accept-reference-changes COUNT --reason` for any change to an existing reference, with COUNT matching the reviewed list exactly, and refuses to run on a grid containing any duplicate canonical key (#17).
- `migrate_references.py` now asserts that every observation value, row key and row order is unchanged, rather than only the number of filled observations (#17).
- Pinned every GitHub Action to a full commit SHA with its release as a trailing comment, resolved from each action's own repository; each pin is the commit its floating major tag already referenced, so no executed code changed.
- Release guide steps 3 and 5 now name the concrete release-semantics command and require a green **Static checks** run on the exact candidate.

- Future material changes must be staged here before release and describe
  observable behavior, compatibility, evidence, and known limitations.

### Validation

- Added negative controls for Excel candidate/source/workflow SHA drift, assertion-count drift, cleanup failure, and false grid export target/hash/row-count claims.
- Added fail-closed meta-validation that rejects missing controls, missing proof scripts, unsupported states, or negative proofs not wired into the hosted pre-gate shim.
- Registered `committed-whitespace`, `vba-procedure-jumps` and `release-semantics` as release-blocking controls, each proved by its own self-test fixtures before the strict accuracy gate.
- Added the `grid-regeneration-safety` release-blocking control. `check_grid_keys.py` rejects duplicate canonical keys and non-canonical schemas in the committed main and holdout grids; `test_grid_regeneration.py` proves the generator, promotion and migration safeguards against temporary grids, plus one real default generator run verified to leave the committed grid byte-identical (#17).
- Added negative controls proving that parameter type/order, `ByVal`/`ByRef`, `Optional` defaults, return types, additive API, renames, and module moves are detected while implementation-only edits do not create API drift.
- Verified the changelog structure, policy links, section ordering, and
  repository-specific versioning scope.

### Known limitations

- Exact-SHA Excel certification from P0.1 / issue #38 is still being integrated separately; once merged into `main`, that release-blocking control must be added to `verification_depth.json` in the same integration change.
- Earlier project history has not been reconstructed. Existing commits, tags,
  releases, and repository documentation remain the authoritative record for
  changes made before this changelog was introduced.

---

