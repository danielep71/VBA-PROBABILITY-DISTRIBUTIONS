# 🚀 VBA Probability Distributions Release Guide

[![Release model: exact source](https://img.shields.io/badge/release-exact%20source-0969da)](#release-invariants)
[![Versioning: SemVer](https://img.shields.io/badge/versioning-SemVer-3f4551)](#versioning)
[![Evidence: required](https://img.shields.io/badge/evidence-required-success)](#evidence-record)
[![Security policy](https://img.shields.io/badge/security-policy-success)](SECURITY.md)
[![License: MPL 2.0](https://img.shields.io/badge/license-MPL_2.0-yellow.svg)](LICENSE)

This maintainer guide turns a reviewed commit into a traceable VBA Probability Distributions release. Source identity, validation, packaging, provenance, and publication must describe the same candidate.

> [!IMPORTANT]
> The first `v1.0.0` release is blocked until the relevant accuracy, regression, provenance, and coverage gates are complete. `VERSION` at `0.0.0` is not a release request.

## 🧭 Release profile

| Property | Requirement |
| --- | --- |
| Project maturity | Pre-1.0 accuracy-sensitive statistical library |
| Version source | [VERSION](VERSION) |
| Version scheme | Semantic Versioning |
| Tag format | Lower-case `vX.Y.Z` matching `VERSION` |
| Changelog | [CHANGELOG.md](CHANGELOG.md) |
| Installation contract | [INSTALLATION.md](INSTALLATION.md) |
| Vulnerability handling | [SECURITY.md](SECURITY.md) |
| License | [MPL 2.0](LICENSE) |

<a id="release-invariants"></a>
## 🔒 Release invariants

A release is valid only when all of these statements are true:

1. The candidate is identified by its full SHA and is reachable from protected `main`.
2. Release scope is frozen; unrelated work is deferred.
3. `VERSION`, changelog, source headers, package metadata, and documentation agree.
4. Static checks pass on the exact candidate.
5. VBA compiles and project-specific certification passes on that candidate.
6. Every binary artifact is built from that candidate, then tested as packaged.
7. Hashes and evidence bind artifacts to the candidate and tool environment.
8. The annotated lower-case tag points to the certified commit.
9. The GitHub Release uses that tag and the reviewed notes and assets.
10. Post-publication checks confirm that a new user can retrieve and validate it.

If an invariant is false, stop and repair the candidate. Never compensate by editing an already-tested artifact manually.

## 🗂️ Authority map

| Concern | Authoritative record |
| --- | --- |
| Current project version | `VERSION` |
| User-visible changes | `CHANGELOG.md` |
| Released source identity | Annotated Git tag and full commit SHA |
| Supported installation | `INSTALLATION.md` |
| Published binaries | GitHub Release assets |
| Integrity | SHA-256 hashes recorded with the release |
| Validation | Evidence record retained for the candidate |
| Vulnerability disclosure | `SECURITY.md` |

<a id="versioning"></a>
## 🧮 Versioning

Use `MAJOR.MINOR.PATCH` without a leading `v` inside `VERSION`.

- **MAJOR**: incompatible public API, behavior, file-format, or migration change.
- **MINOR**: backward-compatible capability.
- **PATCH**: backward-compatible correction or package/documentation fix included in the release.
- Use pre-release identifiers only when their support meaning is documented.

The Git tag adds the lower-case prefix: version `1.2.3` becomes annotated tag `v1.2.3`. Do not use upper-case `V`, moving tags, or a tag that differs from `VERSION`.

<a id="readiness-review"></a>

## ✅ Readiness review

- CORE and SPECIALFUNCS precede all distribution-family modules.
- Every advertised distribution and tail is covered by release evidence.
- Every planned item is merged or explicitly deferred.
- Compatibility and migration consequences are understood.
- Security-sensitive work has completed private handling where necessary.
- The working tree and exported VBA sources are reproducible.
- Maintainers and required reviewers are available.

Bind numerical evidence to the exact candidate, reference implementation and version, grids, tolerances, and timestamp; a summary percentage alone is insufficient. Follow [exact-SHA certification](docs/EXCEL_CERTIFICATION.md) and [provenance](benchmark/PROVENANCE.md), including validation of both exported grids against the committed evidence. Finalizer success and green fixture tests alone are insufficient: the committed record must validate against the release commit, a green **Excel VBA Regression** run of the release commit itself must exist (see [exact-tag binding](#exact-tag-binding)), and every blocker in [`benchmark/release_readiness.json`](benchmark/release_readiness.json) must be resolved.

**Enforcement baseline:** the repository settings, security features and rulesets, including the required checks `Repository integrity (static, no Excel)` and `Strict accuracy gate (pure Python, no Excel)` on `main`, are recorded in [`docs/repo_governance/repository-settings.json`](docs/repo_governance/repository-settings.json) and explained in [SETTINGS.md](docs/repo_governance/SETTINGS.md). `python3 tools/setup_repository.py --verify` compares them with GitHub and fails on drift; a prose statement of enforcement is not evidence. Required approvals remain zero. Exact-release Excel certification and the remaining release gates are tracked under [#31](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/issues/31).

## 1. Freeze and identify the candidate

1. Update refs and start from current protected `main`.
2. Create the release-preparation branch required by repository policy.
3. Record the base and candidate full commit SHAs.
4. Stop feature work on that branch.
5. Review the complete diff from the previous release tag.
6. Confirm that generated and binary changes are intentional.

```bash
git fetch --tags --prune
git rev-parse HEAD
git status --short
git diff --stat <previous-tag>...HEAD
```

For the first release there is no previous tag: omit the tag-diff command, review the entire tracked tree and its evidence, and do not fabricate a previous-release comparison link.

A dirty tree, unknown generated file, or unreviewed binary delta is blocking.

## 2. Synchronize version surfaces

Update every applicable surface in one reviewable change:

- `VERSION`
- `CHANGELOG.md`
- production module headers or documented API version
- README compatibility and accuracy claims
- accuracy-manifest and provenance metadata

Search for the prior version and retain only intentional historical references. Never rewrite historical changelog sections or immutable evidence.

## 3. Finalize the changelog

Move relevant entries from **Unreleased** into a dated `[MAJOR.MINOR.PATCH] - YYYY-MM-DD` section.

- Describe user-visible behavior, not commit mechanics.
- Use Added, Changed, Deprecated, Removed, Fixed, and Security where applicable.
- Link issues or pull requests when they improve traceability.
- Do not claim an artifact, platform, or guarantee that was not certified.
- Keep an empty Unreleased section for future work.
- Add or verify the comparison link for the new version.
- Confirm the result with `python3 tools/check_release_semantics.py --root .`,
  which rejects invalid SemVer, misordered or duplicated releases, impossible
  or regressing dates, wrong comparison links, and a release heading that
  disagrees with `VERSION`.

## 4. Verify documentation and installation

- Follow [INSTALLATION.md](INSTALLATION.md) from a clean environment.
- Verify source paths, import order, module names, prerequisites, and upgrades.
- Confirm README examples use the supported API and current version.
- Check [CONTRIBUTING.md](CONTRIBUTING.md), [SECURITY.md](SECURITY.md), and license links.
- Remove stale commands and promises beyond the tested support matrix.

## 5. Run static gates

- Validate the six-module production manifest and dependency order.
- Require a green **Static checks** workflow
  ([`.github/workflows/static-checks.yml`](.github/workflows/static-checks.yml))
  on the exact candidate SHA. It runs every static check of
  `tools/run_static_checks.py`, listed in [`tools/README.md`](tools/README.md),
  among them committed whitespace, VBA jumps, VBA source and layout rules,
  release semantics, tracked binaries, Markdown links, the Python tests, Ruff,
  mypy and actionlint. The same checks run locally with
  `python3 tools/run_static_checks.py --strict`.
- Scan for credentials or generated noise.
- Confirm accuracy manifests, grids, tolerances, and provenance files are internally consistent.
- Complete the [distribution review](docs/repo_governance/DISTRIBUTION.md) in
  its own issue, including the pinned full-history secret scan, and close it.
  Do not tag or publish while it is open.
- Verify the repository settings and keep the dated record:
  `python3 tools/setup_repository.py --verify --evidence <dir>/settings-{date}.json`
  must report `VERIFIED` (see [`tools/README.md`](tools/README.md)).

Capture commands, tool versions, timestamps, and complete results. Rerun affected gates after any change.

## 6. Certify in Excel

- Import the exact candidate in documented dependency order.
- Run **Debug → Compile VBAProject**.
- Run `Test_STATS_PROBDIST_RunAll` and every release-relevant family suite.
- Run independent accuracy benchmarks and holdout/regression checks with zero unexplained failures or coverage debt.

Certification rules:

- Use a clean workbook or documented clean fixture.
- Import only candidate files.
- Record Excel version, Windows version, and Office bitness.
- Test the advertised environment matrix.
- Treat warnings, repairs, or unexplained numerical deltas as failures.
- If code changes after certification, restart static and Excel validation.
- Finish with a green **Excel VBA Regression** run of the exact commit you will
  tag, after every documentation and evidence commit. The release workflow
  requires that run; a run of any other commit does not count (see
  [exact-tag binding](#exact-tag-binding)).

## 7. Build release artifacts

Pushing the annotated tag (step 9) runs `.github/workflows/release.yml`, which
builds the release outputs from the tagged commit:

- `VBA-Probability-Distributions-X.Y.Z.zip`: the production modules in `src/`
  and the regression harness in `tests/` with CRLF line endings for VBE import,
  plus `LICENSE`, `README.md`, `INSTALLATION.md`, `CHANGELOG.md`, `SECURITY.md`,
  `VERSION` and a `MANIFEST.sha256` of every member;
- `SHA256SUMS` for the bundle;
- `VBA-Probability-Distributions-X.Y.Z.intoto.jsonl`, a GitHub build-provenance
  attestation of both, signed in the workflow through GitHub's OIDC identity;
- the source archive GitHub creates from the tag.

The bundle is built by `.github/scripts/build_release_bundle.py` from Git
objects at the tag, never from a working tree, so the same commit always
produces the same bytes. The workflow takes the checked-out `HEAD^{commit}` as
the candidate and uses that SHA for every check, the build and the
attestation. Before building, it refuses a tag that is not annotated, not
lower-case `vX.Y.Z`, not equal to `VERSION`, not on `main`, or whose peeled
commit is not that candidate. Release semantics, the retained Excel record, the
exact-commit Excel certification below, the README assurance block and a
`certified: true` entry for the version in `benchmark/release_readiness.json`
must all validate at the tagged commit.

<a id="exact-tag-binding"></a>

**Exact-tag Excel binding (#76).** The retained record,
`benchmark/excel_regression_record.json`, names the earlier commit whose grids
Excel exported. A commit cannot contain a record of itself, so that record
proves only that the certified VBA bytes are unchanged at the tag. The workflow
therefore also requires a green **Excel VBA Regression** run of the tagged
commit itself. `.github/scripts/certify_release_candidate.py` lists the
successful `push` and `workflow_dispatch` runs of that workflow whose head
commit is the candidate, downloads the run's `excel-vba-regression-<run id>`
artifact and validates its record with:

```bash
python benchmark/check_excel_certification.py \
  --record <artifact-dir>/excel-certification.json \
  --candidate-sha <tag-commit> --evidence-ref <tag-commit> \
  --require-pass --log-directory <artifact-dir>
```

The record must also name the self-hosted runner and the run it came from. A
record for any other commit fails, even when every VBA byte is identical, for
example a run from before a documentation-only commit. The job summary states
the result, and the certifying record and log are kept for 90 days as the
`release-excel-certification` workflow artifact. Keep your own copy as release
evidence.

Produce that run before you tag. When `main` is at the commit you will tag, run
**Actions → Excel VBA Regression → Run workflow** on `main` and wait for it to
pass:

```bash
gh workflow run excel-vba-regression.yml --ref main
gh run list --workflow excel-vba-regression.yml --commit "$(git rev-parse 'main^{commit}')"
```

If `main` moves before the run starts, the run certifies the newer commit, not
the one you meant to tag. The run's artifact expires after 30 days; tag within
that time or run the workflow again. If a tag push stops only because the run
is missing, run the Excel workflow on the tag (`--ref vX.Y.Z`), then re-run the
failed release jobs.

Rehearse at any time with **Actions → Release → Run workflow**: it runs the
same checks against the exact commit of the ref you choose, builds the bundle
as a workflow artifact, and creates no attestation or release. Exact-commit
certification runs on a rehearsal too, against that commit. When the commit has
no valid Excel run of its own, the job summary reports it as **UNCERTIFIED**
with the reasons and a warning; missing readiness certification is also a
warning. A tag push fails on either, and the attestation and draft jobs run only
for a tag whose exact-commit certification is **CERTIFIED**.

A demo or validation workbook is not part of the automated release; publish one
only if the installation guide promises it, following the steps below.

For each artifact:

1. Start from a clean build location.
2. Use only candidate-controlled inputs.
3. Preserve required form and resource companions.
4. Compile before saving.
5. Exclude developer-only tests unless the artifact promises them.
6. Reopen and run the packaged smoke or regression test.
7. Record filename, size, and SHA-256.
8. Never edit the artifact after hashing.

```powershell
Get-FileHash -Algorithm SHA256 .\dist\<artifact>
```

```bash
sha256sum dist/<artifact>
```

Publish only artifacts promised by the installation guide.

## 8. Review and merge

Every change reaches `main` through a pull request (see the
[merge convention](#merge-convention) below). For a release candidate, make
these items easy to verify:

- intended version, previous tag, and candidate SHA;
- changelog and version synchronization;
- static and Excel results;
- artifact manifest and hashes;
- compatibility, migration, and security notes;
- remaining limitations.

Require configured checks and record the resulting `main` SHA. Steps 6–7 on a
pull-request candidate are pre-merge validation only. Before tagging, repeat
static and Excel certification on the final `main` commit, then rebuild every
publishable artifact from that commit in a clean location. Compile, reopen and
retest the rebuilt packages, and regenerate their filename/size/SHA-256 manifest.
Do not publish artifacts built from the PR head or synthetic merge candidate.
If `main` advances before tagging, retain the exact certified release commit or
restart certification and packaging for the new candidate; never silently retarget.

<a id="merge-convention"></a>

### Merge convention

**Decision (2026-09-30): pull requests only.** Every change reaches `main`
through a pull request, including my own changes and Dependabot updates. There
is no direct-push path.

**Decision (2026-11-02): merge commits.** The `Protect main` ruleset allows
merge commits only; the `Protect release branches` ruleset allows merge commits
or squash, and I use merge commits there too.
[`docs/repo_governance/SETTINGS.md`](docs/repo_governance/SETTINGS.md) records
both rulesets.

- **Every branch commit reaches `main`.** A merge commit keeps the branch
  history, so each commit on the branch must stand on its own: a focused change,
  a message that says what changed and why, and a tree that passes the checks.
  The pull request title and body describe the delivered change, link the
  relevant issues and evidence, and state what was validated; the merge commit
  names the pull request.
- **Release branches.** While a release branch such as `release/1.0.0` is
  active, task pull requests branch from it and target it, and Dependabot
  targets it through `target-branch` in `.github/dependabot.yml`.
  - GitHub closes an issue from a keyword only when a pull request merges into
    the default branch. After a task pull request merges into the release
    branch, I close its issue by hand with its *Resolution* section.
  - The integration pull request merges the release branch into `main` with a
    merge commit. It uses `Refs #` for the issues and pull requests it carries,
    never a closing keyword.
  - Commit and squash messages on a release branch use `Refs #N` only. A
    closing keyword in a commit message acts when the commit reaches `main`,
    not when it is written. `tools/check_closing_keywords.py` checks the
    commits of every pull request.
  - The body of the integration pull request lists every commit it carries
    that uses a closing keyword, from
    `python tools/check_closing_keywords.py --base origin/main --head HEAD --list`.
    Each is recorded in `.github/closing-keyword-exceptions.json` with its
    action. After the merge I check each issue named there, reopen any that
    GitHub closed against its recorded state, and remove the warning from the
    issue.
  - Right after each integration, a sync pull request merges `main` back into
    the release branch, so the next task branch starts from the same history.
  - Integration and sync pull requests follow the same review and check rules
    as any other pull request.
- **Merge the reviewed head with its current integration checks green.** Record
  the reviewed PR head SHA and the target base SHA. On `pull_request` events,
  the current workflows check out GitHub's synthetic merge commit; record that
  tested SHA and its workflow runs separately from the reviewed head. Require
  successful configured checks for the current head/base combination, with the
  branch up to date. A head change needs renewed review; a base change needs
  fresh integration checks. A check that has not run is not green. A required
  check cannot be waived. If a non-required check cannot run, for example because
  the self-hosted Excel runner is offline, merge only when the change cannot
  affect what it tests and record that reasoning in the PR before merging.
- **Keep evidence changes atomic.** A committed grid, its manifest and its
  export record change together, in one commit. Because a merge commit brings
  every branch commit onto `main`, that commit must not be split across the
  branch. `benchmark/check_manifest_provenance.py` examines each commit of the
  pull request's range separately, so a commit that changes a manifest without
  its grid fails the Accuracy Gate before merge. Make each such change in a
  single commit.
- **A merge creates a new source identity.** The merge commit has its own SHA,
  different from the pull request head and from the synthetic merge the checks
  ran on, so exact-SHA Excel certification of either does not certify the
  merged commit. Certify the final `main` commit, rebuild/retest all
  publishable artifacts from it and regenerate their hashes before tagging, as
  required by step 8. Never silently rebind PR evidence or relabel a pre-merge
  package as a final-commit build.
- **Commit order is evidence.** A preregistration committed before the source
  fix it governs (as in PR #27) keeps that order on `main`. Do not reorder or
  squash such commits.
- **Release history check.** The history policy of
  `tools/check_release_semantics.py`, which reports any multi-parent commit
  without a reviewed exception, applies only to the template repository it
  comes from. This repository's `.github/repository-profile.json` does not enable
  it, so pull request merge commits pass the release check.
  [`tests/test_release_history_scope.py`](tests/test_release_history_scope.py)
  pins that, with a release tag and an unreviewed merge commit.
- **Never rewrite history.** Do not force-push `main` or a release branch,
  rewrite published history or move tags. Earlier squash merges and direct
  pushes stay as they are.

GitHub enforces the PR and hosted-check requirements described in
[Readiness review](#readiness-review). Review quality, artifact reconstruction
and exact-release certification remain explicit maintainer responsibilities;
a green merge button does not establish them.

## 9. Create the annotated tag

Tag only the certified commit on `main`: the commit that the green Excel VBA
Regression run of step 6 executed. The release workflow fails for any other
commit.

```bash
git switch main
git pull --ff-only
git rev-parse HEAD
gh run list --workflow excel-vba-regression.yml --commit "$(git rev-parse HEAD)" --status success
git tag -a vX.Y.Z -m "VBA Probability Distributions X.Y.Z"
git show --no-patch --decorate vX.Y.Z
git push origin vX.Y.Z
```

Before pushing, confirm the tag equals `VERSION`, targets the certified commit, and has a matching dated changelog section. Never delete and recreate a public tag to hide a mistake.

<a id="evidence-record"></a>
## 🧾 Evidence record

Retain at least:

- candidate commit SHA
- reviewed PR head, target base and tested synthetic merge SHA (where applicable)
- static-check result
- closed distribution-review issue and its secret-scan result
- repository-settings evidence record (`VERIFIED`)
- Excel version and bitness
- regression outputs, including the Excel VBA Regression run and artifact of
  the tagged commit
- accuracy grid and independent-reference provenance
- artifact SHA-256 hashes
- version and tag
- previous release tag
- release-preparation pull request
- artifact filenames and sizes
- validation timestamps
- known deviations and approving reviewer

A release note summarizes evidence; it does not replace it.

## 10. Publish the GitHub Release

The release workflow creates a **draft** release for the tag, with the three
assets above and notes made of the version's CHANGELOG section plus an asset
table with sizes and SHA-256. Review and complete the draft, then publish it.
Before publishing, make sure it includes:

1. User-facing summary and highlights.
2. Upgrade or migration notes.
3. Supported platform statement.
4. Known limitations.
5. Installation link.
6. Asset table with SHA-256 hashes.
7. Full changelog comparison link.
8. Security-reporting link.

Upload only the already-hashed artifacts the workflow attached. Do not rebuild between tagging and upload. If the workflow failed, fix the cause, delete the draft if one exists, and re-run the workflow; never edit an asset by hand.

## 11. Verify after publication

| Check | Result |
| --- | :---: |
| Tag resolves to the certified SHA | ☐ |
| `VERSION` and changelog match the tag | ☐ |
| Assets download and hashes match | ☐ |
| Installation links and examples work | ☐ |
| Packaged artifact opens and passes its smoke test | ☐ |
| Source archive contains expected release files | ☐ |
| Default branch is ready for the next Unreleased cycle | ☐ |

Do not announce broad availability until these checks pass.

## 🧯 Recovery

### Before tag publication

Fix the branch or pull request, update evidence, and rerun affected gates. Rebuilding invalidates the prior hash and packaged-test result.

### After tag, before GitHub Release

Pause publication and document the correction. Prefer a new patch candidate whenever the tag may have propagated.

### After public release

Do not silently replace assets or move the tag.

- Mark a dangerous release clearly and remove assets only for user safety.
- Publish a corrected patch with a new tag.
- Record the problem and remedy in the changelog.
- Use [SECURITY.md](SECURITY.md) for vulnerabilities.
- Preserve evidence explaining what users received.

## ☑️ Final maintainer checklist

| # | Gate | Done |
| ---: | --- | :---: |
| 1 | Scope frozen and diff reviewed | ☐ |
| 2 | Version surfaces synchronized | ☐ |
| 3 | Changelog finalized | ☐ |
| 4 | Installation verified cleanly | ☐ |
| 5 | Static checks pass; distribution review closed; settings `VERIFIED` | ☐ |
| 6 | Excel certification passes | ☐ |
| 7 | Artifacts built and packaged-tested | ☐ |
| 8 | Hashes and evidence recorded | ☐ |
| 9 | Required pull request merged | ☐ |
| 10 | Annotated tag targets certified main SHA | ☐ |
| 11 | GitHub Release published | ☐ |
| 12 | Post-publication checks pass | ☐ |

## 📚 Related documents

- [README.md](README.md) — project overview
- [INSTALLATION.md](INSTALLATION.md) — deployment and removal
- [CONTRIBUTING.md](CONTRIBUTING.md) — change and review workflow
- [CHANGELOG.md](CHANGELOG.md) — release history
- [SECURITY.md](SECURITY.md) — supported versions and private reporting
- [DISTRIBUTION.md](docs/repo_governance/DISTRIBUTION.md) — distribution review before a public release
- [VERSION](VERSION) — authoritative version
- [LICENSE](LICENSE) — Mozilla Public License 2.0

---

**Release principle:** certify one exact source revision, build from it once, and publish only evidence-backed artifacts derived from it.
