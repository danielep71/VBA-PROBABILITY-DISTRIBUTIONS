# Repository static gates

Generic, Excel-free checks of the repository itself. They complement the
numerical evidence machinery under [`benchmark/`](../benchmark/) and the Excel
regression under [`ci/`](../ci/); they never prove compilation, regression or
accuracy.

The three checkers below are each proved by their own `--self-test` fixtures,
which also run before the strict accuracy gate through
[`benchmark/test_evidence_tools.py`](../benchmark/test_evidence_tools.py) and are
registered in [`benchmark/verification_depth.json`](../benchmark/verification_depth.json).

| Tool | Checks |
| --- | --- |
| `check_committed_whitespace.py` | No whitespace errors in the committed candidate range (`git diff --check` semantics, read only) |
| `check_vba_jumps.py` | Every `GoTo`, `GoSub`, `Resume` and `On Error GoTo` target resolves inside its own procedure, across all tracked `.bas`/`.cls`/`.frm` modules |
| `check_release_semantics.py` | `VERSION` is strict SemVer, `CHANGELOG.md` has exactly one `[Unreleased]`, release headings are ordered and dated, and comparison links match the repository |
| `_gatelib.py` | Shared command-line and report helpers for the three checks above |
| `check_binaries.py` | Every tracked binary is an approved `assets/` image or has a reviewed SHA-256 in [`.github/binary-allowlist.json`](../.github/binary-allowlist.json) ([below](#tracked-binaries)); tested by `tests/test_check_binaries.py` |
| `run_static_checks.py` | Runs every static check with one command and records the run |

## Running every static check

One command runs every static check, locally and in
[`.github/workflows/static-checks.yml`](../.github/workflows/static-checks.yml)
on every push to `main` and `release/**`, every pull request to them and every
`v*` tag:

```bash
python3 tools/run_static_checks.py
```

It works on the repository that contains it, from any directory. It runs:

| Check | Command |
| --- | --- |
| `whitespace-self-test`, `vba-jumps-self-test`, `release-semantics-self-test` | Each checker's `--self-test` fixtures |
| `whitespace` | `check_committed_whitespace.py --mode committed` on the candidate range |
| `vba-jumps` | `check_vba_jumps.py` on every tracked VBA module |
| `release-semantics` | `check_release_semantics.py` |
| `binaries` | `check_binaries.py` on every tracked file ([tracked binaries](#tracked-binaries)) |
| `release-bundle-self-test` | `.github/scripts/build_release_bundle.py --self-test` |
| `python-tests` | `python -m unittest discover -s tests`: every Python test under `tests/` |
| `ruff` | `python -m ruff check .`, restricted by [`ruff.toml`](../ruff.toml) to syntax errors and pyflakes rules |
| `actionlint` | `actionlint -config-file .github/actionlint.yaml` on every workflow |

Every check runs even when an earlier one fails. The run writes, under
`test-results/static-checks/`:

- one `<check>.log` per check, with its command and output;
- `report.json`: the commit, whether the working tree is dirty, the pull request
  head, the Python version and platform, the SHA-256 of every tracked file (its
  working-tree bytes), and each check's command, status and exit code;
- `summary.md`, the Markdown summary, followed by the summaries of the
  whitespace, VBA-jump and release-semantics checks.

The checks still write their own evidence to `test-results/`
(`committed-whitespace.json`, `vba-jumps.json`, `release-semantics.json`, their
`.md` summaries, `ruff.txt` and `actionlint.txt`). The exit code is 1 when any
check failed.

Options:

- `--strict` fails a check whose tool is missing instead of skipping it. CI
  passes it. Without it, a missing Ruff or actionlint is reported as `SKIP` and
  the run says so. Install Ruff with
  `python3 -m pip install -r tools/requirements-lint-ci.txt`.
- `--head REV` and `--base REV` set the committed-whitespace range (default:
  `$CANDIDATE_SHA`, else `HEAD`; `$BASE_SHA`, else none).
- `--pr-head SHA` records the pull request head (default: `$PR_HEAD_SHA`).
- `--actionlint PATH` uses that executable and never one on `PATH` (default:
  `$ACTIONLINT`, else `PATH`). CI passes its checksum-verified download.
- `--summary FILE` also appends the summary to `FILE`, such as
  `$GITHUB_STEP_SUMMARY`.

The run covers the repository only: it compiles no VBA, runs no Excel and
certifies no figure.

### Adding a check

Add one `Check(name, command)` entry to the list returned by `checks()` in
[`run_static_checks.py`](run_static_checks.py). The command runs from the
repository root, and a nonzero exit code fails the check. Optional fields:
`missing` (why its tool is unavailable: the check is skipped, or fails under
`--strict`), `output` (a file that also receives its output) and `summary` (a
Markdown file it writes, added to the run summary). Then add the name to
`test_every_static_check_is_listed` in
[`tests/test_run_static_checks.py`](../tests/test_run_static_checks.py). The
workflow needs no change unless the check needs a new tool installed. A new
`tests/test_*.py` file needs no entry: `python-tests` discovers it.

## Tracked binaries

A binary file cannot be reviewed in a diff, and the demo workbook can carry VBA
macros. `check_binaries.py` therefore admits a tracked binary only by an
approved path or by a reviewed SHA-256:

```bash
python3 tools/check_binaries.py
```

It reads every regular file in the Git index (in CI, the checked-out commit)
and treats a file as binary when [`.gitattributes`](../.gitattributes) unsets
its `text` attribute (the `office-doc`, `opaque` and `vba-form-bin` macros) or
its content contains a NUL byte. It prints every binary it allowed, with its
SHA-256, and fails on:

- a binary that is neither an approved image nor in the allowlist;
- an allowlisted file whose SHA-256 differs from the recorded one, whatever it
  now contains;
- an allowlist entry for a file that is not tracked;
- a NUL-containing file that `.gitattributes` does not declare binary. Git may
  rewrite the line endings of such a file, so an allowlist entry does not admit
  it: add a binary rule to `.gitattributes` first;
- an allowlist that is not valid JSON, or an entry without a 64-digit lowercase
  `sha256` and a non-empty `reason`.

The approved path is narrow on purpose: a `.png`, `.jpg` or `.jpeg` file
directly in `assets/` (the README and social-preview images), whose bytes start
with the signature of that format, so a renamed workbook or archive does not
pass as an image. Every other binary needs an entry in
[`.github/binary-allowlist.json`](../.github/binary-allowlist.json). I allow the
demo workbook `examples/STATS-Distributions demo.xlsm` only there, by digest, so
that any new version of it fails until someone reviews it.

The digest is the SHA-256 of the blob Git stores. Git applies no line-ending
conversion to a file declared binary, so it equals the SHA-256 of the
checked-out file. The allowlist is read from the index too: stage a change
before checking it.

To add or update a binary after reviewing it:

1. For the workbook, open it with macros disabled and inspect its VBA project
   and any external links before trusting the new version.
2. Stage the file and run the check. The failure prints the file's current
   SHA-256 and the entry to write.
3. Set that `sha256` in `.github/binary-allowlist.json`, update the `reason`
   (why the file is tracked and what was reviewed), and commit both in the same
   pull request, so the digest change is reviewed with the file.

Remove the entry when you remove the file.

## Dependency updates

Dependabot checks the GitHub Actions pins, `benchmark/requirements-ci.txt`
(mpmath and SymPy), and `tools/requirements-lint-ci.txt` (Ruff) weekly. Updates
are assigned to the maintainer with priority P3 for initial triage; raise the
priority when a security advisory or other material risk requires it. Review
and merge remain manual. Pull requests carry no milestone: during PR triage,
link each one to an issue and set the milestone on that issue
([`ISSUE_LABELS.md`](../docs/repo_governance/ISSUE_LABELS.md)).

Both CI Python environments install wheels with `--require-hashes`. Review
version and SHA-256 changes together, compare hashes with PyPI's release
metadata and the downloaded wheel, and retain a complete pinned dependency
closure. Reference-library changes require the full assurance and numerical
gates; do not rewrite committed observations to accommodate a dependency bump.
Ruff's lock lists the hash of every published wheel (one per platform) and the
sdist, and Dependabot regenerates the whole set; CI installs the Ubuntu x86-64
wheel. actionlint is separately version- and checksum-pinned in the static
workflow and needs a reviewed manual update.

The optional gmpy2/MPFR and R/Rmpfr research tools are outside these CI locks;
retain their recorded versions with study evidence when rerunning them. The VBA
library itself has no downloaded package dependencies. CodeQL analyzes Python,
JavaScript and Actions, not VBA; a successful analysis run is not proof that
there are no open findings.

## Provenance

These four files are imported **verbatim** from the portfolio template
[`danielep71/EXCEL-VBA-PROJECT-TEMPLATE`](https://github.com/danielep71/EXCEL-VBA-PROJECT-TEMPLATE)
at commit `b903fe4` (template v1.3.0). Fix defects upstream and re-import, so the
copies stay identical to a published template revision.

| File | SHA-256 at import |
| --- | --- |
| `_gatelib.py` | `3af45805f46a1a2535e4be9a334e79c49a05228338e4763e056fe076844c1289` |
| `check_committed_whitespace.py` | `b4b252922567f16bcaea31aa4a3a028289dad7cd7364a65dae85cd5a9a3d98ce` |
| `check_vba_jumps.py` | `fe261b957442551dff8394522de875c8ecc57ad778ae9af9d41da58061efc17e` |
| `check_release_semantics.py` | `94ee669082474da390d685f706c124fff53b9b67d6880cea67e8b92dcf602381` |

`check_release_semantics.py` reads a single field, `repository`, from
[`.github/repository-profile.json`](../.github/repository-profile.json); it uses it
to validate changelog comparison links. That file deliberately records nothing
else. This repository has **not** adopted the template contract, a template
profile or the template's directory layout, and the file must not be read as
claiming any of them. The template's release-history policy applies only to the
template repository itself and is not imported.

### Adapted from `danielep71/MyTemplate` at `9ff7670`

These files adapt a template solution to this repository's own names, checks and
layout; they are not verbatim copies.

| File | Adaptation |
| --- | --- |
| `run_static_checks.py` | The runner of `tools/check_repository.py`: every check runs, per-check logs, `report.json` (commit, dirty state, pull request head, tracked-file SHA-256), Markdown summary and `--strict`. The check list is this repository's own. |
| `tests/test_run_static_checks.py` | The approach of `tests/test_check_repository.py`: injected checks run in temporary Git repositories. |
| `.github/workflows/static-checks.yml` | The single runner step of `.github/workflows/repository-integrity.yml`, with the report directory kept as an artifact. |
| `.github/workflows/pr-milestones.yml` | Pull-request milestone removal on `pull_request_target`, with this repository's workflow conventions, a job summary of each removal and a warning instead of a failure when a pull-request event gets a read-only token. |
| `check_binaries.py` | The binary detection of `tools/check_binaries.py` (`text` unset in `.gitattributes`, or a NUL byte) and its allowlist schema. Here it reads the index rather than the working tree, admits only signature-checked images directly in `assets/`, prints the digest to record, and fails on a NUL-containing file that `.gitattributes` treats as text even when allowlisted. |
| `tests/test_check_binaries.py` | The temporary-repository fixtures of the template's tests, extended to the digest-pinned demo workbook, the image signature, the stale entry and the unstaged allowlist. |
| `.github/binary-allowlist.json` | The template's schema (`schema_version`, and a `sha256` and `reason` per path), holding this repository's demo workbook. |
| `tests/test_pr_milestones_workflow.py` | Offline tests of that workflow: a dependency-free structure check with negative fixtures in place of the template's YAML-parser invariants, and the template's script tests extended to the job summary and the read-only-token path. |
