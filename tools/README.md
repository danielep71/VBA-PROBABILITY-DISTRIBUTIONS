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
| `run_static_checks.py` | Runs every static check with one command and records the run |
| `probe_actionlint.py` | Proves that an actionlint command passes a valid workflow and rejects an invalid one |

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
| `release-bundle-self-test` | `.github/scripts/build_release_bundle.py --self-test` |
| `python-tests` | `python -m unittest discover -s tests`: every Python test under `tests/` |
| `ruff` | `python -m ruff check .`, restricted by [`ruff.toml`](../ruff.toml) to syntax errors and pyflakes rules |
| `mypy` | `python -m mypy`: strict type checking of `tools/` and `.github/scripts/`, configured in [`mypy.ini`](../mypy.ini) |
| `actionlint` | `actionlint -config-file .github/actionlint.yaml` on every workflow |
| `actionlint-probe` | `probe_actionlint.py` with the same actionlint command: a valid workflow must pass, and one with an unknown key and an undefined expression context must fail with both findings, so an invocation that checks nothing cannot pass |

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
`.md` summaries, `ruff.txt`, `mypy.txt` and `actionlint.txt`). The exit code is
1 when any check failed.

Options:

- `--strict` fails a check whose tool is missing instead of skipping it. CI
  passes it. Without it, a missing Ruff, mypy or actionlint is reported as
  `SKIP` and the run says so. Install Ruff and mypy with
  `python3 -m pip install --only-binary=:all: --require-hashes -r tools/requirements-lint-ci.txt`.
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

## Python versions

CI runs one exact Python, `3.12.16`. Each workflow that sets up Python defines
it once, as `PYTHON_VERSION` in its top-level `env`, and every
`actions/setup-python` step reads `${{ env.PYTHON_VERSION }}`:
[`static-checks.yml`](../.github/workflows/static-checks.yml),
[`accuracy-gate.yml`](../.github/workflows/accuracy-gate.yml),
[`release.yml`](../.github/workflows/release.yml) and
[`issue-triage.yml`](../.github/workflows/issue-triage.yml). To move to a new
patch release, change all four together; first confirm that
[actions/python-versions](https://github.com/actions/python-versions) publishes
it for `ubuntu-24.04` x64. The Accuracy Gate's reference wheels
(`benchmark/requirements-ci.txt`) are pure Python, so a patch release does not
change them. I run the issue audit on the same version as its tests rather than
on a newer series, so the audited and the tested interpreter are the same.
`tests/test_python_versions.py` fails if a workflow uses a floating or literal
version, or if the workflows disagree.

Two jobs run short inline scripts on the runner image's own `python3` and set
up no Python: *Keep milestones issue-only* runs no action at all by design, and
the daily traffic export keeps third-party code out of the job that holds its
analytics token.

The minimum Python is **3.11**: every tool under `tools/` and `.github/scripts/`
and every test under `tests/` runs on it, using the standard library only (the
Accuracy Gate's reference stack is pinned separately, on the CI version). I
chose 3.11 because it is the oldest CPython series that still receives security
fixes; 3.10 reached end of life in October 2026. The non-required job
*Minimum Python (3.11, not required)* in `static-checks.yml` runs
`run_static_checks.py` on `MINIMUM_PYTHON_VERSION` (`3.11.18`): every check that
needs no installed tool, the unit tests included. `mypy.ini` (`python_version`)
and `ruff.toml` (`target-version`) check the code against the same floor.

## Dependency updates

Dependabot checks the GitHub Actions pins, `benchmark/requirements-ci.txt`
(mpmath and SymPy), and `tools/requirements-lint-ci.txt` (Ruff, mypy and mypy's
dependencies, in one grouped pull request) weekly. Updates
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
The lint lock lists, for Ruff, mypy and each of mypy's dependencies, the hash of
every file PyPI publishes for the pinned release (every wheel and the sdist),
checked against PyPI's release metadata; CI installs the Ubuntu x86-64 wheels
with `--only-binary=:all: --require-hashes`. The lock must hold mypy's complete
dependency closure: `--require-hashes` refuses any package it does not list. actionlint is separately version- and checksum-pinned in the static
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
| `tests/test_pr_milestones_workflow.py` | Offline tests of that workflow: a dependency-free structure check with negative fixtures in place of the template's YAML-parser invariants, and the template's script tests extended to the job summary and the read-only-token path. |
| `requirements-lint-ci.txt`, `mypy.ini` | mypy 2.4.1 and its dependency closure from `tools/requirements-quality-ci.txt`, with the hash of every published file, and the strict settings of `pyproject.toml` `[tool.mypy]`, scoped to `tools/` and `.github/scripts/`. |
| `probe_actionlint.py`, `tests/test_quality_gates.py` | The actionlint unknown-key and mypy strict-mode probes of `tests/test_quality_gates.py`; the actionlint probe is a runner check that reuses the actionlint check's own command, with stand-in commands as offline fixtures. |
| `.github/workflows/static-checks.yml` (minimum-python job) | The `minimum-python` job of `.github/workflows/repository-integrity.yml`, running this repository's runner without installed tools. |
