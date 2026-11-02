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
| `format_vba.py` | Checks (`--check`) or applies (`--write`) the [VBA source layout](#vba-source-layout), ratcheting `vba_layout_baseline.json`; its fixtures are [`tests/test_format_vba.py`](../tests/test_format_vba.py) |
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
| `vba-layout` | `format_vba.py --check` on every tracked VBA module, against `vba_layout_baseline.json` |
| `release-semantics` | `check_release_semantics.py` |
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

## VBA source layout

[`format_vba.py`](format_vba.py) states the layout of every tracked `.bas`,
`.cls` and `.frm` module under `src/`, `tests/` and `benchmark/`:

1. CRLF line endings, one final line ending and no blank lines at the end of the
   file;
2. no trailing white space;
3. indentation with spaces: each tab of a line's indentation becomes four spaces;
4. procedure declarations (`Sub`, `Function`, `Property Get/Let/Set`, with any
   `Public`, `Private`, `Friend` or `Static`), their `End` statements and labels
   such as `Err_Handler:` start in column one;
5. exactly two blank lines after each procedure's `End` statement, before the
   next procedure, its comment banner or a section banner.

Every other line keeps its indentation. That covers this repository's indented
comment headings inside procedures (`'Declare`, `'Invert`, ...) with their
statements indented under them, the documentation banner after a declaration,
continuation lines (a line after one ending in ` _`, including a continued
comment, is never moved) and blank lines inside procedures and between
module-level declarations. String literals, comments and the header the VBE
writes (`VERSION`, `BEGIN ... END`, `Object` and `Attribute` lines) are never
changed; statement keywords followed by a colon (`Else:`, `Loop:`) are not
labels. The tests prove, on the text of every tracked module, that formatting
leaves the token stream (everything but white space and line layout) unchanged.

I chose these rules from the modules as they are. When I added the check (module
bytes as of `937e3a6`), the only differences between the 35 tracked modules and this layout
are blank lines: 38 missing and 37 surplus, in 17 modules. 494 of the 535 gaps
after an `End` statement are already two blank lines (36 are one, one is none,
four are three or four), and 20 of the 35 modules already end without blank
lines (15 end with two or four). No module has a tab, trailing white space, an
LF-only line ending, or a procedure, `End` or label line out of column one.

```bash
python3 tools/format_vba.py --check            # every tracked module, with the baseline
python3 tools/format_vba.py --check --no-baseline
python3 tools/format_vba.py --write src/M_STATS_PROBDIST_CORE.bas
```

Every module is bound by hash to the Excel certification record or the
observation manifest, so I do not reformat one outside an Excel evidence wave.
Until then, [`vba_layout_baseline.json`](vba_layout_baseline.json) lists the
modules still out of layout, and the `vba-layout` check ratchets it: a module
out of layout that it does not list fails, a listed module that is in layout
fails until its entry is removed, and a listed path that is not a tracked
module fails. To reformat during an Excel wave, run `--write` on the listed
modules, re-import them into the VBE, refresh their evidence, remove their
entries, and delete the baseline file once it lists none. Exit codes: 0 pass,
1 a module or baseline entry fails, 2 an unreadable or malformed baseline.

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
| `format_vba.py` | `tools/format_vba.py`: the same five layout rules, kept because they match this repository's dominant layout; tracked modules in place of the template's source directories, a committed baseline with a two-way ratchet for modules bound to Excel evidence, `--no-baseline`, statement keywords excluded from labels, and continued comments left in place. |
| `tests/test_format_vba.py` | `tests/test_format_vba.py`: its rule, literal and header cases, extended with token-stream, idempotence and drift-undo checks on every tracked module's text and the baseline ratchet in temporary Git repositories. |
| `tests/test_pr_milestones_workflow.py` | Offline tests of that workflow: a dependency-free structure check with negative fixtures in place of the template's YAML-parser invariants, and the template's script tests extended to the job summary and the read-only-token path. |
