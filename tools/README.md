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
| `check_vba_source.py`, `check_vba_conditionals.py`, `check_vba_visibility.py` | The VBA source rules of [VBA source checks](#vba-source-checks); their fixtures are unit tests in [`tests/test_check_vba.py`](../tests/test_check_vba.py), not `--self-test` |
| `vba_lexer.py` | Shared comment, string, continuation and statement lexing for those three checks |

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
| `vba-source`, `vba-conditionals`, `vba-visibility` | `check_vba_source.py`, `check_vba_conditionals.py` and `check_vba_visibility.py` on every tracked VBA module |
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

## VBA source checks

Three checks read every tracked `.bas`, `.cls` and `.frm` file: the production
modules in `src/`, the test module in `tests/`, and the exporters and probes
under `benchmark/`. `run_static_checks.py` runs them; each also runs on its
own, for example `python3 tools/check_vba_source.py`. Their positive and
negative fixtures are in [`tests/test_check_vba.py`](../tests/test_check_vba.py).
They share the lexer in `vba_lexer.py`, so a keyword in a comment (`'Declare`,
`Rem Public Sub`), in a continued comment or in a string literal is never read
as code.

| Check | Enforces |
| --- | --- |
| `check_vba_source.py` | Each component in `src/` or `tests/` (no subdirectories) or under `benchmark/`; its file name starts with `M_STATS_PROBDIST_`, has at most 31 characters and equals its one `Attribute VB_Name`, unique ignoring case; a `.bas` export starts with that line; ASCII only; `Option Explicit`; `Option Private Module` in the modules of `OPTION_PRIVATE_REQUIRED` and never in a module that declares a `Public K_STATS_` procedure; no `Option Base`, tab or trailing white space; closed string literals; LF in the Git index with CRLF checkout (`git ls-files --eol` and the `vba-source` attribute of `.gitattributes`) |
| `check_vba_conditionals.py` | `#If`, `#ElseIf`, `#Else`, `#End If` and module `#Const` evaluated in three compilation environments: VBA6 Win32, VBA7 Win32 and VBA7 Win64. An unknown symbol, an unbalanced block, or a malformed, duplicate or non-Boolean `#Const` fails. Every `Declare` reachable under VBA7 is `PtrSafe` |
| `check_vba_visibility.py` | Every procedure and module-level declaration states `Public` or `Private` (`Global` counts as `Public`), and `Friend` appears only in a class module. No public name (procedure, `Const`, `Enum`, `Type`, variable or `Declare`) is declared twice in one standard module or in two standard modules of a collision group, in an environment they share |

**Option Private Module.** I require it in `M_STATS_PROBDIST_CORE` and
`M_STATS_PROBDIST_SPECIALFUNCS`, the internal numerical layers: their
`Public PROB_*` routines are shared across the project but stay out of the
worksheet Function Wizard and out of other projects. The four modules that
declare the `K_STATS_*` worksheet functions must not have it, because it would
hide those functions. Today the two internal layers have it and no other module
does. The test module and the benchmark exporters are neither required nor
forbidden to have it. A path in `OPTION_PRIVATE_REQUIRED` that is no longer a
tracked component fails the check, so a rename cannot drop the rule silently.

**Compilation environments.** The symbols are Microsoft's compiler constants:
`VBA6` (True under VBA7 too), `VBA7`, `Win16` (always False here), `Win32` (True
on 64-bit Windows too), `Win64` and `Mac` (always False here). A module `#Const`
is defined from its line on, in the environments that reach it, and its value
may use the compiler constants and the module's earlier `#Const` symbols. A
symbol defined nowhere fails, although VBA would read it as `Empty`: a
project-level conditional-compilation argument is invisible in an exported
module, and the check fails closed rather than guess.
`M_STATS_PROBDIST_CORE` uses `#Const PROB_WRITE_STATUS_BAR = False`, and
`benchmark/M_STATS_PROBDIST_ACCURACYEXPORT.bas` uses `#If Win64`. No module
contains a `Declare` statement today.

**Collision groups.** One group, `one VBA project` (`COLLISION_GROUPS`), holds
every standard module under `src/`, `tests/` and `benchmark/`. The Excel
regression imports the production and test modules into one workbook
(`regression_sources` in
[`.github/excel-evidence-policy.json`](../.github/excel-evidence-policy.json)),
each exporter or probe is imported into a workbook that already holds the
production modules, and the main-grid manifest binds production, test, exporter
and study modules together. One group is the superset of every import set, so
no combination of them can raise *Ambiguous name detected*. `Enum` members and
private names are not compared, and class-module members are not in the
project's global namespace. `benchmark/check_public_api.py` remains the freeze
of the `K_STATS_*` signatures; this check keeps no inventory of its own.

What these checks cannot prove:

- they neither compile nor run VBA, and they resolve no call;
- they model the three environments statically, not on a 32-bit or 64-bit
  host, and do not prove that a `Declare` uses the right argument widths;
- `check_vba_jumps.py` keeps its own lexer and reads both branches of an `#If`
  together rather than per environment.

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
| `tests/test_pr_milestones_workflow.py` | Offline tests of that workflow: a dependency-free structure check with negative fixtures in place of the template's YAML-parser invariants, and the template's script tests extended to the job summary and the read-only-token path. |
| `vba_lexer.py` | The shared lexer of `tools/vba_lexer.py`, without the facade-prefix and house-style lookups. |
| `check_vba_source.py` | `tools/check_vba_source.py`: this repository's flat `src/` and `tests/` and nested `benchmark/` homes and `M_STATS_PROBDIST_` names in place of role homes and prefixes; `Option Private Module` required in a named set of modules and forbidden beside `K_STATS_` worksheet functions; no form or workbook rules. |
| `check_vba_conditionals.py` | `tools/check_vba_conditionals.py`, extended to module `#Const` symbols and to the `Win16` and `Mac` constants, with one directive walk shared by the reachable-source model and the `PtrSafe` rule. |
| `check_vba_visibility.py` | The visibility and collision parts of `tools/check_vba_public_api.py`, without its inventory, signature and public-variable rules; multi-identifier `Const` and variable statements are read name by name, and collisions are evaluated per collision group. |
| `tests/test_check_vba.py` | The fixture repositories and the source, conditional, public-API and lexer cases of `tests/test_check_vba.py`, rewritten for this repository's modules and rules. |
