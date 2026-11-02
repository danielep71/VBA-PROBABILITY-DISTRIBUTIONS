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
| `check_vba_source.py`, `check_vba_conditionals.py`, `check_vba_visibility.py` | The VBA source rules of [VBA source checks](#vba-source-checks); their fixtures are unit tests in [`tests/test_check_vba.py`](../tests/test_check_vba.py), not `--self-test` |
| `vba_lexer.py` | Shared comment, string, continuation and statement lexing for those three checks |
| `probe_actionlint.py` | Proves that an actionlint command passes a valid workflow and rejects an invalid one |
| `check_links.py` | Relative links and anchors in every tracked Markdown file (offline); external links with `--external` ([Markdown links](#markdown-links)) |

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
| `binaries` | `check_binaries.py` on every tracked file ([tracked binaries](#tracked-binaries)) |
| `release-bundle-self-test` | `.github/scripts/build_release_bundle.py --self-test` |
| `markdown-links` | `check_links.py`: relative links and anchors in every tracked Markdown file |
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

## CI test coverage

`check_ci_test_coverage.py` fails when a tracked `test_*.py` file does not run
in CI on every pull request. It runs on every event in the Accuracy Gate's
`Select accuracy proofs` job, after `python -m unittest discover -s tests`, and
its live check and fixtures (`tests/test_ci_test_coverage.py`) are registered in
`benchmark/verification_depth.json`.

The checker is deliberately conservative: it recognizes a few step shapes and
reports everything else, so an ambiguous step has to be restructured. A file
counts as run when either of these executes it:

- a `run:` step of a workflow whose `pull_request` trigger has no path filter,
  no `branches-ignore` and no negative (`!`) branch pattern, and whose
  `branches`, if any, include the default branch.
  The step must use the default or `bash` shell and must not be
  `continue-on-error`. Every line must be recognized: blank lines, comments, a
  top-level `cd <dir>`, plain commands (redirections and leading assignments
  allowed), and `python`/`python3` running `<file>`,
  `-m unittest <file or module>` or `-m unittest discover`. One other line
  withholds credit from the whole step: a command list, a pipeline, a subshell,
  a substitution, a heredoc, a loop, `exit`, `set +e` or another `if`.
  unittest counts only modules that define a `unittest.TestCase`, and discovery
  only inside packages without a `load_tests` hook.
- an entry of a runner list that such a step invokes. Only
  `benchmark/test_evidence_tools.py` is registered, and its `--checks-only`
  filter is applied. `benchmark/refresh_evidence.py` is not registered: CI never
  runs it.

A run gated by the accuracy-proof selection counts only if
`select_accuracy_checks.py` does not treat a change to that test file as
metadata-only. Two gates are recognized:

- `if: needs.select.outputs.full == 'true'`;
- the `then` branch of a single
  `if [[ "$VAR" == "true" ]]; then` ... `else` ... `fi` block, where `VAR` is
  bound exactly to `${{ needs.select.outputs.full }}`. The `else` branch does
  not count. The operand must be double-quoted (`"$VAR"`, `"${VAR}"` or the
  expression itself) and compared with `true` by `==` or `=`, in `[[ ]]` or
  `[ ]`. The gate is matched on the raw line, so quoting that stops expansion,
  such as `'$VAR'`, is reported rather than credited.

Otherwise a job or step `if:` counts only when it is exactly `always()`,
`success()` or `!cancelled()`. A job gets no credit when a job it needs has any
other condition. A step that may rebind the selection variable gets none either:
`VAR=`, `export`, `declare`, `read`, `printf -v`, `${VAR:=...}` and `unset` all
count as rebinding. The report names each withheld step and the reason.

**Threat model and limits.** The guard catches a test file that no CI step runs
by accident: a new file nobody listed, a renamed one, or a step that silently
skips it. It does not try to defeat a deliberately misleading workflow; review
does that. Anything outside the recognized shapes is reported rather than
credited. Its failure mode is therefore a false report, which you resolve by
simplifying the step, never a false pass. Workflows are read with a small parser for block-style YAML; a
construct outside it is an error (exit 2), not a pass.

To add a test, add a benchmark test to the `commands` list in
`benchmark/test_evidence_tools.py`. Put a `unittest` module under `tests/`,
where discovery picks it up; its file and package names must be Python identifiers, because discovery silently skips any other name. `EXEMPT` in the checker takes a path and a reason.
I keep it empty; the checker reports an exemption that is stale or already
covered.

```bash
python3 tools/check_ci_test_coverage.py --root .
python3 -m unittest discover -s tests
```

## Repository settings

[`setup_repository.py`](setup_repository.py) compares the live GitHub settings
with the baseline in
[`docs/repo_governance/repository-settings.json`](../docs/repo_governance/repository-settings.json):
visibility, merge methods and repository features, the five security features,
and the five repository rulesets with their required checks and allowed merge
methods. [SETTINGS.md](../docs/repo_governance/SETTINGS.md) explains each
value. It calls the REST API through an authenticated `gh` CLI and needs
Python 3.10 or later.

```bash
python3 tools/setup_repository.py                  # preview: differences and the plan; read-only
python3 tools/setup_repository.py --verify \
  --evidence settings-{date}.json                  # read-only; exit 1 on drift
python3 tools/setup_repository.py --apply          # writes the plan, then reads everything back
```

- **Preview** (the default) lists every difference and the writes that would
  remove it. **Verify** reads the same state and exits with status 1 on any
  difference, including a value GitHub does not report to the token in use.
  Neither writes.
- **Apply** needs administrative access and refuses to run unless the
  repository's visibility already matches the baseline (the tool never
  changes visibility) and, before it writes a ruleset, every required check
  has succeeded on the default-branch HEAD. Replacing a ruleset that exists
  under a baseline name also needs `--update-existing`, because the whole
  ruleset is replaced. There is no automatic rollback; rerun the preview after
  a failure.
- `--evidence PATH` writes a JSON record of every managed value as observed,
  with the check time, the baseline's SHA-256 and the result (`VERIFIED` or
  `DRIFT`). `{date}` in the path becomes the UTC date.
- Rulesets outside the baseline are named as notes and never changed.
- When the `vulnerability-alerts` endpoint cannot be read, enabled Dependabot
  security updates count as proof that alerts are on, because GitHub allows
  them only with alerts; the record names that source.

Change a setting by editing the baseline in a pull request, then apply it.
A change made in the GitHub settings first shows up as drift. The fixtures in
[`tests/test_setup_repository.py`](../tests/test_setup_repository.py) run
offline, through the static-check runner's `python-tests` check, against API responses recorded from this
repository on 2026-11-02 and copies of them changed to show drift.

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
condition is evaluated only in the environments that can enter its branch:
inside the enclosing active branch and, for an `#ElseIf`, before any branch of
its block was taken. A `#Const` defined inside a branch can therefore be used
within that branch. A symbol undefined in an environment that evaluates it
fails, although VBA would read it as `Empty`: a
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

## Python versions

CI runs one exact Python, `3.12.16`. Each workflow that sets up Python defines
it once, as `PYTHON_VERSION` in its top-level `env`, and every
`actions/setup-python` step reads `${{ env.PYTHON_VERSION }}`:
[`static-checks.yml`](../.github/workflows/static-checks.yml),
[`accuracy-gate.yml`](../.github/workflows/accuracy-gate.yml),
[`release.yml`](../.github/workflows/release.yml),
[`issue-triage.yml`](../.github/workflows/issue-triage.yml) and
[`link-check.yml`](../.github/workflows/link-check.yml). To move to a new
patch release, change all five together; first confirm that
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

## Markdown links

`check_links.py` reads every tracked `*.md` file. Offline, on every change
through the `markdown-links` check, it fails when:

- a relative link (Markdown, or an HTML `href`/`src` attribute) names a missing
  or empty file, a path outside the repository or a root-relative path; paths are
  URL-decoded first, so `my%20file.md` names `my file.md`, and character
  references are decoded, so `notes&amp;examples.md` names `notes&examples.md`,
  and backslash escapes are honoured, so `notes\).md` names `notes).md`;
- a fragment names no anchor of the target Markdown file. Anchors are the
  heading slugs GitHub generates from the rendered heading (lower case;
  punctuation and emphasis delimiters dropped, so `__API__` gives `api` while
  `snake_case` keeps its underscore; emoji dropped but their spaces kept;
  duplicates suffixed `-1`, `-2`, ...) and explicit `<a id>`/`<a name>` anchors.

The checker approximates GitHub's Markdown rendering; it is not a full
CommonMark parser. Where the rendering is ambiguous it accepts rather than
reports: a heading's slug with its `_` emphasis left in place is accepted
too, and lines that may be indented code are skipped.

Links inside fenced code blocks, indented code blocks (following CommonMark,
relative to the enclosing list item) and code spans, including spans that
continue over line breaks, are ignored. External links, including this
repository's wiki and issues, are not requested offline.

The [Link check](../.github/workflows/link-check.yml) workflow runs
`check_links.py --external` weekly and on demand. It is not a required check
and never runs on pull requests. Each external link is requested once, with
timeouts and two retries. A wiki page is requested as its raw Markdown, which
answers 404 when the page is missing. HTTP 403, 429 and 5xx answers and network
failures are warnings, because rate limits, bot protection and outages say
nothing about the link; any other non-2xx answer fails the run. Links that are
deliberately not requested are listed, with their reason, in `UNCHECKED` in
`check_links.py`: currently the code-scanning alert pages, which only signed-in
users with write access can open.

Fixtures for every rule are in
[`tests/test_check_links.py`](../tests/test_check_links.py). To check the
external links locally:

```bash
python3 tools/check_links.py --external --repository danielep71/VBA-PROBABILITY-DISTRIBUTIONS
```

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
else; the repository settings baseline is a separate file,
[`docs/repo_governance/repository-settings.json`](../docs/repo_governance/repository-settings.json).
This repository has **not** adopted the template contract, a template
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
| `check_links.py` | `tools/check_links.py`: offline relative-link and GitHub-slug anchor check, and the `--external` request with warnings for 403, 429 and 5xx. This repository is public, so its own wiki pages (as raw Markdown) and issues are requested; only the pages in `UNCHECKED` are skipped. |
| `tests/test_check_links.py` | `tests/test_check_links.py`, with fixtures for percent-encoded paths, HTML attributes, duplicate and punctuated headings, wiki-page requests and `UNCHECKED`. |
| `.github/workflows/link-check.yml` | `.github/workflows/link-check.yml`: the weekly and manual external-link run, with this repository's action pins and Python version. |
| `tests/test_pr_milestones_workflow.py` | Offline tests of that workflow: a dependency-free structure check with negative fixtures in place of the template's YAML-parser invariants, and the template's script tests extended to the job summary and the read-only-token path. |
| `tools/setup_repository.py` | The template's settings script: preview/apply/check modes, ruleset comparison and read-back; reworked for this repository's baseline (merge methods, features, two required checks), a `--verify` mode and a JSON evidence record |
| `tests/test_setup_repository.py` | The template's fake-API tests, rebuilt on API responses recorded from this repository |
| `docs/repo_governance/repository-settings.json` | The template's settings configuration format, filled with this repository's live settings |
| `docs/repo_governance/DISTRIBUTION.md` | The template's distribution review, adapted to a public repository and given a pinned Gitleaks history scan |
| `vba_lexer.py` | The shared lexer of `tools/vba_lexer.py`, without the facade-prefix and house-style lookups. |
| `check_vba_source.py` | `tools/check_vba_source.py`: this repository's flat `src/` and `tests/` and nested `benchmark/` homes and `M_STATS_PROBDIST_` names in place of role homes and prefixes; `Option Private Module` required in a named set of modules and forbidden beside `K_STATS_` worksheet functions; no form or workbook rules. |
| `check_vba_conditionals.py` | `tools/check_vba_conditionals.py`, extended to module `#Const` symbols and to the `Win16` and `Mac` constants, with one directive walk shared by the reachable-source model and the `PtrSafe` rule. |
| `check_vba_visibility.py` | The visibility and collision parts of `tools/check_vba_public_api.py`, without its inventory, signature and public-variable rules; multi-identifier `Const` and variable statements are read name by name, and collisions are evaluated per collision group. |
| `tests/test_check_vba.py` | The fixture repositories and the source, conditional, public-API and lexer cases of `tests/test_check_vba.py`, rewritten for this repository's modules and rules. |
| `requirements-lint-ci.txt`, `mypy.ini` | mypy 2.4.1 and its dependency closure from `tools/requirements-quality-ci.txt`, with the hash of every published file, and the strict settings of `pyproject.toml` `[tool.mypy]`, scoped to `tools/` and `.github/scripts/`. |
| `probe_actionlint.py`, `tests/test_quality_gates.py` | The actionlint unknown-key and mypy strict-mode probes of `tests/test_quality_gates.py`; the actionlint probe is a runner check that reuses the actionlint check's own command, with stand-in commands as offline fixtures. |
| `.github/workflows/static-checks.yml` (minimum-python job) | The `minimum-python` job of `.github/workflows/repository-integrity.yml`, running this repository's runner without installed tools. |
