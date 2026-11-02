# Repository static gates

Generic, Excel-free checks of the repository itself. They complement the
numerical evidence machinery under [`benchmark/`](../benchmark/) and the Excel
regression under [`ci/`](../ci/); they never prove compilation, regression or
accuracy.

All three run in [`.github/workflows/static-checks.yml`](../.github/workflows/static-checks.yml)
on every push to `main`, every pull request and every `v*` tag. Each is proved
by its own `--self-test` fixtures, which also run before the strict accuracy
gate through [`benchmark/test_evidence_tools.py`](../benchmark/test_evidence_tools.py)
and are registered in [`benchmark/verification_depth.json`](../benchmark/verification_depth.json).

| Tool | Checks |
| --- | --- |
| `check_committed_whitespace.py` | No whitespace errors in the committed candidate range (`git diff --check` semantics, read only) |
| `check_vba_jumps.py` | Every `GoTo`, `GoSub`, `Resume` and `On Error GoTo` target resolves inside its own procedure, across all tracked `.bas`/`.cls`/`.frm` modules |
| `check_release_semantics.py` | `VERSION` is strict SemVer, `CHANGELOG.md` has exactly one `[Unreleased]`, release headings are ordered and dated, and comparison links match the repository |
| `_gatelib.py` | Shared command-line and report helpers for the three checks above |

The same workflow also lints every tracked Python file for correctness with
Ruff, restricted by [`ruff.toml`](../ruff.toml) to syntax errors and pyflakes
rules. CI installs it from the hash-locked
[`requirements-lint-ci.txt`](requirements-lint-ci.txt).

Run locally from the repository root:

```bash
python3 tools/check_committed_whitespace.py --root . --self-test
python3 tools/check_committed_whitespace.py --root . --mode committed --head HEAD
python3 tools/check_vba_jumps.py --root . --self-test
python3 tools/check_vba_jumps.py --root .
python3 tools/check_release_semantics.py --root . --self-test
python3 tools/check_release_semantics.py --root .
python3 -m ruff check .
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
offline, in the static workflow, against API responses recorded from this
repository on 2026-11-02 and copies of them changed to show drift.

## Dependency updates

Dependabot checks the GitHub Actions pins, `benchmark/requirements-ci.txt`
(mpmath and SymPy), and `tools/requirements-lint-ci.txt` (Ruff) weekly. Updates
are assigned to the maintainer with priority P3 for initial triage; raise the
priority when a security advisory or other material risk requires it. Review
and merge remain manual. Set the applicable milestone during PR triage.

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
else; the repository settings baseline is a separate file,
[`docs/repo_governance/repository-settings.json`](../docs/repo_governance/repository-settings.json).
This repository has **not** adopted the template contract, a template
profile or the template's directory layout, and the file must not be read as
claiming any of them. The template's release-history policy applies only to the
template repository itself and is not imported.

### Adapted from `danielep71/MyTemplate` at `9ff7670`

These files are adapted, not copied: names, layout and scope follow this
repository, and template-only features are dropped.

| File | What was adapted |
| --- | --- |
| `tools/setup_repository.py` | The template's settings script: preview/apply/check modes, ruleset comparison and read-back; reworked for this repository's baseline (merge methods, features, two required checks), a `--verify` mode and a JSON evidence record |
| `tests/test_setup_repository.py` | The template's fake-API tests, rebuilt on API responses recorded from this repository |
| `docs/repo_governance/repository-settings.json` | The template's settings configuration format, filled with this repository's live settings |
| `docs/repo_governance/DISTRIBUTION.md` | The template's distribution review, adapted to a public repository and given a pinned Gitleaks history scan |
