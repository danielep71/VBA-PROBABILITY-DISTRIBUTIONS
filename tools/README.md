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
Ruff's lock selects the Ubuntu x86-64 CI wheel; it is not a portable developer
lock. actionlint is separately version- and checksum-pinned in the static
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
