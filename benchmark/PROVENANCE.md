# Accuracy evidence provenance

A green accuracy summary must prove more than "the committed observations still
pass every contract." It must prove:

> these exact observations were produced by this exact source, under a recorded
> Excel environment, and every active row was evaluated.

The numerical observations are produced by VBA running in Excel. The repository
therefore keeps **two complementary evidence layers**:

1. a canonical Excel-host certification record for the exact source candidate
   that Excel executed; and
2. source/grid manifests for the main accuracy grid and the independent holdout.

Neither layer substitutes for the other.

## Exact-SHA Excel certification

`.github/excel-evidence-policy.json` defines the regression entry point, exact
assertion count, imported source inventory, and supported grid exporters. The
self-hosted Excel workflow emits `excel-certification.json`; the contract is
validated by `excel_certification.py` / `check_excel_certification.py`.

A green regression record binds:

- a full 40-character candidate Git SHA;
- SHA-256 of the canonical Git bytes of every imported VBA source;
- Excel version/build and actual Office bitness;
- runner/workflow identity and timestamps;
- import, execution-backed compile, regression, and cleanup stages;
- the exact policy assertion count and pass/fail totals; and
- SHA-256 of the retained regression log.

The normal regression workflow does **not** export numerical grids. Its main and
holdout entries therefore remain `exported: false`. A green regression run by
itself cannot authorize a provenance rebind.

See [`../docs/EXCEL_CERTIFICATION.md`](../docs/EXCEL_CERTIFICATION.md) for the
record schema, validation command, and fresh-export finalization procedure.

## Two numerical evidence bindings

The main grid and independent holdout are separate Excel exports and therefore
have separate provenance manifests:

- `observation_manifest.json` binds the main accuracy grid to every production,
  test, exporter, and study `.bas` module that can affect or populate it;
- `holdout/holdout_manifest.json` binds the independent holdout to the six
  production modules and `holdout/M_STATS_PROBDIST_HOLDOUT.bas`, plus the exact
  holdout bytes, row count, schema, registry, source commit, export timestamp,
  and Excel environment.

The main manifest hashes the LF-normalized source content used by its existing
manifest contract. The canonical Excel certification record is stricter about
source identity: it hashes the exact canonical Git blob bytes. Both approaches
are stable across ordinary CRLF/LF working-tree checkout differences.

**Evidence snapshot — 2026-09-30, source baseline `d2e1592`.** The latest retained Excel runtime result is 909/909 PASS on candidate `74041b3` (Excel 16.0 build 20228, 64-bit); the certified VBA bytes remain unchanged. This is not a new Excel run on the documentation commit. Main and holdout manifests are source-bound, but release export certification is incomplete: [#47](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/issues/47) tracks the retained main-grid digest mismatch and [#29](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/issues/29) tracks holdout export certification. The numerical gate is green with 36 unclaimed main-grid rows tracked by [#22](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/issues/22). No stable release has been published; [#31](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/issues/31) tracks readiness.

## What the hosted gate enforces

- **Main source binding.** `compute_errors.py` verifies the main manifest before
  evaluating contracts. Changed, added, removed, malformed, or unbound source,
  grid, registry, or schema evidence fails closed as `STALE EVIDENCE`.
- **Independent-holdout binding.** After the main binding verifies,
  `compute_errors.py` requires a valid holdout manifest. Missing or stale
  holdout evidence blocks before the holdout analyzer can contribute a release
  verdict.
- **Source changes trigger the gate.** `accuracy-gate.yml` runs on production,
  test, and benchmark changes.
- **Every active row is evaluated.** This remains separately enforced by the
  evaluation preflight and measured-row invariant.
- **Manifest changes are checked per commit.** The guard never relies on an
  aggregate push diff; a grid change in one commit cannot launder a manifest
  rebind in another.

The Accuracy Gate fetches the full commit graph because exact-restoration
validation and its historical regression controls are inherently
history-dependent. A depth-1 checkout is not sufficient evidence for those
rules.

## Fresh-export operating procedure

Commands below run from the repository root. The current #47 blocker must be resolved before final release certification; these steps describe the required evidence chain, not a claim that the retained export certificate currently passes.

### 1. Certify the exact source candidate in Excel

Check out and record the full candidate SHA. Import its exact production and test modules into a clean workbook, run the self-hosted regression workflow, and retain its green `excel-certification.json` and regression log. Validate the artifact as described in [EXCEL_CERTIFICATION.md](../docs/EXCEL_CERTIFICATION.md).

This establishes runtime/source identity. It does not establish a fresh numerical export.

### 2. Export the selected numerical evidence in Excel

From the same exact source candidate, run the appropriate exporter:

- main: `Export_Accuracy_Observations`;
- holdout: the dedicated holdout exporter.

Also run `Export_ExcelEnvironment` so `benchmark/excel_environment.json` records
the Excel version/build/bitness used for the export.

Do not claim a grid that was not actually exported in this session.

### 3. Finalize the certification record

Immediately after the real export, from the unchanged exact candidate checkout,
run only the flags corresponding to grids actually exported:

```bash
python benchmark/finalize_excel_certification.py \
  --record <artifact-dir>/excel-certification.json \
  --main \
  --holdout \
  --from-fresh-export
```

`--main` and `--holdout` are independent. The finalizer fails unless:

- the record candidate equals full current `HEAD`;
- the policy-selected regression source bytes still match that candidate;
- `excel_environment.json` agrees with the regression record on Excel
  version/build/bitness;
- the relevant exporter module is included in the certified source inventory;
- the selected grid exists and its SHA-256 and row count can be bound; and
- the completed record structure and source inventory revalidate.

The finalizer hashes working-tree grids; it does not yet check the committed grid bytes. The evidence commit must pass the canonical validator afterward, including grid digest and row-count checks. See #47 and the commands in [EXCEL_CERTIFICATION.md](../docs/EXCEL_CERTIFICATION.md).

It writes `benchmark/excel_regression_record.json`. The tool cannot observe an
Excel export itself: `--from-fresh-export` is an explicit operator assertion,
which is why this step must immediately follow the real Excel export.

### 4. Bind only the grids actually exported

For example, after both grids were exported:

```bash
python benchmark/refresh_evidence.py \
  --bind-exported-main \
  --bind-exported-holdout
```

Use only `--bind-exported-main` or only `--bind-exported-holdout` when only that
grid was refreshed. Plain `refresh_evidence.py` regenerates derived material but
never asserts a fresh export.

### 5. Commit the evidence atomically

Commit the changed grid(s), corresponding manifest(s), finalized canonical Excel
record, environment record, and regenerated summaries together when those files
belong to the same evidence event.

A byte-identical re-export is also valid. In that case the grid file may not
appear changed in Git, so the manifest guard uses the finalized canonical Excel
record to prove the fresh export rather than relying on a grid diff that cannot
exist.

## Manifest provenance guard

`check_manifest_provenance.py` examines every commit in scope separately. A
commit modifying a manifest is legal only if, in that same commit, one of these
conditions holds:

| | Condition |
| --- | --- |
| **A** | The manifest's own grid is also modified. This is the ordinary fresh-export shape. |
| **B** | `benchmark/excel_regression_record.json` is also modified and the canonical validator proves a green exact-SHA run **plus an explicit fresh-export claim for that exact grid**, whose committed SHA-256 and row count match. |
| **C** | The commit modifies only the manifest and restores bytes identical to an earlier committed version. This is the exact-restoration repair path. |

B permits a genuinely fresh export that produces a byte-identical grid. It requires a valid canonical export record. The retained record currently fails the main-grid digest check (#47) and does not claim a holdout export (#29), so it cannot currently authorize this exception.


## Why binding is verified, not declared

`write_manifest.py` refuses a bare invocation, and `--from-fresh-export` is
checked rather than believed: a byte-identical grid alone does not establish
that an export occurred, so the writer requires independent certification. The documented exception is the one
`check_manifest_provenance.py` already recognises - a real export that
reproduces a byte-identical grid, evidenced by a validated
`excel_regression_record.json` naming that grid as exported.

Both interlocks exist because the flag alone proved insufficient. It was passed
once with no export behind it, rebinding the main manifest to a newer commit
while the observations were unchanged: the strict gate's failure moved from
`STALE EVIDENCE` to `STALE HOLDOUT EVIDENCE` and the main binding read clean.
Reverted before it was committed. A promise is cheaper to make than an export,
so the promise is now checked.

## Canonical Excel certification record

The manifests bind committed observation bytes to source. Separately, `excel-certification.json` / retained `benchmark/excel_regression_record.json` bind an Excel runtime session to one full candidate SHA and exact imported VBA bytes. A regression-only record has every grid entry `exported: false` and cannot authorize a manifest rebind. Only a record finalized from a real fresh export, with the relevant grid marked exported and matching SHA-256/row count, may satisfy the byte-identical-grid exception in `check_manifest_provenance.py`. See [EXCEL_CERTIFICATION.md](../docs/EXCEL_CERTIFICATION.md).
