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

The holdout record is intentionally absent during the current pre-export phase
of v1.0.0. The committed holdout observations came from older source. Creating a
current-source holdout manifest before a real re-export would be false
provenance.

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

A numerical provenance binding may be refreshed only after the relevant grid
was really exported from the exact source candidate.

### 1. Obtain green exact-SHA regression evidence

Run the self-hosted Excel regression on the exact candidate. Download
`excel-certification.json` and `test-result.txt`, then validate them against the
candidate checkout as documented in `docs/EXCEL_CERTIFICATION.md`.

This establishes the runtime/source identity. It does **not** yet establish a
fresh numerical export.

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
- the completed canonical record revalidates.

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

Condition B exists for legitimate byte-identical re-exports. A green
regression-only record is rejected because it has `exported: false` and carries
no grid hash/row-count claim.

This guard was introduced after historical commit `9fba175` rebound the main
manifest to newer source without re-exporting the observations. The truthful
manifest was restored in `c496f1b`. Those historical cases remain regression
controls, together with synthetic negative fixtures for wrong source hashes,
wrong target, wrong grid digest/row count, missing fields, `exported: false`,
and a two-commit push.

## Evidence-only follow-up commits

The canonical record's `candidate_sha` identifies the source revision actually
executed in Excel. A later evidence-only commit may retain that record only while
every certified VBA source byte is unchanged. The validator checks the evidence
commit against the certified source inventory. Any later change to a certified
VBA source invalidates the retained record and requires a new Excel run.

This avoids the self-referential requirement that an evidence JSON file somehow
be present in the source commit that was already executed.

## Maintainer controls outside portable Python

- Require both the Accuracy Gate and Excel/VBA Regression checks on protected
  release branches.
- Keep the Excel runner isolated and interactive as documented in
  `docs/EXCEL_VBA_CI.md`.
- Do not infer a fresh export from a passing regression, unchanged grid, or
  unchanged accuracy summary.
- Preserve failed/non-green runtime evidence when useful for diagnosis; never
  translate `NOT_RUN`, `FAIL`, `TIMEOUT`, or cleanup failure into PASS.

## Target automation boundary

The long-term automation model remains two-stage:

1. an eligible Windows/Excel host executes the exact candidate and, when
   requested, exports numerical observations and produces canonical runtime
   evidence; then
2. hosted Python verifies the source/grid binding, evaluates every contract, and
   regenerates the release summaries.

The trust boundary is deliberate: portable Python verifies what Excel produced;
it does not pretend to execute Excel itself.
