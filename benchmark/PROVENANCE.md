# Accuracy evidence provenance (P1-03)

A green accuracy summary must prove more than "the committed observations still
pass every contract." It must prove:

> these exact observations were produced by this exact source, and every active
> row was evaluated.

The observations in `probability_accuracy_grid.csv` are produced by VBA (the
kernels in `src/`, plus the export and study macros) running in Excel. Nothing
in the grid recorded *which* source produced them, so an algorithm could change
in `src/**` while the committed observations — and the green summary — stayed
put, and the hosted accuracy workflow did not even re-run.

## Two evidence bindings

The main grid and independent holdout are separate Excel exports and therefore
have separate provenance records:

- `observation_manifest.json` binds the main accuracy grid to every production,
  test, exporter, and study `.bas` module that can affect or populate it;
- `holdout/holdout_manifest.json` binds the independent holdout to the six
  production modules and `holdout/M_STATS_PROBDIST_HOLDOUT.bas`, plus the exact
  holdout bytes, row count, schema, registry, source commit, export timestamp,
  and Excel environment.

Both use SHA-256 over LF-normalized content. A CRLF checkout and an LF checkout
therefore verify identically, while any substantive byte change fails closed.

The holdout record is intentionally absent during Phase 0 of v1.0.0. The 559
committed observations were exported at older source commit `4553afa`; creating
a current-source manifest for them would be false provenance. The first
`holdout_manifest.json` will be written immediately after the #23 real-Excel
export. Until then, the strict gate reaches `STALE HOLDOUT EVIDENCE` as soon as
the earlier stale main-grid binding has been refreshed.

## What is enforced now (in-repo, hosted, reproducible)

- **Source binding.** `observation_manifest.json` records a content hash
  (SHA-256 over LF-normalized bytes, so CRLF and LF hash identically) of every
  `.bas` file — kernels, exporters, study macros, and tests — the full
  `probability_accuracy_grid.csv` **observation bytes**, the
  `accuracy_contracts.csv` **registry**, and the grid schema version and
  columns. `compute_errors.py` recomputes those against the checked-out tree
  **before** evaluating any contract and **fails the gate** (`STALE EVIDENCE`,
  nonzero exit) if any module changed / was added / removed, if a single
  `observed_vba` value was edited (even with unchanged columns), if a threshold
  in the registry changed, or if the schema drifted. It also fails if the
  manifest is absent or predates content binding (unless
  `--allow-missing-manifest`, for local development only). The binding is thus
  *these exact observation bytes, this exact source, this exact contract
  registry* — not merely a matching structure.
- **Independent-holdout binding.** After the main binding verifies,
  `compute_errors.py` requires `holdout/holdout_manifest.json` and checks the
  production modules, dedicated exporter, exact holdout bytes and row count,
  schema, and contract registry. Missing, malformed, stale, added, removed, or
  changed inputs block before the holdout analyzer can contribute a release
  verdict. Main provenance is checked first so a known stale-main state retains
  its exact diagnostic rather than being masked by a downstream failure.
- **The gate now runs on source changes.** `accuracy-gate.yml` triggers on
  `src/**` and `tests/**` (as well as `benchmark/**`). A source edit therefore
  re-runs the gate, which then fails on the manifest mismatch until the
  observations are re-exported and the manifest re-written.
- **Provenance in the summary.** `accuracy_summary.md` opens with the bound
  source commit, export timestamp, Excel version/build/bitness, module count,
  and schema version.
- **Every active row evaluated.** Enforced separately by the P1-02 preflight
  (`_contract_eval.dispositions` + the `measured == to_measure` invariant).

## Operating procedure (run at every export)

Whenever `src/**`, the exporters, the tests, either observation grid, or the
contract registry changes, the affected manifest must be re-written — the gate
will not certify the summary otherwise:

1. Import the current source into the workbook and re-export the observations
   (`Export_Accuracy_Observations`, plus any affected study macro).
2. Re-run the manifest writer against the freshly exported grid, recording the
   environment that produced it:

   ```
   python write_manifest.py --commit-sha <sha> \
       --excel-version <ver> --excel-build <build> --office-bitness <32|64>
   ```

3. Commit the grid **and** `observation_manifest.json` together.

For a fresh independent-holdout export, write its binding in the same evidence
operation:

```
python write_manifest.py --holdout --commit-sha <sha> \
    --excel-version <ver> --excel-build <build> --office-bitness <32|64>
```

Or, after both grids have just been exported and
`benchmark/excel_environment.json` is current, run:

```
python refresh_evidence.py --bind-exported-holdout
```

That explicit flag writes both the normal main binding and the holdout binding
before regenerating summaries. Plain `refresh_evidence.py` never creates a
holdout binding, so regenerating documentation cannot accidentally claim that
historical observations came from current source.

The baseline manifest was written against the current committed observations
under the assumption that they are current (they were re-exported through the
P1-01, F-envelope, and beta_f_inverse work). Regenerate it at the next full
export for a from-scratch binding, and populate the environment fields, which
are currently `unrecorded`.

## Still owned by the maintainer (outside the repo)

- **Branch protection.** Require *both* the `Accuracy Gate` and the Excel/VBA
  regression checks on `main`. The Excel regression exercises current source at
  broad behavioral tolerances; the accuracy gate certifies the tight external
  contracts against source-bound evidence. Neither substitutes for the other.
- **Environment capture.** Have each export path record Excel version/build and
  Office bitness (either from the exporter macro or passed to
  `write_manifest.py`) so the `unrecorded` fields become real.

## Strongest target design (two-stage, when automation is practical)

- **Stage 1 — self-hosted Windows/Excel:** import exact current source, export
  observations, write the manifest (including the source commit and environment),
  upload both as a build artifact.
- **Stage 2 — hosted Python:** download the artifact, verify the manifest
  against the checked-out source, evaluate every contract, publish the summary.

This removes the manual re-export step and makes the binding automatic: the only
observations the gate ever sees are ones Stage 1 just produced from that revision.


## Writing a manifest requires a fresh export

A manifest asserts that the committed observations were produced by the
checked-out source. Only a real Excel export makes that true.

`write_manifest.py` refuses a bare invocation and exits non-zero. Writing
requires `--from-fresh-export`; `--dry-run` previews without touching
anything. `refresh_evidence.py` no longer binds during ordinary
regeneration - it refreshes summaries only. Binding requires
`--bind-exported-main` or `--bind-exported-holdout`, each of which passes
`--from-fresh-export` to the writer explicitly.

This exists because of `9fba175`. A bare `write_manifest.py` rebound the
main manifest to source three commits newer than the observations. The
seven-line signature was unchanged, but the strict gate's failure had moved
from `STALE EVIDENCE` (two mismatches, truthful) to `STALE HOLDOUT EVIDENCE`,
and the main binding read clean - one `write_manifest.py --holdout` away
from a green gate on stale evidence. Restored in `c496f1b`.

### The CI guard

`check_manifest_provenance.py` examines each commit **separately**, never a
push's aggregate diff: otherwise one commit touching the grid conceals
another commit's rebind. A commit modifying a manifest is legal only if, in
that same commit, one of these holds:

| | Condition |
| --- | --- |
| A | the manifest's own grid is also modified - what a real export looks like |
| B | `benchmark/excel_regression_record.json` is also modified **and validates**: it must identify that grid as an exported target and bind the export session, source identity, Excel version/build/bitness, regression totals, and each grid's SHA-256 and row count, matching the committed grid |
| C | the commit modifies the manifest and nothing else, and the result is byte-identical to an earlier committed version - the repair path, which cannot launder a rebind because a rebind produces content that has never existed |

B exists so that a re-export producing a byte-identical grid - a `.bas`
change altering no observation value - remains legal. The record does not
exist yet; it is produced by Phase 1's first export session, so until then
only A and C can apply, which is correct for a phase in which no export can
happen.
