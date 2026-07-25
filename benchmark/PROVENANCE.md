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

Whenever `src/**`, the exporters, the tests, the observation grid, or the
contract registry change, the manifest must be re-written — the gate will not
certify the summary otherwise:

1. Import the current source into the workbook and re-export the observations
   (`Export_Accuracy_Observations`, plus any affected study macro).
2. Re-run the manifest writer against the freshly exported grid, recording the
   environment that produced it:

   ```
   python write_manifest.py --commit-sha <sha> \
       --excel-version <ver> --excel-build <build> --office-bitness <32|64>
   ```

3. Commit the grid **and** `observation_manifest.json` together.

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
- **Environment capture.** Have the export path record Excel version/build and
  Office bitness (either from the exporter macro or passed to `write_manifest.py`)
  so the `unrecorded` fields become real.

## Strongest target design (two-stage, when automation is practical)

- **Stage 1 — self-hosted Windows/Excel:** import exact current source, export
  observations, write the manifest (including the source commit and environment),
  upload both as a build artifact.
- **Stage 2 — hosted Python:** download the artifact, verify the manifest
  against the checked-out source, evaluate every contract, publish the summary.

This removes the manual re-export step and makes the binding automatic: the only
observations the gate ever sees are ones Stage 1 just produced from that revision.
