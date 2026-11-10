# Fresh public Gamma main/holdout export — expanded #242 suite

[Excel VBA Regression run 42529642623](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/actions/runs/42529642623)
executed candidate `e91c535a8f7ef2ba3830d9c690b39966b74cc6b3` on Excel 16.0
build 20430 / 64-bit. The regression passed **14,878 / 14,878 assertions**;
import, execution-backed compile, regression and cleanup passed. The same job
then exported the main grid in `full` mode and the independent holdout.

| Export | Total rows | Freshly filled | Changed observations | Error cells |
| --- | ---: | ---: | ---: | ---: |
| Main | 2,093 | 1,737 | 112 | 0 |
| Holdout | 659 | 659 | 5 | 0 |

The other 356 rows in the main CSV belong to other evidence sets and were
preserved byte-for-byte. They are not claimed as fresh exports. Coordinates,
reference values, regimes, expected-error declarations and contract thresholds
are unchanged. These are re-exports of existing datasets, not new untouched
holdout cohorts. The separately frozen fresh cohort is recorded in #242.

## Accuracy and caller movement

The strict main gate passes **166 / 166 active contracts**, with zero failures,
known limitations or pending evaluations; one additional registry row remains
characterization-only. The independent holdout passes **83 / 83 contracts**.
The 27-row main coverage debt remains unchanged and explicitly registered.

| Public function | Changed main observations | Changed holdout observations |
| --- | ---: | ---: |
| ChiSquare_Cumulative | 23 | 0 |
| ChiSquare_Survival | 21 | 0 |
| ChiSquare_InverseCumulative | 25 | 0 |
| Gamma_Cumulative | 17 | 0 |
| Gamma_Survival | 12 | 0 |
| Gamma_InverseCumulative | 2 | 0 |
| Poisson_Cumulative | 4 | 2 |
| Poisson_Survival | 8 | 3 |

All other observation cells are unchanged, including Poisson inverses and all
non-Gamma-family functions. The recorded per-row comparison is
[`public_export_movement.json`](public_export_movement.json), using the existing high/low
observation encoding and fixed decimal references at 80-digit Decimal precision.

The large/extreme-shape Gamma CDF worst errors fall from approximately
`2.90e-14` / `2.97e-13` to `8.37e-17` / `4.23e-17`; the corresponding survival
errors fall from `2.59e-14` / `5.59e-14` to `1.20e-16` / `1.25e-16`.
Poisson CDF/SF worst main errors fall from `2.56e-14` / `2.91e-14` to
`8.40e-17` / `9.32e-17`.

Not every changed row improves: 22 main-grid rows have a larger relative
error than before, all at or below `1.136e-15` and within their unchanged
contracts. Sixteen are chi-square inverses, one is a Gamma inverse, and five
are forward/survival outputs. Gamma inverse's two-row worst error increases
from `2.57e-16` to `8.35e-16`, still below its `2e-14` contract; chi-square
inverse's overall worst improves from `1.31e-15` to `1.14e-15`.
All five changed holdout rows improve. No threshold was loosened and no
unexpected movement was hidden by changing coordinates or references.

## Retention and provenance

| Item | Identity |
| --- | --- |
| Export artifact | `13637073116` |
| Original ZIP SHA-256 | `582cd4efbca45e6e8fff287d815a9370aca39e93b35fbb9024a5e4c71386d15b` |
| Main raw CSV SHA-256 | `3e94db6a501c4ac44f1e57e2beeab43c9f80c21e084b7030b100d5bb68c7aedc` |
| Holdout raw CSV SHA-256 | `bc7cdb7b2906e2e122e6e5fba2128b28a470fcd365691e10650ab7956397075f` |

[`excel_public_export_artifact.json`](excel_public_export_artifact.json) retains the original
ZIP bytes losslessly, including the export manifest/log, raw CSVs, both
environment records and original regression certificate/log. The exported
candidate SHA is preserved. `apply_excel_export.py --grid main --grid holdout`
accepted the artifact from a clean checkout of that exact candidate, verified
source and raw-file hashes and permitted cell changes, finalized the canonical
record, and bound the two actual exports. The grids, manifests, environment,
certification record, generated summaries and README are committed atomically
as a direct child of the exported candidate.

The normal-tail study binds the whole main CSV, so the legitimate re-export
made two of its tooling checks stale. Its documented
`validate_kernel.py --regenerate --reason ...` path was run against the new
grid. It reported PASS and appended the reason and changed identity. The
coefficients, all scientific result sections, predictions, A6 and original
holdout exposure remain exactly equal to the preceding record; the original
holdout was not evaluated again. The two affected tooling checks pass after
this regeneration. The retained Excel record's pre-commit digest check is
validated against the complete evidence snapshot before commit and against
the actual evidence commit afterward.

Observed export macro times were 217.59 seconds (main) and 7.64 seconds
(holdout); the export adapter, including import and cleanup, took 267.42
seconds. These are observed workload times, not a full-domain performance
guarantee.

## Acceptance boundary

This export removes stale main/holdout evidence for #242's own expanded test
module and 14,878-assertion policy. It is a separate source-specific execution
from #241's 1,947-assertion export. The two original raw main CSVs and the two
raw holdout CSVs are byte-identical, and the six production Git blobs also
match; no #241 certificate or manifest is reused here.

Both PRs remain open and draft. This evidence does not establish a global
accumulated-error bound, full-domain performance, another Office host, or
release acceptance. The 27-row coverage debt and other registered release
blockers remain. No issue, PR, release or milestone is closed by this update.
