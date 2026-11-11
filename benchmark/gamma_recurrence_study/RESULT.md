# Recurrence reuse: measurement blocked (#307)

No performance conclusion or production promotion is supported yet.

| Run | Candidate | Completed samples | Outcome |
|---|---|---:|---|
| [43030484382](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/actions/runs/43030484382) | 2c82328753b413f9da8c5a6a6a281914051c464c | 193 / 216 | Excel RPC failure during round 3; cleanup also failed |
| [43036251207](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/actions/runs/43036251207) | ab54518473bd14930bf277b532ac047d115de717 | 0 / 216 | Excel COM activation failed before any measurement |

All 193 completed samples in the first attempt preserve baseline output bits.
Its four generated source files match the artifact metadata and local
regeneration byte for byte. Both downloaded artifact ZIP digests were verified.
These checks preserve traceability; they do not make an incomplete run valid.
All partial samples and both errors are retained in `results/`.

The [frozen design](DESIGN.md) and the [host recovery record](HOST_RECOVERY.md)
describe the candidate operations, full sample matrix and repeat procedure.
The repeat keeps numerical transformations, cases, ordering, counts and the
5% target unchanged, but bounds each Excel instance to one round. It failed
before this mitigation could be exercised. The original host failure's root
cause remains undetermined.

Production source, routing, thresholds and accuracy manifests are unchanged.
The untouched routing holdout was not read or evaluated. Restore Excel COM
startup on the runner, repeat the entire study, then assess bit identity and
cost. Do not splice partial measurements or select candidates from this report.
