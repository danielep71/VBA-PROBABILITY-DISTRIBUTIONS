# Accuracy summary

Generated 2026-07-18 by `compute_errors.py` from `probability_accuracy_grid.csv`.

Reference values are mpmath at 50 digits (see `generate_reference_values.py`). Observed values are produced by the VBA library via `M_STATS_PROBDIST_ACCURACY_EXPORT.bas`. Each function is checked against the accuracy claim published in its source comment.

| Function | Claim | Metric | Max error | At input | Points | Verdict |
|---|---|---|---:|---|---:|---|
| ChiSquare_Cumulative | rel<=2.6E-10 | rel | — | not measured | 0/25 | ⏳ pending |
| ChiSquare_InverseCumulative | rel<=4.7E-12 | rel | — | not measured | 0/25 | ⏳ pending |
| ChiSquare_Survival | rel<=2.6E-10 | rel | — | not measured | 0/25 | ⏳ pending |
| F_Cumulative | rel<=1.1E-10 | rel | — | not measured | 0/16 | ⏳ pending |
| F_InverseCumulative | rel<=5.9E-13 | rel | — | not measured | 0/12 | ⏳ pending |
| F_Survival | rel<=1.1E-10 | rel | — | not measured | 0/16 | ⏳ pending |
| LogChoose | rel<=3.2E-16 | rel | — | not measured | 0/30 | ⏳ pending |
| LogGamma | rel<6.1E-14 | rel | — | not measured | 0/40 | ⏳ pending |
| LogGammaHalfDiff | rel<=2.1E-15 | rel | — | not measured | 0/30 | ⏳ pending |
| StirlingError | abs<=3E-17 | abs | — | not measured | 0/12 | ⏳ pending |
| StudentT_Cumulative | rel<=1.3E-12 | rel | — | not measured | 0/25 | ⏳ pending |
| StudentT_Density | rel<=8.4E-15 | rel | — | not measured | 0/25 | ⏳ pending |
| StudentT_InverseCumulative | rel<=3.0E-12 | rel | — | not measured | 0/25 | ⏳ pending |
| StudentT_Survival | rel<=1.3E-12 | rel | — | not measured | 0/25 | ⏳ pending |

> **No observed values present yet.** Run the export macro in Excel to fill the `observed_vba` column, then re-run `compute_errors.py`.
