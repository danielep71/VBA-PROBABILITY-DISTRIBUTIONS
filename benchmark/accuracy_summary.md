# Accuracy summary

Generated 2026-07-18 by `compute_errors.py` from `probability_accuracy_grid.csv`.

Reference values are mpmath at 50 digits (see `generate_reference_values.py`). Observed values are produced by the VBA library via `M_STATS_PROBDIST_ACCURACY_EXPORT.bas`. Each function is checked against the accuracy claim published in its source comment.

| Function | Claim | Metric | Max error | At input | Points | Verdict |
|---|---|---|---:|---|---:|---|
| Beta_Cumulative | rel<=2E-14 | rel | 5.43e-15 | `0.8, 5.0, 1.0` | 3/3 | ✅ PASS |
| Beta_Density | rel<=5E-15 | rel | 2.78e-15 | `0.3, 2.0, 5.0` | 3/3 | ✅ PASS |
| Beta_InverseCumulative | rel<=5E-15 | rel | 6.17e-16 | `0.95, 2.0, 5.0` | 2/2 | ✅ PASS |
| Beta_Mean | rel<=5E-15 | rel | 1.39e-16 | `2.0, 3.0` | 2/2 | ✅ PASS |
| Beta_StdDev | rel<=5E-15 | rel | 3.80e-17 | `5.0, 2.0` | 2/2 | ✅ PASS |
| Beta_Survival | rel<=5E-15 | rel | 4.36e-15 | `0.3, 2.0, 5.0` | 3/3 | ✅ PASS |
| Beta_Variance | rel<=5E-15 | rel | 7.20e-17 | `5.0, 2.0` | 2/2 | ✅ PASS |
| ChiSquare_Cumulative | rel<=2.6E-10 | rel | 4.04e-14 | `0.5, 30.0` | 25/25 | ✅ PASS |
| ChiSquare_InverseCumulative | rel<=4.7E-12 | rel | 2.15e-14 | `0.95, 1.0` | 25/25 | ✅ PASS |
| ChiSquare_Survival | rel<=2.6E-10 | rel | 3.74e-14 | `80.0, 30.0` | 25/25 | ✅ PASS |
| Exponential_Cumulative | rel<=5E-15 | rel | 6.56e-17 | `3.0, 0.5` | 3/3 | ✅ PASS |
| Exponential_Density | rel<=5E-15 | rel | 1.04e-16 | `3.0, 0.5` | 3/3 | ✅ PASS |
| Exponential_InverseCumulative | rel<=5E-15 | rel | 4.08e-16 | `0.95, 2.0` | 2/2 | ✅ PASS |
| Exponential_Survival | rel<=5E-15 | rel | 1.04e-16 | `3.0, 0.5` | 3/3 | ✅ PASS |
| F_Cumulative | rel<=1.1E-10 | rel | 2.59e-14 | `1.0, 10.0, 30.0` | 16/16 | ✅ PASS |
| F_InverseCumulative | rel<=5.9E-13 | rel | 1.65e-14 | `0.5, 10.0, 30.0` | 12/12 | ✅ PASS |
| F_Survival | rel<=1.1E-10 | rel | 2.98e-14 | `1.0, 10.0, 30.0` | 16/16 | ✅ PASS |
| Gamma_Cumulative | rel<=2E-14 | rel | 1.71e-14 | `5.0, 3.0, 2.0` | 3/3 | ✅ PASS |
| Gamma_Density | rel<=2E-14 | rel | 1.61e-14 | `5.0, 3.0, 2.0` | 3/3 | ✅ PASS |
| Gamma_InverseCumulative | rel<=2E-14 | rel | 9.54e-15 | `0.5, 2.0, 1.0` | 2/2 | ✅ PASS |
| Gamma_Mean | rel<=5E-15 | rel | 0.00e+00 | `2.0, 3.0` | 2/2 | ✅ PASS |
| Gamma_StdDev | rel<=5E-15 | rel | 1.36e-16 | `5.0, 2.0` | 2/2 | ✅ PASS |
| Gamma_Survival | rel<=2E-14 | rel | 1.60e-14 | `2.0, 2.0, 1.0` | 3/3 | ✅ PASS |
| Gamma_Variance | rel<=5E-15 | rel | 0.00e+00 | `2.0, 3.0` | 2/2 | ✅ PASS |
| LogChoose | rel<=3.2E-16 | rel | 1.94e-16 | `2.0, 1.0` | 30/30 | ✅ PASS |
| LogGamma | rel<=6.1E-14 | rel | 3.77e-15 | `8.376776400682919` | 40/40 | ✅ PASS |
| LogGammaHalfDiff | rel<=2E-14 | rel | 1.53e-14 | `1.6102620275609393` | 30/30 | ✅ PASS |
| Lognormal_Cumulative | rel<=5E-15 | rel | 1.80e-16 | `0.5, 0.0, 1.0` | 3/3 | ✅ PASS |
| Lognormal_Density | rel<=3E-14 | rel | 2.36e-14 | `3.720075976020836e-44, 0.0, 2.5` | 4/4 | ✅ PASS |
| Lognormal_InverseCumulative | rel<=5E-15 | rel | 1.69e-16 | `0.025, 0.0, 1.0` | 2/2 | ✅ PASS |
| Lognormal_InverseSurvival | rel<=5E-15 | rel | 1.40e-16 | `0.025, 0.0, 1.0` | 2/2 | ✅ PASS |
| Lognormal_Mean | rel<=5E-15 | rel | 4.66e-17 | `0.0, 1.0` | 2/2 | ✅ PASS |
| Lognormal_ParamMeanLog | rel<=5E-15 | rel | 1.68e-17 | `10.0, 3.0` | 2/2 | ✅ PASS |
| Lognormal_ParamStdDevLog | rel<=5E-15 | rel | 1.17e-16 | `10.0, 3.0` | 2/2 | ✅ PASS |
| Lognormal_StdDev | rel<=5E-15 | rel | 8.73e-18 | `0.5, 0.25` | 2/2 | ✅ PASS |
| Lognormal_Survival | rel<=5E-15 | rel | 1.94e-16 | `2.0, 0.5, 0.25` | 3/3 | ✅ PASS |
| Lognormal_Variance | rel<=5E-15 | rel | 3.08e-16 | `0.0, 1.0` | 2/2 | ✅ PASS |
| NormalStandard_Cumulative | rel<=5E-15 | rel | 9.04e-16 | `-2.0` | 6/6 | ✅ PASS |
| NormalStandard_Density | rel<=5E-15 | rel | 2.47e-16 | `-2.0` | 6/6 | ✅ PASS |
| NormalStandard_IntervalProbability | rel<=5E-15 | rel | 1.39e-16 | `-1.96, 1.96` | 3/3 | ✅ PASS |
| NormalStandard_InverseCumulative | rel<=5E-15 | rel | 1.80e-15 | `0.999` | 5/5 | ✅ PASS |
| NormalStandard_InverseCumulativeFast | rel<=5E-9 | rel | 8.05e-10 | `0.975` | 5/5 | ✅ PASS |
| NormalStandard_InverseSurvival | rel<=5E-15 | rel | 1.80e-15 | `0.999` | 5/5 | ✅ PASS |
| NormalStandard_Survival | rel<=2E-14 | rel | 1.52e-14 | `3.0` | 6/6 | ✅ PASS |
| Normal_Cumulative | rel<=5E-15 | rel | 3.22e-17 | `110.0, 100.0, 15.0` | 3/3 | ✅ PASS |
| Normal_Density | rel<=5E-15 | rel | 3.23e-16 | `1.96, 0.0, 1.0` | 3/3 | ✅ PASS |
| Normal_InverseCumulative | rel<=5E-15 | rel | 1.05e-16 | `0.025, 10.0, 2.0` | 2/2 | ✅ PASS |
| Normal_InverseSurvival | rel<=5E-15 | rel | 1.83e-16 | `0.99, 100.0, 15.0` | 2/2 | ✅ PASS |
| Normal_Survival | rel<=5E-15 | rel | 8.30e-16 | `1.96, 0.0, 1.0` | 3/3 | ✅ PASS |
| Normal_ZScore | rel<=5E-15 | rel | 4.00e-19 | `110.0, 100.0, 15.0` | 3/3 | ✅ PASS |
| StirlingError | abs<=3E-17 | abs | 2.52e-17 | `501.0` | 12/12 | ✅ PASS |
| StudentT_Cumulative | rel<=1.3E-12 | rel | 2.94e-14 | `1.0, 1000.0` | 25/25 | ✅ PASS |
| StudentT_Density | rel<=2E-14 | rel | 1.93e-14 | `20.0, 1000.0` | 25/25 | ✅ PASS |
| StudentT_InverseCumulative | rel<=3.0E-12 | rel | 9.96e-14 | `0.95, 1000.0` | 25/25 | ✅ PASS |
| StudentT_Survival | rel<=1.3E-12 | rel | 1.55e-13 | `1.0, 1000.0` | 25/25 | ✅ PASS |
| Uniform_Cumulative | rel<=5E-15 | rel | 0.00e+00 | `3.0, 0.0, 10.0` | 2/2 | ✅ PASS |
| Uniform_Density | rel<=5E-15 | rel | 7.99e-19 | `2.5, 1.0, 4.0` | 2/2 | ✅ PASS |
| Uniform_InverseCumulative | rel<=5E-15 | rel | 0.00e+00 | `0.5, 0.0, 10.0` | 2/2 | ✅ PASS |
| Uniform_Survival | rel<=5E-15 | rel | 0.00e+00 | `3.0, 0.0, 10.0` | 2/2 | ✅ PASS |
| Weibull_Cumulative | rel<=5E-15 | rel | 2.40e-16 | `0.5, 3.0, 1.0` | 3/3 | ✅ PASS |
| Weibull_Density | rel<=5E-15 | rel | 1.73e-16 | `0.5, 3.0, 1.0` | 3/3 | ✅ PASS |
| Weibull_InverseCumulative | rel<=5E-15 | rel | 6.71e-17 | `0.95, 2.0, 2.0` | 2/2 | ✅ PASS |
| Weibull_Mean | rel<=2E-14 | rel | 7.82e-15 | `1.5, 1.0` | 2/2 | ✅ PASS |
| Weibull_StdDev | rel<=5E-15 | rel | 4.12e-15 | `2.0, 2.0` | 2/2 | ✅ PASS |
| Weibull_Survival | rel<=5E-15 | rel | 7.91e-17 | `0.5, 3.0, 1.0` | 3/3 | ✅ PASS |
| Weibull_Variance | rel<=2E-14 | rel | 8.26e-15 | `2.0, 2.0` | 2/2 | ✅ PASS |
| PROB_LogBeta | rel<=5E-15 | rel | — | documented | 0/0 | 🔷 KNOWN LIMITATION |

> **Verdict tally** — FAIL: 0, KNOWN LIMITATION: 1, CHARACTERIZATION ONLY: 0, PENDING: 0.

> States: **PASS** meets the contract; **FAIL** exceeds it and must be investigated; **KNOWN LIMITATION** is a documented defect tracked in `accuracy_contracts.csv` (does not read as green); **CHARACTERIZATION ONLY** is measured for the record but not held to a pass/fail claim; **PENDING** is not yet measured. Errors are computed in Decimal from the two-part hi;lo export, so a miss is a real miss (no precision-floor exemption).
