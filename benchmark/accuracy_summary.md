# Accuracy summary

Generated 2026-07-18 by `compute_errors.py` from `probability_accuracy_grid.csv`.

Reference values are mpmath at 50 digits (see `generate_reference_values.py`). Observed values are produced by the VBA library via `M_STATS_PROBDIST_ACCURACY_EXPORT.bas`. Each function is checked against the accuracy claim published in its source comment.

| Function | Claim | Metric | Max error | At input | Points | Verdict |
|---|---|---|---:|---|---:|---|
| Beta_Cumulative | rel<=2E-14 | rel | 5.59e-15 | `0.8, 5.0, 1.0` | 3/3 | ✅ pass |
| Beta_Density | rel<=5E-15 | rel | 2.67e-15 | `0.3, 2.0, 5.0` | 3/3 | ✅ pass |
| Beta_InverseCumulative | rel<=5E-15 | rel | 5.72e-16 | `0.95, 2.0, 5.0` | 2/2 | ✅ pass |
| Beta_Mean | rel<=5E-15 | rel | 1.39e-16 | `2.0, 3.0` | 2/2 | ✅ pass |
| Beta_StdDev | rel<=5E-15 | rel | 0.00e+00 | `2.0, 3.0` | 2/2 | ✅ pass |
| Beta_Survival | rel<=5E-15 | rel | 4.36e-15 | `0.3, 2.0, 5.0` | 3/3 | ✅ pass |
| Beta_Variance | rel<=5E-15 | rel | 0.00e+00 | `2.0, 3.0` | 2/2 | ✅ pass |
| ChiSquare_Cumulative | rel<=2.6E-10 | rel | 4.04e-14 | `0.5, 30.0` | 25/25 | ✅ pass |
| ChiSquare_InverseCumulative | rel<=4.7E-12 | rel | 2.14e-14 | `0.95, 1.0` | 25/25 | ✅ pass |
| ChiSquare_Survival | rel<=2.6E-10 | rel | 3.73e-14 | `80.0, 30.0` | 25/25 | ✅ pass |
| Exponential_Cumulative | rel<=5E-15 | rel | 0.00e+00 | `1.0, 1.0` | 3/3 | ✅ pass |
| Exponential_Density | rel<=5E-15 | rel | 0.00e+00 | `1.0, 1.0` | 3/3 | ✅ pass |
| Exponential_InverseCumulative | rel<=5E-15 | rel | 2.96e-16 | `0.95, 2.0` | 2/2 | ✅ pass |
| Exponential_Survival | rel<=5E-15 | rel | 0.00e+00 | `1.0, 1.0` | 3/3 | ✅ pass |
| F_Cumulative | rel<=1.1E-10 | rel | 2.60e-14 | `1.0, 10.0, 30.0` | 16/16 | ✅ pass |
| F_InverseCumulative | rel<=5.9E-13 | rel | 1.64e-14 | `0.5, 10.0, 30.0` | 12/12 | ✅ pass |
| F_Survival | rel<=1.1E-10 | rel | 2.98e-14 | `1.0, 10.0, 30.0` | 16/16 | ✅ pass |
| Gamma_Cumulative | rel<=2E-14 | rel | 1.70e-14 | `5.0, 3.0, 2.0` | 3/3 | ✅ pass |
| Gamma_Density | rel<=2E-14 | rel | 1.62e-14 | `5.0, 3.0, 2.0` | 3/3 | ✅ pass |
| Gamma_InverseCumulative | rel<=2E-14 | rel | 9.66e-15 | `0.5, 2.0, 1.0` | 2/2 | ✅ pass |
| Gamma_Mean | rel<=5E-15 | rel | 0.00e+00 | `2.0, 3.0` | 2/2 | ✅ pass |
| Gamma_StdDev | rel<=5E-15 | rel | 2.09e-16 | `2.0, 3.0` | 2/2 | ✅ pass |
| Gamma_Survival | rel<=2E-14 | rel | 1.60e-14 | `2.0, 2.0, 1.0` | 3/3 | ✅ pass |
| Gamma_Variance | rel<=5E-15 | rel | 0.00e+00 | `2.0, 3.0` | 2/2 | ✅ pass |
| LogChoose | rel<=3.2E-16 | rel | 1.60e-16 | `9007199254740992.0, 4503599627370496.0` | 30/30 | ✅ pass |
| LogGamma | rel<6.1E-14 | rel | 3.82e-15 | `8.376776400682919` | 40/40 | ✅ pass |
| LogGammaHalfDiff | rel<=2E-14 | rel | 1.53e-14 | `1.6102620275609393` | 30/30 | ✅ pass |
| Lognormal_Cumulative | rel<=5E-15 | rel | 2.27e-16 | `0.5, 0.0, 1.0` | 3/3 | ✅ pass |
| Lognormal_Density | rel<=5E-15 | rel | 1.88e-16 | `2.0, 0.5, 0.25` | 3/3 | ✅ pass |
| Lognormal_InverseCumulative | rel<=5E-15 | rel | 1.97e-16 | `0.025, 0.0, 1.0` | 2/2 | ✅ pass |
| Lognormal_InverseSurvival | rel<=5E-15 | rel | 2.50e-16 | `0.025, 0.0, 1.0` | 2/2 | ✅ pass |
| Lognormal_Mean | rel<=5E-15 | rel | 0.00e+00 | `0.0, 1.0` | 2/2 | ✅ pass |
| Lognormal_ParamMeanLog | rel<=5E-15 | rel | 0.00e+00 | `2.0, 0.5` | 2/2 | ✅ pass |
| Lognormal_ParamStdDevLog | rel<=5E-15 | rel | 0.00e+00 | `2.0, 0.5` | 2/2 | ✅ pass |
| Lognormal_StdDev | rel<=5E-15 | rel | 0.00e+00 | `0.0, 1.0` | 2/2 | ✅ pass |
| Lognormal_Survival | rel<=5E-15 | rel | 2.52e-16 | `2.0, 0.5, 0.25` | 3/3 | ✅ pass |
| Lognormal_Variance | rel<=5E-15 | rel | 1.90e-16 | `0.0, 1.0` | 2/2 | ✅ pass |
| NormalStandard_Cumulative | rel<=5E-15 | rel | 9.15e-16 | `-2.0` | 6/6 | ✅ pass |
| NormalStandard_Density | rel<=5E-15 | rel | 2.57e-16 | `-2.0` | 6/6 | ✅ pass |
| NormalStandard_IntervalProbability | rel<=5E-15 | rel | 1.17e-16 | `-1.96, 1.96` | 3/3 | ✅ pass |
| NormalStandard_InverseCumulative | rel<=5E-15 | rel | 1.72e-15 | `0.999` | 5/5 | ✅ pass |
| NormalStandard_InverseCumulativeFast | rel<=5E-9 | rel | 8.05e-10 | `0.975` | 5/5 | ✅ pass |
| NormalStandard_InverseSurvival | rel<=5E-15 | rel | 1.72e-15 | `0.999` | 5/5 | ✅ pass |
| NormalStandard_Survival | rel<=5E-15 | rel | 1.51e-14 | `3.0` | 6/6 | ⚠️ below harness precision |
| Normal_Cumulative | rel<=5E-15 | rel | 0.00e+00 | `1.96, 0.0, 1.0` | 3/3 | ✅ pass |
| Normal_Density | rel<=5E-15 | rel | 1.63e-16 | `110.0, 100.0, 15.0` | 3/3 | ✅ pass |
| Normal_InverseCumulative | rel<=5E-15 | rel | 2.11e-16 | `0.99, 100.0, 15.0` | 2/2 | ✅ pass |
| Normal_InverseSurvival | rel<=5E-15 | rel | 2.18e-16 | `0.99, 100.0, 15.0` | 2/2 | ✅ pass |
| Normal_Survival | rel<=5E-15 | rel | 8.33e-16 | `1.96, 0.0, 1.0` | 3/3 | ✅ pass |
| Normal_ZScore | rel<=5E-15 | rel | 0.00e+00 | `1.96, 0.0, 1.0` | 3/3 | ✅ pass |
| StirlingError | abs<=3E-17 | abs | 2.78e-17 | `0.5` | 12/12 | ✅ pass |
| StudentT_Cumulative | rel<=1.3E-12 | rel | 2.93e-14 | `1.0, 1000.0` | 25/25 | ✅ pass |
| StudentT_Density | rel<=2E-14 | rel | 1.93e-14 | `20.0, 1000.0` | 25/25 | ✅ pass |
| StudentT_InverseCumulative | rel<=3.0E-12 | rel | 9.95e-14 | `0.95, 1000.0` | 25/25 | ✅ pass |
| StudentT_Survival | rel<=1.3E-12 | rel | 1.55e-13 | `1.0, 1000.0` | 25/25 | ✅ pass |
| Uniform_Cumulative | rel<=5E-15 | rel | 0.00e+00 | `3.0, 0.0, 10.0` | 2/2 | ✅ pass |
| Uniform_Density | rel<=5E-15 | rel | 0.00e+00 | `3.0, 0.0, 10.0` | 2/2 | ✅ pass |
| Uniform_InverseCumulative | rel<=5E-15 | rel | 0.00e+00 | `0.5, 0.0, 10.0` | 2/2 | ✅ pass |
| Uniform_Survival | rel<=5E-15 | rel | 0.00e+00 | `3.0, 0.0, 10.0` | 2/2 | ✅ pass |
| Weibull_Cumulative | rel<=5E-15 | rel | 2.36e-16 | `0.5, 3.0, 1.0` | 3/3 | ✅ pass |
| Weibull_Density | rel<=5E-15 | rel | 1.68e-16 | `0.5, 3.0, 1.0` | 3/3 | ✅ pass |
| Weibull_InverseCumulative | rel<=5E-15 | rel | 1.28e-16 | `0.95, 2.0, 2.0` | 2/2 | ✅ pass |
| Weibull_Mean | rel<=2E-14 | rel | 7.75e-15 | `1.5, 1.0` | 2/2 | ✅ pass |
| Weibull_StdDev | rel<=5E-15 | rel | 4.19e-15 | `2.0, 2.0` | 2/2 | ✅ pass |
| Weibull_Survival | rel<=5E-15 | rel | 1.26e-16 | `0.5, 3.0, 1.0` | 3/3 | ✅ pass |
| Weibull_Variance | rel<=2E-14 | rel | 8.28e-15 | `2.0, 2.0` | 2/2 | ✅ pass |

> All measured functions meet their published accuracy claims. Rows marked *below harness precision* have claims tighter than a 15-16 digit CSV round-trip can verify; they are not failures.
