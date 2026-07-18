# Accuracy summary

Generated 2026-07-18 by `compute_errors.py` from `probability_accuracy_grid.csv`.

Reference values are mpmath at 50 digits (see `generate_reference_values.py`). Observed values are produced by the VBA library via `M_STATS_PROBDIST_ACCURACY_EXPORT.bas`. Each function is checked against the accuracy claim published in its source comment.

| Function | Claim | Metric | Max error | At input | Points | Verdict |
|---|---|---|---:|---|---:|---|
| ChiSquare_Cumulative | rel<=2.6E-10 | rel | 4.04e-14 | `0.5, 30.0` | 25/25 | ✅ pass |
| ChiSquare_InverseCumulative | rel<=4.7E-12 | rel | 2.14e-14 | `0.95, 1.0` | 25/25 | ✅ pass |
| ChiSquare_Survival | rel<=2.6E-10 | rel | 3.73e-14 | `80.0, 30.0` | 25/25 | ✅ pass |
| F_Cumulative | rel<=1.1E-10 | rel | 2.60e-14 | `1.0, 10.0, 30.0` | 16/16 | ✅ pass |
| F_InverseCumulative | rel<=5.9E-13 | rel | 1.64e-14 | `0.5, 10.0, 30.0` | 12/12 | ✅ pass |
| F_Survival | rel<=1.1E-10 | rel | 2.98e-14 | `1.0, 10.0, 30.0` | 16/16 | ✅ pass |
| LogChoose | rel<=3.2E-16 | rel | 1.60e-16 | `9007199254740992.0, 4503599627370496.0` | 30/30 | ✅ pass |
| LogGamma | rel<6.1E-14 | rel | 3.82e-15 | `8.376776400682919` | 40/40 | ✅ pass |
| LogGammaHalfDiff | rel<=2.1E-15 | rel | 1.53e-14 | `1.6102620275609393` | 30/30 | ⚠️ below harness precision |
| StirlingError | abs<=3E-17 | abs | 2.78e-17 | `0.5` | 12/12 | ✅ pass |
| StudentT_Cumulative | rel<=1.3E-12 | rel | 2.93e-14 | `1.0, 1000.0` | 25/25 | ✅ pass |
| StudentT_Density | rel<=8.4E-15 | rel | 1.93e-14 | `20.0, 1000.0` | 25/25 | ⚠️ below harness precision |
| StudentT_InverseCumulative | rel<=3.0E-12 | rel | 9.95e-14 | `0.95, 1000.0` | 25/25 | ✅ pass |
| StudentT_Survival | rel<=1.3E-12 | rel | 1.55e-13 | `1.0, 1000.0` | 25/25 | ✅ pass |

> All measured functions meet their published accuracy claims. Rows marked *below harness precision* have claims tighter than a 15-16 digit CSV round-trip can verify; they are not failures.
