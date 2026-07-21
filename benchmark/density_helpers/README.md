# Density helpers coverage

Benchmark coverage for four public UDFs that previously had no rows in the main
accuracy grid, so their header accuracy language was uncorroborated. Each is now
measured against a 50-digit mpmath reference on the main grid and gated by a
contract in `../accuracy_contracts.csv` (`evidence = density_helpers`).

| Function | What it computes | Worst measured rel. error | Contract |
|---|---|---:|---:|
| `ChiSquare_Density` | chi-square density | 2.37E-14 | 1E-13 |
| `F_Density` | F density | 4.81E-14 | 1E-13 |
| `Normal_IntervalProbability` | P(a ≤ X ≤ b) for a normal (a normal-density integral) | 1.17E-15 | 1E-14 |
| `Lognormal_ParametersFromMeanStdDev` | arithmetic (Mean, StdDev) → (MeanLog, StdDevLog) | 1.17E-16 | 5E-15 |

The name is approximate: the first three are density or density-integral
computations, while `Lognormal_ParametersFromMeanStdDev` is a moment-to-parameter
converter grouped here because it was part of the same coverage gap. It returns a
1×2 array `[MeanLog, StdDevLog]`, benchmarked as two rows distinguished by
`regime` (`param_meanlog` / `param_stddevlog`).

References were cross-checked against SciPy (chi-square and F densities agree to
~1E-14, SciPy's double-precision limit; the interval probability agrees exactly).

## Regenerating

The observations are already committed in the main grid. To re-measure after
changing any of these functions:

1. Import `density_helpers.bas`, `Debug > Compile`.
2. Run `Export_DensityHelpers`; select `../probability_accuracy_grid.csv`.
   It fills only rows tagged `evidence_set = density_helpers` and leaves every
   other observation untouched.
3. Re-run `python compute_errors.py` from `benchmark/` to re-gate.
