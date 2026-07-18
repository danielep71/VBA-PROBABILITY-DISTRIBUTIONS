# Accuracy benchmarks

Reproducible evidence for the accuracy claims published in the VBA source
comments — so those claims are artifacts a contributor can regenerate, not
statements that have to be taken on trust.

Each function that publishes a measured error level (in `SPECIALFUNCS` and
`TFAMILY`) is evaluated on a fixed input grid, compared against a 50-digit
mpmath reference, and checked against its own claim.

## Why two phases

The reference values are computed in Python (mpmath). The library under test is
VBA and can only run inside Excel. So the harness is split: Python owns the
reference and the error analysis; a small Excel macro owns the observed values.
Neither side trusts the other's numbers — Python never sees the library's code,
and the macro never reads the reference column.

## Files

| File | Role |
|---|---|
| `generate_reference_values.py` | Builds the input grid and 50-digit mpmath reference values. Phase 1. |
| `probability_accuracy_grid.csv` | The grid: inputs, reference, and an empty `observed_vba` column. |
| `M_STATS_PROBDIST_ACCURACYEXPORT.bas` | Excel macro that fills `observed_vba` by calling the library. Phase 2. |
| `compute_errors.py` | Joins observed vs reference, finds max-error locations, checks each claim, writes the summary. Phase 3. |
| `accuracy_summary.md` | The generated verdict table. |
| `environment.txt` | Python and dependency versions, reference precision, date. |

## Running it

```
# Phase 1 — reference (Python)
python generate_reference_values.py            # -> probability_accuracy_grid.csv

# Phase 2 — observed (Excel)
#   Import M_STATS_PROBDIST_ACCURACYEXPORT.bas into the workbook,
#   put probability_accuracy_grid.csv beside the workbook,
#   run Export_Accuracy_Observations. It fills the observed_vba column.

# Phase 3 — analysis (Python)
python compute_errors.py                       # -> accuracy_summary.md
```

`compute_errors.py` degrades honestly: any row whose `observed_vba` is still
empty is reported as *not measured* and excluded from the pass/fail check, so a
partial export produces a partial — not a misleading — summary.

> **Observed-value format.** VBA writes each observation as a two-part sum `hi;lo` (two 15-digit numbers), because VBA cannot emit more than ~15 significant digits in one literal. `compute_errors.py` sums the parts to recover the full-precision Double, so the harness measures accuracy below the 15-digit floor. A single number is also accepted for backward compatibility.

## Claims under test

Taken verbatim from the `ACCURACY` comments in the source:

The harness covers four modules: the special-function kernels
(`SPECIALFUNCS`), the test-statistic families (`TFAMILY`), the normal and
lognormal family (`NORMALFAMILY`), and the continuous distributions
(`CONTINUOUS`) — 66 functions in total.

**Special functions**

| Function | Published claim |
|---|---|
| `PROB_LogGamma` | relative error < 6.1E-14 for Z in [1E-8, 1E+50] |
| `PROB_LogGammaHalfDiff` | relative error <= 2E-14 for Z > 0 (tested range) |
| `PROB_StirlingError` | absolute error <= 3E-17 for N >= 0.5 |
| `PROB_LogChoose` | relative error <= 3.2E-16 for N in [2, 2^53], all K |

**Test-statistic families**

| Function | Published claim |
|---|---|
| Student t density | relative error <= 2E-14 (tested range) |
| Student t cumulative / survival | relative error <= 1.3E-12 |
| Student t quantile | relative error <= 3.0E-12 |
| Chi-square cumulative / survival | relative error <= 2.6E-10 |
| Chi-square quantile | relative error <= 4.7E-12 |
| F cumulative / survival | relative error <= 1.1E-10 |
| F quantile | relative error <= 5.9E-13 |

**Normal and lognormal family**

| Function | Published claim |
|---|---|
| Standard normal density / CDF / survival / inverse / inverse survival / interval | relative error <= 5E-15 |
| Standard normal fast inverse (raw Acklam) | relative error <= 5E-9 |
| General normal density / CDF / survival / inverse / inverse survival / z-score | relative error <= 5E-15 |
| Lognormal density / CDF / survival / inverse / inverse survival | relative error <= 5E-15 |
| Lognormal mean / variance / std dev / parameter conversion | relative error <= 5E-15 |

The normal-family `~1E-15` source comments are interpreted as a hard bound of
5E-15 for the harness. `NormalStandard_Survival` measures about 1.5E-14 at
moderate tail values, above that bound; it is reported as *below harness
precision* rather than a failure, and its claim is left as documented.

**Continuous distributions**

Bounds were set from the measured worst-case error over the tested grid, not
from source comments (the module publishes none). Exponential is parameterized
by rate (Lambda), not scale.

<!-- Generated from accuracy_contracts.csv by render_contract_table.py. Do not hand-edit. -->

| Function | Metric | Threshold | Domain |
|---|---|---|---|
| Beta_Cumulative | relative | 2E-14 | balanced to moderately unbalanced arguments |
| Beta_Density | relative | 5E-15 | balanced to moderately unbalanced arguments |
| Beta_InverseCumulative | relative | 5E-15 | balanced to moderately unbalanced arguments |
| Beta_Mean | relative | 5E-15 | balanced to moderately unbalanced arguments |
| Beta_StdDev | relative | 5E-15 | balanced to moderately unbalanced arguments |
| Beta_Survival | relative | 5E-15 | balanced to moderately unbalanced arguments |
| Beta_Variance | relative | 5E-15 | balanced to moderately unbalanced arguments |
| ChiSquare_Cumulative | relative | 2.6E-10 | full tested range |
| ChiSquare_InverseCumulative | relative | 4.7E-12 | full tested range |
| ChiSquare_Survival | relative | 2.6E-10 | full tested range |
| Exponential_Cumulative | relative | 5E-15 | full tested range |
| Exponential_Density | relative | 5E-15 | full tested range |
| Exponential_InverseCumulative | relative | 5E-15 | full tested range |
| Exponential_Survival | relative | 5E-15 | full tested range |
| F_Cumulative | relative | 1.1E-10 | balanced to moderately unbalanced arguments |
| F_InverseCumulative | relative | 5.9E-13 | balanced to moderately unbalanced arguments |
| F_Survival | relative | 1.1E-10 | balanced to moderately unbalanced arguments |
| Gamma_Cumulative | relative | 2E-14 | full tested range |
| Gamma_Density | relative | 2E-14 | full tested range |
| Gamma_InverseCumulative | relative | 2E-14 | full tested range |
| Gamma_Mean | relative | 5E-15 | full tested range |
| Gamma_StdDev | relative | 5E-15 | full tested range |
| Gamma_Survival | relative | 2E-14 | full tested range |
| Gamma_Variance | relative | 5E-15 | full tested range |
| LogChoose | relative | 3.2E-16 | full tested range |
| LogGamma | relative | 6.1E-14 | full tested range |
| LogGammaHalfDiff | relative | 2E-14 | full tested range |
| Lognormal_Cumulative | relative | 5E-15 | full tested range |
| Lognormal_Density | relative | 5E-15 | full tested range |
| Lognormal_InverseCumulative | relative | 5E-15 | full tested range |
| Lognormal_InverseSurvival | relative | 5E-15 | full tested range |
| Lognormal_Mean | relative | 5E-15 | full tested range |
| Lognormal_ParamMeanLog | relative | 5E-15 | full tested range |
| Lognormal_ParamStdDevLog | relative | 5E-15 | full tested range |
| Lognormal_StdDev | relative | 5E-15 | full tested range |
| Lognormal_Survival | relative | 5E-15 | full tested range |
| Lognormal_Variance | relative | 5E-15 | full tested range |
| NormalStandard_Cumulative | relative | 5E-15 | full tested range |
| NormalStandard_Density | relative | 5E-15 | full tested range |
| NormalStandard_IntervalProbability | relative | 5E-15 | full tested range |
| NormalStandard_InverseCumulative | relative | 5E-15 | full tested range |
| NormalStandard_InverseCumulativeFast | relative | 5E-9 | full tested range |
| NormalStandard_InverseSurvival | relative | 5E-15 | full tested range |
| NormalStandard_Survival | relative | 5E-15 | full tested range |
| Normal_Cumulative | relative | 5E-15 | full tested range |
| Normal_Density | relative | 5E-15 | full tested range |
| Normal_InverseCumulative | relative | 5E-15 | full tested range |
| Normal_InverseSurvival | relative | 5E-15 | full tested range |
| Normal_Survival | relative | 5E-15 | full tested range |
| Normal_ZScore | relative | 5E-15 | full tested range |
| StirlingError | absolute | 3E-17 | full tested range |
| StudentT_Cumulative | relative | 1.3E-12 | full tested range |
| StudentT_Density | relative | 2E-14 | full tested range |
| StudentT_InverseCumulative | relative | 3.0E-12 | full tested range |
| StudentT_Survival | relative | 1.3E-12 | full tested range |
| Uniform_Cumulative | relative | 5E-15 | full tested range |
| Uniform_Density | relative | 5E-15 | full tested range |
| Uniform_InverseCumulative | relative | 5E-15 | full tested range |
| Uniform_Survival | relative | 5E-15 | full tested range |
| Weibull_Cumulative | relative | 5E-15 | full tested range |
| Weibull_Density | relative | 5E-15 | full tested range |
| Weibull_InverseCumulative | relative | 5E-15 | full tested range |
| Weibull_Mean | relative | 2E-14 | full tested range |
| Weibull_StdDev | relative | 5E-15 | full tested range |
| Weibull_Survival | relative | 5E-15 | full tested range |
| Weibull_Variance | relative | 2E-14 | full tested range |

**Known limitations**

| Function | Domain | Notes |
|---|---|---|
| PROB_LogBeta | shape ratio >= 1E-1, or <= 1E-15 | Middle band (ratio 1E-2..1E-13) does not meet claim; see logbeta_study |

The Gamma and Beta inverse functions are iterative, yet measure near machine
epsilon (Gamma 9.7E-15, Beta 5.7E-16), so they hold the same tight bounds as
the closed-form functions.

**Unbalanced-argument caveat (Beta, F).** These Beta bounds, and the F functions
that depend on the incomplete beta, were verified for balanced-to-moderately-
unbalanced arguments (shape ratio min/max down to about 0.1). For more extreme
imbalance the accuracy degrades: `PROB_LogBeta`'s defining three-log-gamma
identity cancels, with measured relative error growing from roughly 1E-14 near
ratio 1E-2 to a few percent near ratio 1E-14, before a one-term asymptotic
restores full precision below ratio 1E-15. The closed-form Beta mean, variance
and standard deviation do not use `PROB_LogBeta` and are unaffected; Student t is
largely protected by its half-integer normalization. See `logbeta_study/` for the
measured curve. Repositioning or extending the asymptotic switch is deferred to a
validated future pass.

## Metric note

`PROB_StirlingError` is checked on **absolute** error, not relative: its source
comment explains that relative error is the wrong metric there (it reaches
1.5E-13 near N = 501, where delta itself is 1.67E-04), because what propagates
into a log-probability is the absolute error. The harness honours that choice
per function via the `metric` column.

## Reference integrity

The mpmath references were cross-checked against SciPy for every function; the
two independent oracles agree to about 1E-14 or better across the grid. The
survival references are computed on the upper tail directly (never `1 - CDF`),
so they stay accurate in the deep tail where a naive subtraction would collapse.
