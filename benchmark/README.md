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
| `M_STATS_PROBDIST_ACCURACY_EXPORT.bas` | Excel macro that fills `observed_vba` by calling the library. Phase 2. |
| `compute_errors.py` | Joins observed vs reference, finds max-error locations, checks each claim, writes the summary. Phase 3. |
| `accuracy_summary.md` | The generated verdict table. |
| `environment.txt` | Python and dependency versions, reference precision, date. |

## Running it

```
# Phase 1 — reference (Python)
python generate_reference_values.py            # -> probability_accuracy_grid.csv

# Phase 2 — observed (Excel)
#   Import M_STATS_PROBDIST_ACCURACY_EXPORT.bas into the workbook,
#   put probability_accuracy_grid.csv beside the workbook,
#   run Export_Accuracy_Observations. It fills the observed_vba column.

# Phase 3 — analysis (Python)
python compute_errors.py                       # -> accuracy_summary.md
```

`compute_errors.py` degrades honestly: any row whose `observed_vba` is still
empty is reported as *not measured* and excluded from the pass/fail check, so a
partial export produces a partial — not a misleading — summary.

## Claims under test

Taken verbatim from the `ACCURACY` comments in the source:

| Function | Published claim |
|---|---|
| `PROB_LogGamma` | relative error < 6.1E-14 for Z in [1E-8, 1E+50] |
| `PROB_LogGammaHalfDiff` | relative error <= 2.1E-15 for Z > 0 |
| `PROB_StirlingError` | absolute error <= 3E-17 for N >= 0.5 |
| `PROB_LogChoose` | relative error <= 3.2E-16 for N in [2, 2^53], all K |
| Student t density | relative error <= 8.4E-15 |
| Student t cumulative / survival | relative error <= 1.3E-12 |
| Student t quantile | relative error <= 3.0E-12 |
| Chi-square cumulative / survival | relative error <= 2.6E-10 |
| Chi-square quantile | relative error <= 4.7E-12 |
| F cumulative / survival | relative error <= 1.1E-10 |
| F quantile | relative error <= 5.9E-13 |

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
