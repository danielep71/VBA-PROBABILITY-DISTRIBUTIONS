# Unbalanced-Beta switch study (PROB_LogBeta)

## Why this exists

`PROB_LogBeta(A, B)` computes `Log(Beta(A, B))` two ways:

- **general identity** — `LogGamma(A) + LogGamma(B) - LogGamma(A + B)`;
- **asymptotic branch** — `LogGamma(Small) - Small * Log(Large)`, used only when
  `Small / Large <= PROB_EPS` (1E-15).

The general identity cancels catastrophically as the arguments become
unbalanced: `LogGamma(Large)` and `LogGamma(Large + Small)` are both large and
nearly equal, so their difference loses precision proportional to
`macheps * (Large / Small)`. The asymptotic branch avoids this, but currently
only engages at a ratio of 1E-15.

`PROB_LogBeta` is reached by `Beta_Density` directly and by the incomplete-beta
CDF kernel (so Beta, F and Student-t all depend on it), which makes any
unbalanced-argument error user-visible.

This study measures the **actual** `PROB_LogBeta` across the ratio range so the
switch point can be repositioned from evidence, not a model.

## Files

| File | Role |
|---|---|
| `generate_logbeta_switch.py` | Writes `logbeta_switch_grid.csv` — 90 reference rows (5 `Small` values x 18 ratios) at 50+ mpmath digits. |
| `logbeta_switch_grid.csv` | The grid: `arg1 = Large`, `arg2 = Small`, reference, empty `observed_vba`. |
| `M_STATS_PROBDIST_ACCURACYEXPORT.bas` | Export macro with an added `Case "LogBeta"` that calls `PROB_LogBeta(A1, A2)`. |
| `analyze_logbeta_switch.py` | Reads the filled grid and prints relative error vs ratio, flags rows above the 5E-15 Beta claim, and marks where the branch fires. |

## How to run

1. Import `M_STATS_PROBDIST_ACCURACYEXPORT.bas` into the workbook (Debug > Compile).
2. Run `Export_Accuracy_Observations`; when the file dialog appears, pick
   `logbeta_switch_grid.csv`.
3. Commit the filled CSV.
4. Analysis (done for you): `python3 analyze_logbeta_switch.py`.

## Expected pattern (from a Python model of the branch logic)

The general identity is expected to exceed the 5E-15 claim from roughly
ratio 1E-2 down to 1E-14 (worsening to a few percent near 1E-14), then the
branch engages at 1E-15 and returns to ~1E-16 accuracy. If the real
`PROB_LogBeta` confirms this, the switch is mispositioned by many orders of
magnitude and should fire far earlier (nearer 1E-3) rather than being tightened.

The VBA run is the authority; this expectation only sizes the study.
