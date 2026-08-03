# Student t density at large df

Measures `K_STATS_StudentT_Density` across df 1E2–1E20, so it can be held to a
contract like the other four continuous densities.

## Why this study exists

CR-P1-01B gave the Gamma, Beta, ChiSquare and F densities a measured envelope
(`PROB_DENSITY_SHAPE_MAX = 1E20`) and frozen contracts. **StudentT_Density was
left out.** It carried no envelope and no large-df study, and its only coverage
above df 1E6 asserted merely that it returned *a* value:

```vba
AssertTrue "t density not enveloped (large df)", _
    (Not IsError(K_STATS_StudentT_Density(0#, 1E+18)))
```

That is a liveness check, not an accuracy check. The t density was also a
plausible carrier of the very defect CR-P1-01 fixed elsewhere, because it needs

```
LogGamma((df+1)/2) - LogGamma(df/2)
```

a difference of two large log-gammas of exactly the kind that lost accuracy in
the Gamma, Beta and F densities.

## What the measurement found

**It does not carry that defect.** `PROB_LogGammaHalfDiff` evaluates the
difference by an asymptotic series in 1/Z above a cutoff instead of subtracting
two large values, so the cancellation never forms. Measured against mpmath at 50
digits, the density is machine-precise to df 1E20 across the body and into the
tail — worst about 1E-14 at x = 10, where the `(df+1)/2 * Log1p(x*x/df)` term is
largest.

This study therefore exists to **contract** that behaviour, not to repair it.
Being able to say "measured, not merely believed" is the point.

## References

mpmath at 50 digits. For df >= 1E8 each reference is additionally cross-checked
against the standard normal density, which the t density approaches as df grows;
the gap is O(1/df), so at large df that limit is an independent check on the
reference rather than a restatement of it.

## Files

| file | role |
| --- | --- |
| `generate_t_density_large_df.py` | builds the grid with cross-checked references |
| `t_density_large_df_grid.csv` | the grid; `observed_vba` filled by the macro |
| `M_STATS_PROBDIST_TDENS.bas` | `Export_TDensityLargeDf` writes `observed_vba` |
| `analyze_t_density_large_df.py` | worst relative error by df and by x |

## Procedure

1. `python generate_t_density_large_df.py`
2. Import `M_STATS_PROBDIST_TDENS.bas`, run `Export_TDensityLargeDf`
3. `python analyze_t_density_large_df.py`
4. Promote the points into the main grid and freeze a contract
