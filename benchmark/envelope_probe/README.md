# Envelope probe

Measures whether the StudentT, ChiSquare and F degrees-of-freedom envelopes are
still justified after the CR-P1-02 prefactor repair.

## The problem this solves

Current caps:

| family | constant | cap |
| --- | --- | --- |
| StudentT CDF/survival/inverse | `PROB_T_MAX_DF` | 1E6 |
| ChiSquare CDF/survival/inverse | `PROB_CHI_MAX_DF` | 1E6 |
| F CDF/survival/inverse | `PROB_F_MAX_DF` | 1E5 |

All three were set while the incomplete-gamma and incomplete-beta kernels still
carried the CR-P1-02 prefactor cancellation. The P2-02 chi divergence at
df = 1E16 (survival wrong by ~7E2) routes through the incomplete gamma at
A = df/2 = 5E15, where the old prefactor was wrong by roughly e^46 — so the cap
may have been measuring a defect rather than an intrinsic limit.

**The public functions reject df above their cap, so they cannot answer this.**

## Approach

The kernels beneath them are `Public` and carry no envelope, so this study calls
them directly with exactly the arguments each family would pass:

```
ChiSquare_Cumulative(X, df)  -> PROB_TryGammaRegularizedP(df/2, X/2)
StudentT_Cumulative(t, df)   -> PROB_TryBetaRegularized(x, y, df/2, 1/2)
                                x = df / (df + t*t)
F_Cumulative(f, d1, d2)      -> PROB_TryBetaRegularized(x, y, d1/2, d2/2)
                                x = d1*f / (d1*f + d2)
```

No source change is needed to run it. That matters: it means the caps are not
disturbed while the evidence for changing them is being collected.

## References

mpmath at 30 digits by the converging route, **revalidated at 45 digits**, with
anything that fails revalidation dropped rather than trusted. The generator
aborts before writing if the worst disagreement exceeds tolerance.

Reference availability, not ambition, sets the upper end of each family's range:
Chi to 1E9, StudentT to 1E8, F to 1E12. A cap may be raised only to a df that
was actually measured here.

## Files

| file | role |
| --- | --- |
| `generate_envelope_probe.py` | builds the grid with self-checked references |
| `envelope_probe_grid.csv` | the grid; `observed_vba` filled by the macro |
| `M_STATS_PROBDIST_ENVPROBE.bas` | `Export_EnvelopeProbe` writes `observed_vba` |
| `analyze_envelope_probe.py` | worst error per family and df, against current caps |

## Procedure

1. `python generate_envelope_probe.py`
2. Import `M_STATS_PROBDIST_ENVPROBE.bas`, run `Export_EnvelopeProbe`
3. `python analyze_envelope_probe.py`
4. Raise each cap to the highest **measured** df, or reaffirm it, citing this study

An `ERROR` cell is a measurement: it records that the kernel refused rather than
returning an unconverged value, which is the reachable-domain edge.
