# Inverse probe

Measures where the StudentT, ChiSquare and F **inverse** functions actually
converge, so their envelopes can be set from evidence rather than inherited.

## Why the inverses are separate

The forward envelopes were raised to df 1E8 (StudentT, ChiSquare) and 1E10 (F)
on the evidence of `benchmark/envelope_probe`. **That study measured the forward
kernels only.** The inverses sit on top of them with safeguarded Newton plus
bisection and their own iteration budget, and the difference is not theoretical:

```
K_STATS_F_InverseCumulative(0.5, 1E6, 3)  ->  #NUM!
```

df 1E6 is comfortably inside the raised forward cap of 1E10, yet the inverse
refuses. Extending a forward measurement to the inverse would therefore have
widened the accepted domain past the demonstrated one — the precise failure
CR-P1-01B was reopened for. The inverses consequently kept their previously
validated caps:

| function | constant | cap |
| --- | --- | --- |
| `F_InverseCumulative` | `PROB_F_INV_MAX_DF` | 1E5 |
| `StudentT_InverseCumulative` | `PROB_T_INV_MAX_DF` | 1E6 |
| `ChiSquare_InverseCumulative` | `PROB_CHI_INV_MAX_DF` | 1E6 |

## Approach

As in `envelope_probe`, the public inverses reject beyond their caps, so this
grid probes the `Public` kernels directly with exactly the arguments each public
function would pass:

```
ChiSquare_InverseCumulative(p, df) -> PROB_TryGammaInvP(p, q, df/2)
F_InverseCumulative(p, d1, d2)     -> PROB_TryBetaInvRegularized(p, q, d1/2, d2/2)
StudentT_InverseCumulative(p, df)  -> the same beta inverse with B = 1/2
```

Both kernels take the probability **and its complement** as a pair, so no tail
is reconstructed by subtraction.

## Design note: ratio, not just magnitude

The known failure is at df ratio 1E6:3 — not at large balanced df. A grid that
varied only magnitude would have missed the case that motivated the study, so
the beta shapes cover balanced pairs, strongly unbalanced pairs in **both**
orientations, and the `B = 1/2` shape the t distribution produces.

## References

The reference is the quantile itself, found by safeguarded Newton on the
high-precision CDF (a Newton step inside the bracket, a bisection step
otherwise, so the bracket always shrinks). Each reference is then **fed back
through the CDF** and kept only if it reproduces the target probability; the
generator aborts before writing if the worst residual exceeds tolerance.

## Files

| file | role |
| --- | --- |
| `generate_inverse_probe.py` | builds the grid with round-trip-checked references |
| `inverse_probe_grid.csv` | the grid; `observed_vba` filled by the macro |
| `M_STATS_PROBDIST_INVPROBE.bas` | `Export_InverseProbe` writes `observed_vba` |
| `analyze_inverse_probe.py` | worst error and refusals per shape, against current caps |

## Procedure

1. `python generate_inverse_probe.py`
2. Import `M_STATS_PROBDIST_INVPROBE.bas`, run `Export_InverseProbe`
3. `python analyze_inverse_probe.py`
4. Raise each inverse cap to the highest **measured clean** shape, or reaffirm it

A refusal is a measurement: it is a clean `#NUM!`, never a wrong value, but it
caps the usable domain regardless of how accurate the converging points are.
