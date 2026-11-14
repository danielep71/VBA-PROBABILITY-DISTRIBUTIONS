# #14: the Gamma inverse at the final scale — result of the first holdout

Recorded 2026-11-14 (Europe/Rome). `PREREGISTRATION.md` and its Addendum 1
were committed and pushed at `91a7705` and `ca669ec`, with every observed
field of `gamma_inverse_scale_grid.csv` blank. The rule was wired at
`9e1e531` and this result was exported from that commit.

## Evidence

- Excel run 43585348023 on `9e1e531`, runner `ITPF545VX1`, Excel 16.0 build
  20430, 64-bit: 1,933 of 1,933 regression assertions and 13,005 of 13,005
  in the study and slow production batteries; evidence commit `ffad5ec`, the
  direct child of the candidate.
- The main grid, the consolidated holdout and this holdout were exported in
  that one run.
- `gamma_inverse_scale_grid.csv`: 455 rows, raw SHA-256
  `e54edd818ecafb6704147cc899aa01cbc69e0eab497fff1a620620d6d512f81a`.

## Verdict

`python3 holdout_analyze.py`: **REJECTED, 5 failures, all of rule G7.**

| Rule | Outcome |
| --- | --- |
| G1 accuracy | holds on 431 scored rows; worst relative error 2.30E-14 against 1E-12 |
| G2 limits | holds: 10 exact zeros, 10 subnormal quantiles, 4 `#NUM!` |
| G3 monotone | holds on 120 adjacent pairs |
| G4 families | holds on 60 rows at scale 2 |
| G5 round trip | holds on 431 rows; worst residual 5.69E-14 |
| G6 numbers | holds |
| G7 seam | **fails at shapes 0.5, 1, 2.5, 7 and 15** |

At those five shapes all thirteen probes fall on the ordinary side of the
crossover, by the `L0` Excel exported. Near -40 one unit in the last place
of `L0` is 7.1E-15, and Excel's `L0` differs from the model's by up to
three units there. At these shapes a
unit in the last place of the probability moves `L0` by far less than that,
so thirty-two of them on either side never reach the crossover: the range
of `L0` across a cluster is one to three Doubles. The model's `LogGamma1p`
is the correctly rounded value and production's `PROB_TryLogGamma1p` is not;
the registration named that difference and G7 was the rule for it.

The rule is rejected, as the registration requires. Nothing is retuned
against this holdout.

## Predictions against the outcome

1. **G1 to G7 on all 455 rows: not met.** G7 fails.
2. **Rows answered in log space below 1E-14: not met.** The worst is
   2.30E-14. By decade of shape:

   | Shape | Worst relative error, log space |
   | --- | ---: |
   | 1E-8 to 0.1 | 2.1E-16 |
   | 0.1 to 1 | 2.30E-14 |
   | 1 to 10 | 1.03E-14 |
   | 10 to 16 | 3.0E-15 |

   The pair arithmetic is not the cause: where `LogGamma(1 + A)` is tiny in
   absolute value the result is at 2E-16. Between 0.1 and 2 the error is that
   of `PROB_TryLogGamma1p`, which there takes `PROB_LogGamma(1 + A)` near its
   zeros at 1 and 2. The model assumed a correctly rounded value and so
   predicted 3.5E-16.
3. **No existing observation changes: met.** The main grid changed in its 13
   new rows only, and the consolidated holdout not at all. The 1,922
   assertions that existed pass unchanged.
4. **Cost: not measured.** The cost benchmark was not run on this candidate.

The four contracts registered for the main grid were exported in the same
run. Their guarantee holds on all 13 rows; two targets were missed:

| Contract | Target | Worst |
| --- | ---: | ---: |
| `Gamma_InverseCumulative` quantile error | 1E-14 | 1.47E-14 |
| `Gamma_InverseCumulative` tail residual | 1E-14 | 9.74E-15 |
| `ChiSquare_InverseCumulative` quantile error | 1E-14 | 1.46E-14 |
| `ChiSquare_InverseCumulative` tail residual | 1E-14 | 7.28E-15 |

## What follows

The candidate is not merged. A second candidate is registered in
`PREREGISTRATION_2.md` with its own untouched holdout. The 455 rows above
are revealed and serve it as development data.
