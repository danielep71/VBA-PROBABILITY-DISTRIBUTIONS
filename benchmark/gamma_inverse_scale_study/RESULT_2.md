# #14: the Gamma inverse at the final scale — result of the second holdout

Recorded 2026-11-14 (Europe/Rome). `PREREGISTRATION_2.md` was committed and
pushed at `901174a`, with every observed field of
`gamma_inverse_scale_grid2.csv` blank. The second candidate was wired at
`664cd48` and this result was exported from that commit.

## Evidence

- Excel run 43592089271 on `664cd48`, runner `ITPF545VX1`, Excel 16.0 build
  20430, 64-bit: 1,933 of 1,933 regression assertions and 13,005 of 13,005
  in the study and slow production batteries; evidence commit `6d85020`, the
  direct child of the candidate.
- The main grid, the consolidated holdout and the second holdout were
  exported in that one run.
- `gamma_inverse_scale_grid2.csv`: 525 rows, raw SHA-256
  `914dec804a073a44dca6516766d716efb4ef1f66c9d308514dd92b87f441ff18`.

## Verdict

`python3 holdout_analyze.py --holdout 2`: **ACCEPTED: G1 to G7 hold on every
row.** The command is part of the evidence suite, so the verdict is
recomputed in CI.

| Count | Value |
| --- | ---: |
| Rows | 525 |
| Quantiles scored against a normal reference (G1) | 501 |
| Exact zeros, subnormal quantiles, `#NUM!` (G2) | 10, 10, 4 |
| Adjacent seam pairs checked for order (G3) | 160 |
| Rows at scale 2, Chi-square equal to Gamma bit for bit (G4) | 60 |
| Round trips through `Gamma_Cumulative` (G5) | 501 |
| Rows Excel answered in log space, on the unit scale | 409, 116 |

All ten seam clusters hold both branches (G7), which the first holdout's
did not.

Worst relative error, against a bound of 1E-12:

| Rows | Worst relative error |
| --- | ---: |
| Log space, shape up to 1 | 2.1E-16 |
| Log space, shape from 1 to 2 | 1.7E-16 |
| Log space, shape above 2 | 7.64E-15 |
| Unit-scale solver | 1.09E-14 |
| Round trip, residual against the probability | 4.03E-14 |

## Predictions against the outcome

1. **G1 to G7 on all 525 rows: met.**
2. **Log space below 5E-15, and below 1E-15 up to shape 2: met in part.**
   Up to shape 2 the worst is 2.1E-16. Above 2 it is 7.64E-15, over the
   predicted 5E-15: the five largest errors are the five rows at the
   successor of 2, where `PROB_TryLogGamma1p` takes over. The source of
   `LogGamma(1 + A)` therefore leaves a step in accuracy at shape 2, from
   2E-16 to 8E-15. Both sides are far inside the guarantee, and the
   prediction was reported, not gating.
3. **Only the 13 rows of the new regime change, and the four contracts meet
   their target: met.** Six of the 13 changed against the first candidate's
   export; no other row of the main grid and no row of the consolidated
   holdout did.

   | Contract | Target | First candidate | Second candidate |
   | --- | ---: | ---: | ---: |
   | `Gamma_InverseCumulative` quantile error | 1E-14 | 1.47E-14 | 2.13E-16 |
   | `Gamma_InverseCumulative` tail residual | 1E-14 | 9.74E-15 | 1.54E-15 |
   | `ChiSquare_InverseCumulative` quantile error | 1E-14 | 1.46E-14 | 1.80E-16 |
   | `ChiSquare_InverseCumulative` tail residual | 1E-14 | 7.28E-15 | 3.95E-15 |

4. **The 1,933 regression assertions pass unchanged: met.**
5. **Chi-square equals Gamma at scale 2 bit for bit on all 60 rows: met.**

Not predicted, and observed: for shapes up to 2 the public quantile equals
the model's `candidate_2` bit for bit on all 363 rows answered in log space,
and Excel's `L0` equals the model's on 416 of the 417 rows with shape up to
1.

## Cost

The #298 benchmark of the same run. Milliseconds per call, Gamma inverse at
probabilities 0.1 and 0.9, against run 43573743649 before #14:

| Shape | Before | After |
| --- | ---: | ---: |
| 0.001 | refused after 1,050; 1,310 | 0.021, 0.019 |
| 0.01 | 0.27, 95 | 0.022, 52 |
| 0.5 | 0.05, 0.15 | 0.046, 0.11 |

At shape 0.001 and probability 0.1 the quantile is `0.1^1000`, below every
Double: the call used to run for a second and return `#NUM!`, and now
returns zero, as the representability policy states. At shape 0.01 and
probability 0.9 the unit quantile is about `exp(-11)`, above the crossover,
so the unit-scale solver still answers at the kernel's cost. Timing depends
on the host; a second run moves single figures by up to a factor of two.

## What the rule does not cover

- A unit quantile above `exp(-40)` is solved on the unit scale as before.
  For shapes below 0.01 that solver uses the double-double kernel and stays
  slow (#348).
- Above shape 2 the log branch is as accurate as `PROB_TryLogGamma1p`,
  about 8E-15 next to shape 2.
- A subnormal final quantile is returned as the Double the exponential
  gives; G2 holds it to one unit of the subnormal spacing.
