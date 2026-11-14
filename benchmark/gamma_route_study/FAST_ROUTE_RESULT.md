# #314: the fast Gamma route — result

Recorded 2026-11-14 (Europe/Rome). `FAST_ROUTE_PREREGISTRATION.md` and its
Addendum 1 were committed and pushed at `e5e5ec4`, with every observed field
of `holdout_route_grid.csv` blank. The rule was wired at `c7042c5` and this
result was exported from that commit.

## Evidence

- Excel run 43569839650 on `c7042c5`, runner `ITPF545VX1`, Excel 16.0 build
  20430, 64-bit: 1,922 of 1,922 regression assertions and 13,005 of 13,005
  in the study and slow production batteries; evidence commit `3a56c8d`, the
  direct child of the candidate.
- The main grid, the consolidated holdout and the routing holdout were
  exported in that one run.
- `holdout_route_grid.csv`: 371 rows, raw SHA-256
  `c1018ea57210f6c3db8e6ff81833fb3ebd208e04c00725a4a32a746ed8b3410e`.

## Verdict

`python3 holdout_route_analyze.py`: **ACCEPTED: R1 to R6 hold on every row.**
The command is part of the evidence suite, so the verdict is recomputed in CI.

| Count | Value |
| --- | ---: |
| Rows | 371 |
| Tails scored against their reference (R1) | 566 |
| Tails with an exact reference of 0 or 1, returned exactly | 100 |
| Tails below the floor, returned by the kernel bit for bit (R5) | 33 |
| Tails within one percent below the floor, scored by neither | 43 |
| Adjacent pairs checked for order (R3) | 702 |

Worst relative error by kind of row, against a bound of 1E-13:

| Kind | Worst relative error |
| --- | ---: |
| Holdout points | 1.49E-15 |
| Transition probes | 4.33E-15 |
| Shape probes at `A = 1` | 1.64E-15 |

## Predictions against the outcome

1. **R1 to R6 on all 371 rows, worst error below 1E-14: met.** The worst is
   4.33E-15.
2. **Bit equality with the model wherever it does not delegate: not fully
   met.** 694 of 708 compared values equal the model's. The other 14 differ
   by one or two units in the last place, all at the holdout shape 6.17, on
   series, continued-fraction and complemented values alike. The model
   matched Excel at all 392 fit points; this shape is not one of them. The
   prediction was reported, not gating, and every one of the 14 is within
   R1. The model stays a development tool and is not evidence of the
   production value.
3. **Changed rows of the grids: met in part.**

   | Statement | Predicted | Observed |
   | --- | --- | --- |
   | Main grid, cumulative and survival rows outside the #13 regimes | 87 of 163 change, at most 23 ulps | 88 change, at most 23 ulps |
   | Consolidated holdout, the same functions | 3 of 6, at most 4 ulps | 3, at most 4 ulps |
   | Inverses | may change | 19 change: Chi-square 17, Gamma 2 |
   | Any other function | none | none |
   | Positive-ratio regimes of #13 | none | **9 change** |
   | Every changed row within its guarantee | all | all |

   The extra row is a Chi-square survival: 26 changed where 25 were
   predicted. The prediction on the #13 regimes was wrong. Their cumulative
   dispatch recovers only part of the regime (`AMENDMENT_1.md`); the rest
   goes to the ordinary route, and that route is now the routed one. Eight rows of
   `ChiSquare_Cumulative.positive_ratio_subnormal` moved by at most 145 ulps,
   about 3.2E-14 relative, and one of `Gamma_Cumulative.positive_ratio_subnormal`
   by 7. All nine stay inside their 5E-13 contracts, and the accuracy gate
   passes on all 172 active contracts.
4. **Cost: see below.**

## A corner the sweep under-sampled

Where the argument is so small that `|A ln X|` is large, the series forms its leading factor as `exp(A ln X - X - LogGamma(A + 1))`.
The rounding of that exponent is amplified by its size: at the largest
exponent the routine accepts before the floor of 1E-290 hands the call to
the kernel, the bound on the relative error is about 7.4E-14. That is inside
the guarantee of 1E-13, with a thin margin. The seeded sweep of the
development replay placed few points there, and the nine rows of the #13
regimes above are the measured instance: 3.2E-14. The holdout has no failure
in this corner, and the acceptance rules are not retuned for it. A wider
margin there would need its own registration.

## What stays with the kernel

As registered: an upper tail below `X = A + 1` when the lower tail exceeds
0.9, any directly returned tail below 1E-290, the inverses for shapes below
0.01, and any call the binary64 routine declines. Their accuracy and their
cost are unchanged.
