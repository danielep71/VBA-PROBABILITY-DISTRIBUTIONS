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
4. **Cost: not met as worded.** The prediction said microseconds up to
   shape 10,000. The binary64 routines do take 4 to 77 microseconds, but a
   public call takes 0.06 to 0.23 ms. The gain over the kernel is as
   expected; the absolute figure was optimistic by more than a factor of
   ten. The second part held: a lower-region survival with `S > 0.9` keeps
   the kernel's cost. See the next section.

## Cost per call

Excel run 43573743649 on `cc20679`, the same runner and build; the #298
benchmark, each figure the mean of repeated calls over at least half a
second. Timing depends on the host, and a second run moved single figures
by up to a factor of two. Milliseconds per call:

| Shape, argument | Kernel P, Q | Gamma CDF, survival | Chi-square CDF, survival | Poisson CDF, survival |
| --- | ---: | ---: | ---: | ---: |
| 0.0001, 0.5 | 13, 14 | 0.09, **15** | 0.12, **16** | |
| 0.5, 0.25 | 18, 17 | 0.09, 0.09 | 0.12, 0.10 | |
| 2.5, 1 | 20, 20 | 0.07, 0.08 | 0.13, 0.13 | |
| 2.5, 6 | 38, 36 | 0.08, 0.09 | 0.12, 0.09 | |
| 100, 90 | 37, 36 | 0.07, 0.08 | 0.13, 0.15 | 0.14, 0.13 |
| 100, 110 | 37, 32 | 0.11, 0.13 | 0.17, 0.13 | 0.07, 0.06 |
| 10,000, 9,900 | 397, 327 | 0.14, 0.15 | 0.21, 0.18 | 0.23, 0.21 |
| 10,000, 10,100 | 158, 156 | 0.10, 0.12 | 0.15, 0.17 | 0.15, 0.14 |

The kernel columns are the double-double entry points, which answered every
one of these calls before #314. The Chi-square call is at twice the argument
and twice the shape; the Poisson call is at count `shape - 1` and mean
`argument`. The two figures in bold are the registered exception: at shape
0.0001 the lower tail is above 0.9, so the survival stays with the kernel.

The binary64 routine of before #242, timed alone at the same points, takes
0.004 to 0.077 ms. Most of a public call is therefore not the series or the
continued fraction. Where the rest goes has not been measured.

Inverses, milliseconds per call at probabilities 0.1 and 0.9:

| Shape | Gamma inverse | Chi-square inverse |
| --- | ---: | ---: |
| 0.001 | refused after 1,050; 1,310 | refused after 1,350; 1,340 |
| 0.01 | 0.27, **95** | 0.32, **88** |
| 0.5 | 0.05, 0.15 | 0.17, 0.31 |
| 2.5 | 0.15, 0.21 | 0.17, 0.21 |
| 100 | 1.3, 0.35 | 1.7, 0.48 |
| 10,000 | 21, 21 | 23, 18 |

The Poisson inverse at probability 0.5 takes 0.19 to 0.30 ms at mean about
100 and 0.73 to 1.1 ms at mean about 10,000.

Three things in that table are not improvements, and none is new:

- Shape 0.001 is below the limit of Addendum 1, so the solver uses the
  kernel, at more than a second per call. At probability 0.1 the quantile is
  `0.1^1000`, below the Double range, and the call is refused.
- At shape 0.01 and probability 0.9 the solver works where the lower tail
  exceeds 0.9, the registered exception, and pays the kernel's cost.
- At shape 10,000 an inverse takes about 20 ms, far more than its tail
  evaluations cost through the route. The solver was not profiled here.

The regression of 1,922 assertions takes 5 s on the runner, against 57 s
before the route, and the full export of the main grid 1 s against 3 min
40 s. The two slow production batteries still take about 225 s, because
they call the double-double entry points directly, and stay apart from the
pull-request regression.

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

## A correction after the verdict: the shape ceiling

A review of pull request #347 found that the registration's "no shape
ceiling" was wrong: near `X = A` the route's exponent loses accuracy in
proportion to the shape, and from about 1E18 the tail misses the bound of
1E-13 without the routine declining. No row of the holdout, the sweep or the
main grid is that large, so the verdict above did not see it.

Addendum 2 of the preregistration records the measurement and the
correction: shapes above 1E15 are delegated to the kernel. The verdict above
was computed on `c7042c5`, without the ceiling. No holdout row has a
shape above it, so the corrected candidate must return the same grid bit for
bit. The evidence commit that follows the corrected candidate records that
export, and CI recomputes the verdict on it.

## What stays with the kernel

Shapes above 1E15, by Addendum 2. As registered: an upper tail below `X = A + 1` when the lower tail exceeds
0.9, any directly returned tail below 1E-290, the inverses for shapes below
0.01, and any call the binary64 routine declines. Their accuracy and their
cost are unchanged.
