# #496: the Hypergeometric mass function as pairs of Doubles — result, 2026-11-26

Registration `PREREGISTRATION.md` (`1b71dd1`). The verdicts of before were
frozen from Excel run 45870098635 (`b109c7a`, `a56af69`). The candidate is
`db24a3d`. Excel run 45873953950 (Excel 16.0 build 20719, 64-bit) compiled
it and passed the regression, 3,067 of 3,067 assertions. The export was
judged on the run's artifact before it was applied, then applied as
`8ff608e`.

**Verdict: ACCEPTED by its rules, with one departure from the registration
that the maintainer decides on.** M1 to M4 hold, and the strict accuracy gate
passes on the main grid and the consolidated holdout with no contract
changed. The departure is below.

    python hyper_study.py --check

## The rules

| Id | Rule | Outcome |
| --- | --- | --- |
| M1 | Every `hg` point within 1E-14 | held: worst 2.0E-15 (6.6E-13 before) |
| M2 | No `hg` point inferior to Excel; every row a number, no library error | held: none of 69 (10 before) |
| M3 | The point of the issue within 1E-14 | held: 1.0E-15 (3.5E-14 before; Excel 1.9E-14) |
| M4 | No loss above 3E-14 and no more inferior points in the floor grid and the four other holdouts | held: no newly inferior point; 53 against 54 in the floor grid, the holdouts unchanged |

## The predictions

| | Prediction | Outcome |
| --- | --- | --- |
| 1 | M1 to M4 hold | held |
| 2 | Worst `hg` error below 1E-14; above 3E-13 before | held: 2.0E-15; 6.6E-13 |
| 3 | At least 10 of 69 `hg` inferior before; none after | held, at its bound: exactly 10; none |
| 4 | The point of the issue within 3E-15 | held: 1.0E-15 |
| 5 | No observation of the holdout of #498 changes | held: none of 180 |

Prediction 5 is what shows that splitting `PROB_TryBinomialLogPair` out of
`PROB_TryBinomialMassPair` moved code and not arithmetic.

## A departure from the registration

The registration said: "`PROB_DS_TryHypergeometricLogMass` keeps its code, so
`K_STATS_Hypergeometric_LogPMF` and the cumulative and survival functions,
whatever they call, keep their results."

That is wrong for the cumulative and survival functions.
`PROB_DS_TryHypergeometricLowerTail` and `PROB_DS_TryHypergeometricUpperTail`
take their anchor mass from `PROB_DS_TryHypergeometricPMF`, the procedure
the rule changes, and sum from it. The search for callers that preceded the
registration missed both lines; it was made carelessly, and the claim was
written without being checked.

What the wave measured on those functions:

| Grid | Function | Rows changed | Worst before | Worst after |
| --- | --- | ---: | ---: | ---: |
| main | `Hypergeometric_Survival` | 9 | 9.6E-16 | 7.1E-16 |
| main | `Hypergeometric_Cumulative` | 1 | 1.8E-16 | 2.0E-16 |
| holdout | `Hypergeometric_Survival` | 2 | 1.6E-14 | 1.8E-16 |
| holdout | `Hypergeometric_Cumulative` | 2 | 8.5E-16 | 7.4E-16 |

Fourteen observations change. Eight are closer to the reference and six
further, none by more than 3.2E-16, which is about one unit of a Double; the
gates pass. No rule of the
registration judges these functions apart from M4, which holds.

The two ways out, for the maintainer:

- keep the candidate, with the scope it really has: the mass function, and
  through their anchor the cumulative and survival functions;
- restrict the rule to `K_STATS_Hypergeometric_PMF`, as registered, which
  moves the call out of `PROB_DS_TryHypergeometricPMF` and needs one more
  wave.

`K_STATS_Hypergeometric_LogPMF` is not changed, as registered.

## What the wave measured elsewhere

- 17 observations of the main grid change and 7 of the consolidated holdout:
  the fourteen above and 10 of `Hypergeometric_PMF`, whose worst falls from
  3.5E-14 to 1.1E-15.
- Excel is the less accurate at many `hg` points: up to 1.1E-12, and at
  `HYPGEOM.DIST(1341, 3333, 12345, 77777, FALSE)` its relative error is 1.
- The disposition of the point of the issue is removed from
  `floor_dispositions.json`. 53 inferior points remain in the floor grid,
  all disposed: 30 against #494, 16 against #495, and the seven of the fast
  normal inverse.
- The 27 `nb` points read as before, as they must: the rule does not touch
  the Negative Binomial mass. Eight are inferior to Excel, three above
  1E-13 (#504).

## Not done here

- The contract of the function is not tightened. The cost per call is not
  measured.
- The edges of the Binomial terms keep the route of before.
- No assertion was added to the VBA regression module.
