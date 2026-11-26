# #504: the Negative Binomial and Geometric mass functions as pairs of Doubles — result, 2026-11-26

Registration `PREREGISTRATION.md` (`82aab5d`). The verdicts of before were
frozen from Excel run 45883720628 (`18df981`, `52c048c`). The candidate is
`d2a92cb`, after a commit that only lays out the two procedures it touches.
Excel run 45888759918 (Excel 16.0 build 20719, 64-bit) compiled it and
passed the regression, 3,067 of 3,067 assertions. The export was judged on
the run's artifact before it was applied, then applied as `fe2cd56`.

**Verdict: ACCEPTED.** N1 to N5 hold, and the strict accuracy gate passes on
the main grid and the consolidated holdout with no contract changed. There is
no departure from the registration.

    python nb_study.py --check

## The condition of R2

The registration wrote the Geometric rule to enter the candidate if and only
if a `ge` point of the holdout was inferior to Excel in the export of before.
16 of the 48 were, two of them above 1E-13. R2 applies, the candidate carries
it, and the `ge` points are judged by N1 and N2. N5 has nothing to forbid.

## The rules

| Id | Rule | Outcome |
| --- | --- | --- |
| N1 | Every judged point within 1E-14 | held: `nb` worst 8.7E-16 (3.9E-13 before); `ge` worst 4.2E-16 (1.2E-13 before) |
| N2 | No judged point inferior to Excel; every row a number, no library error | held: none of 52 `nb` (10 before), none of 48 `ge` (16 before) |
| N3 | The eight points of the issue within 1E-14 | held: worst 8.8E-16 (6.1E-13 before) |
| N4 | No loss above 3E-14 and no more inferior points in the floor grid and the five earlier holdouts | held: no newly inferior point anywhere; the holdout of #496 goes from 8 inferior to none |
| N5 | If R2 does not apply, no `ge` observation changes | not engaged: R2 applies |

## The predictions

| | Prediction | Outcome |
| --- | --- | --- |
| 1 | N1 to N5 hold | held |
| 2 | Worst `nb` error below 1E-14; above 3E-13 before | held: 8.7E-16; 3.9E-13 |
| 3 | At least 8 of 52 `nb` inferior before; none after | held: 10; none |
| 4 | The eight points of the issue within 3E-15 | held: 8.8E-16 |
| 5 | R2 applies | held: 16 `ge` points inferior before |
| 6 | Outside the two mass functions no observation changes | held: see below |

## What changed, grid by grid

Every changed row of every exported grid is a `NegativeBinomial_PMF` or a
`Geometric_PMF` row.

| Grid | Rows changed | Of which |
| --- | ---: | --- |
| holdout of #504 | 96 of 100 | 51 `nb`, 45 `ge`; 91 improved, 5 rose, by at most 2.8E-16 |
| floor grid of #493 | 34 | 25 `NegativeBinomial_PMF`, 9 `Geometric_PMF`; 6 rose, by at most 3.4E-16 |
| holdout of #496 | 27 | all 27 `NegativeBinomial_PMF`; one rose, by 1.3E-16 |
| main grid | 29 | 22 `NegativeBinomial_PMF`, 7 `Geometric_PMF` |
| consolidated holdout | 5 | 3 `NegativeBinomial_PMF`, 2 `Geometric_PMF` |
| holdouts of #175, #498, #501; older holdout; survival boundary; large-df grid | 0 | |

The rises are of a fraction of one unit in the last place, at points that
were and remain below 1E-15.

## Against Excel

On the 52 `nb` points `NEGBINOM.DIST` is off by up to 5.9E-13, and above
1E-13 at 17; the library is now the closer of the two at 49. On the 48 `ge`
points Excel's worst is 9.4E-14, and the library is the closer at 43.

## What the model got right and wrong

The model put the rule within 1.4E-15 (`nb`) and 5.0E-16 (`ge`) on the
holdout; Excel measured 8.7E-16 and 4.2E-16. Its figures for the source of
before were not faithful, as the registration said: it gave 6.2E-13 for the
worst `nb` point where Excel measured 3.9E-13.

## Not done here

- The counts `x = 0` and `k = 0`, the two LogPMF functions, and the
  cumulative, survival and inverse functions are as they were.
- No contract was tightened.
- The cost per call of the pair routes is not measured; it is measured
  before the release.
