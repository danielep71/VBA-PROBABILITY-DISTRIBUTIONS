# #175: the incomplete-beta prefactor — result of the first candidate, 2026-11-26

Registration `c5c84e1`, Amendment 1 and candidate `dec7696`. Excel run
45776476298 (Excel 16.0 build 20719, 64-bit); the regression passed 3,067 of
3,067 assertions. The export of the candidate was judged on the run's
artifact and was **not** applied to the repository; `wave1_verdicts.json`
keeps its verdicts.

**Verdict: REJECTED.** H2 fails at two points. The bounds are not retuned.
The source is restored to the state of before the rule (the house-style
layout of the three procedures stays), and a second registration follows.

## The rules

| Id | Rule | Outcome |
| --- | --- | --- |
| H1 | Every `ib` point within 2E-13 | held: worst 9.93E-14 (5.02E-13 before) |
| H2 | No holdout point newly inferior to Excel, and no more inferior than before | **failed**: 38 inferior against 173 before, but two of the 38 were not inferior before |
| H3 | Every row holds an observation, no library error | held |
| H4 | The twelve points of the issue within 3E-14; no newly inferior point in the floor grid | held: worst 5.1E-15; 61 inferior against 123, none new |

The two points of H2:

| Point | Before | After | Excel |
| --- | ---: | ---: | ---: |
| `Beta_Cumulative(0.993276, 700, 3)` | 7.6E-17 | 1.66E-14 | 7.6E-17 |
| `Beta_Density(0.812111, 150, 90)` | 5.2E-15 | 1.26E-14 | 2.5E-16 |

The first is a loss of 1.7E-14 at a point where the literal form happened to
be correctly rounded. The second is at the level of rounding: the far
deviance moved by 7E-15.

*Correction, 2026-11-26.* This section first said that with one shape small
and the other large the literal form is exact, and that the region of R1 was
therefore too wide. The model does not support that: for shapes with a ratio
below 0.1 the rule is better at 404 development points and worse at 22
(`PREREGISTRATION_2.md`). The loss at this point is the scatter of two forms
that each carry a few units of rounding, not a region where the old form
wins.

## The predictions

| # | Prediction | Outcome |
| --- | --- | --- |
| 1 | H1 to H4 hold | **failed** (H2) |
| 2 | Worst `ib` error below 1E-13; above 2E-13 before | held: 9.93E-14; 5.02E-13 |
| 3 | Each of the twelve points within 1E-14 | held: worst 5.1E-15 |
| 4 | Fewer inferior points in the holdout and in the floor grid | held: 173 to 38; 123 to 61 |
| 5 | No `dv` point worse by more than 1E-14; their worst does not rise | **failed**: `Binomial_PMF(1466, 200000, 0.01)` went from 2.7E-14 to 5.8E-14. The worst fell from 4.94E-12 to 2.14E-13, and 28 of 54 improved against 3 worse |
| 6 | Student t with df below 1999 and `t` up to 16 no longer inferior to `T.DIST.RT` | **failed** as stated: 40 of 140 points inferior against 61. With df below 1999 they fell from 53 to 31, and 17 of the 31 have `t` up to 16 |

## What the wave showed

By function, inferior points of the holdout before and after: Beta cumulative
27 to 4, Beta survival 29 to 1, Binomial cumulative 9 to 1 and survival 11 to
1, Negative binomial cumulative 11 to 0 and survival 12 to 2, F cumulative 10
to 2 and survival 15 to 1, Student t survival 18 to 9, Gamma density 6 to 1,
Poisson PMF 6 to 1, Binomial PMF 10 to 7, Beta density 3 to 2, F density 3
to 3, Student t density 3 to 3.

The x87 port of the model gave 5.02E-13 for the worst `ib` point before the
rule, and Excel measured 5.02E-13. For the candidate it gave 8.4E-14 where
Excel measured 9.93E-14.

The candidate changed 191 observations of the main grid, 65 of the
consolidated holdout and 26 of the older-families holdout. Whether every
contract held was not established: the export was not applied.

`below1e5_study.py --check` (#365) rejects the candidate's export by its rule
G2, which requires the rows outside the rule of #365 to be bit-identical to
their export of 2026-11-25. Any change to the prefactor moves those rows, so
that frozen comparison has to be re-based when a prefactor rule is adopted.

## What the second registration takes from this

- The region of R1 stays: the model finds no part of it where the old form
  is better on the whole.
- H2 counted a point crossing 1E-14 by a few E-15 the same as a real loss.
  A criterion that separates the two has to be written before new data are
  seen, and judged on points this wave has not revealed.
- Predictions 5 and 6 were stated more strongly than the model supported.
