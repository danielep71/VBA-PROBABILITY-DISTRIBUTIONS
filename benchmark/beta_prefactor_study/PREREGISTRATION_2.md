# #175: the incomplete-beta prefactor — second registration

Registered 2026-11-26 (Europe/Rome), after the first candidate was rejected
(`RESULT_1.md`) and before any point of the new holdout is evaluated. The 329
points are committed in `../excel_function_map/prefactor_holdout2_grid.csv`
with references only. No production source changes in the commit that carries
this file: the source is that of before the rule.

## What is the same, and what changes

**The rule is unchanged.** R1 and R2 of `PREREGISTRATION.md`, with its
Amendment 1: the Loader prefactor whenever both shapes are at least 1/2,
decided by `PROB_IBetaUsesLoader` for `PROB_TryBetaRegularized` and
`PROB_TryBetaTailsPair`; one logarithm of a ratio in the far branch of
`PROB_TryDeviancePart`, with the two logarithms kept outside 1E-100 to
1E+100. The candidate is the source of `dec7696`, restored unchanged.

**The region is unchanged, on evidence.** The first result suggested that the
rule loses where the shapes are very unequal. The model does not support
narrowing it. On the 2235 development points, by the ratio of the smaller
shape to the larger (the `double` port; "better" and "worse" count a change
of more than 5E-15 that ends or starts above 1E-14):

| Ratio of the shapes | Points | Above 1E-14 before | Above 1E-14 under the rule | Better | Worse |
| --- | ---: | ---: | ---: | ---: | ---: |
| below 0.01 | 448 | 144 | 56 | 114 | 15 |
| 0.01 to 0.03 | 324 | 111 | 15 | 108 | 5 |
| 0.03 to 0.1 | 432 | 186 | 10 | 182 | 2 |
| 0.1 to 0.3 | 484 | 251 | 21 | 250 | 0 |
| 0.3 to 1 | 547 | 260 | 27 | 258 | 2 |

The rule helps at every ratio. At the level of 1E-14 both forms carry a few
units of rounding, and a point where the old form happened to be exact can
come out a little worse: 24 such points of 2235 in the model, 2 of 317 in the
first wave. No region removes them.

**The rules of judgement change.** H2 of the first registration counted any
point that crossed the 1E-14 floor, by however little, as a rejection. The
first wave had two: 1.26E-14 and 1.66E-14. That rule could not tell rounding
from a loss, and it is replaced here, before new data, by one that can:

- a newly inferior point is a **loss** when the library's error there is
  above 3E-14, three times the floor and the bound H4 already used for the
  points of the issue;
- newly inferior points below that are allowed up to 2 percent of the
  holdout, and each is reported;
- the count of inferior points must at least halve.

These thresholds are informed by the first wave (0.6 percent newly inferior,
worst 1.66E-14, the count down to 22 percent) and are set with margin on it.
They are judged on points that wave did not reveal.

## The rows

`prefactor_study2.py --build` writes 329 points, 440 rows, in the format of
the Excel floor grid. No point is in the first holdout
(`test_prefactor2.py` checks it).

| Points | Class | What they hold |
| ---: | --- | --- |
| 112 | `ib` | `Beta_Cumulative` and `Beta_Survival` at eleven new shape pairs with a + b below 1000, from -6 to +7 standard deviations |
| 40 | `ib` | `NegativeBinomial_Cumulative` and `Survival` at four new `(s, p)` |
| 40 | `ib` | `Binomial_Cumulative` and `Survival` at four new `(n, p)` |
| 48 | `ib` | `F_Cumulative` and `F_Survival` at four new pairs of degrees of freedom |
| 35 | `ib` | `StudentT_Survival` at df from 9 to 1700 and t from 0.6 to 14 |
| 54 | `dv` | `Binomial_PMF`, `Poisson_PMF`, `Gamma_Density`, `Beta_Density`, `F_Density`, `StudentT_Density`, far from the mean |

The grid is exported once from the source of before, and
`prefactor_study2.py --freeze-before` keeps those verdicts and the floor
grid's in `prefactor2_before.json`. Then the candidate is restored and
exported.

## Acceptance rules

`prefactor_study2.py --check` implements them on the exported grids.

| Id | Rule |
| --- | --- |
| K1 | Every `ib` point is within **2E-13** of the reference |
| K2 | No holdout point is newly inferior to Excel with a library error above **3E-14**; newly inferior points are at most **2 percent** of the holdout; at most **half** as many points are inferior as before |
| K3 | Every row holds an observation, and no library value is an error |
| K4 | Each of the twelve points of the issue is within **3E-14** of the reference; the floor grid of #493 has no newly inferior point with a library error above **3E-14**, and fewer inferior points than before |

Also required in the same wave: the full Excel regression, and the strict
accuracy gate on the main grid and both holdouts with no contract changed.

Reported, not gating: every newly inferior point with its figures; the `dv`
points before and after.

## Development evidence

The model on the new holdout's `ib` points (`prefactor_study2.py --model`):
before, worst 1.03E-12 (`x87`) and 1.33E-12 (`double`); under the rule,
5.4E-14 and 5.3E-14, with 3 and 4 points above 3E-14.

The first wave is development evidence too, and is revealed: on its 317
points the candidate brought the inferior points from 173 to 38 and the
worst `ib` error from 5.02E-13 to 9.93E-14, with the two points above.

## Predictions

1. K1 to K4 hold.
2. The worst `ib` error is below 1.5E-13; before the rule it is above 5E-13.
3. At most a third as many holdout points are inferior as before.
4. Each of the twelve points of the issue is within 1E-14.
5. The worst `dv` error falls, and no `dv` point rises by more than 5E-14.

No prediction is made on how many Student t points remain inferior to
`T.DIST.RT`: the first registration's was wrong.

## After acceptance

- Newly inferior points below 3E-14 are recorded against #494 in
  `floor_dispositions.json`; dispositions of points that stop being inferior
  are removed.
- `below1e5_study.py --check` (#365) holds, by its rule G2, the rows outside
  the rule of #365 to their export of 2026-11-25. A prefactor rule moves
  those rows. On acceptance `below1e5_before.json` is frozen again from the
  accepted export, and `BELOW1E5_RESULT.md` records it.
- The balanced Beta contracts are tightened to what Excel measures, and the
  header of `PROB_TryBetaRegularized` is corrected, in separate commits.

## Rejection policy

A failure of K1 to K4, of the regression or of the gate rejects the rule. The
bounds are not retuned and the result is recorded.
