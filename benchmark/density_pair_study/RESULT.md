# #501: the Beta, F and Student t densities as pairs of Doubles — result, 2026-11-26

Registration `PREREGISTRATION.md` (`1ff715e`). The verdicts of before were
frozen from Excel run 45842548435 (`1c245b5`, `e704a29`). The candidate is
`84740c8`. Excel run 45848441378 (Excel 16.0 build 20719, 64-bit) compiled
it and passed the regression, 3,067 of 3,067 assertions. The export was
judged on the run's artifact before it was applied, then applied as
`a207de6`.

**Verdict: ACCEPTED.** L1 to L4 hold, and the strict accuracy gate passes on
the main grid and the consolidated holdout with no contract changed.

    python density_study.py --check

## The rules

| Id | Rule | Outcome |
| --- | --- | --- |
| L1 | Every holdout point within 1E-14 | held: worst 7.5E-16 (1.73E-12 before) |
| L2 | No holdout point inferior to Excel; every row a number, no library error | held: none of 114 (39 before) |
| L3 | The ten points of the issue within 1E-14 | held: worst 1.8E-16 |
| L4 | No loss above 3E-14 and no more inferior points in the floor grid, the two holdouts of #175 and the holdout of #498 | held: no newly inferior point; 54 against 59, 21 against 29, 14 against 16, none against none |

By class:

| Class | Points | Worst before | Worst after | Inferior before | After |
| --- | ---: | ---: | ---: | ---: | ---: |
| `Beta_Density` | 51 | 3.8E-13 | 7.5E-16 | 15 | 0 |
| `F_Density` | 38 | 1.73E-12 | 4.7E-16 | 13 | 0 |
| `StudentT_Density` | 25 | 6.9E-13 | 2.2E-16 | 11 | 0 |

`F_Density(6, 300, 900)`, the one point of the issue above 1E-13, goes from
1.3E-13 to 1.5E-17; Excel is at 1.6E-14 there.

## The predictions

| | Prediction | Outcome |
| --- | --- | --- |
| 1 | L1 to L4 hold | held |
| 2 | Worst holdout error below 5E-15; above 5E-13 before | held: 7.5E-16; 1.73E-12 |
| 3 | At least 15 of 114 inferior before; none after | held: 39; none |
| 4 | The ten points within 3E-15 | held: 1.8E-16 |
| 5 | No newly inferior point in the four other grids | held |

The model under the rule returns the three worst figures Excel measured
(7.5E-16, 4.7E-16, 2.2E-16). For the source of before it was less good: its
port of the Student t density gave 6E-14 on the holdout where Excel measured
6.9E-13. The registration had said that port was not faithful.

## What the wave measured elsewhere

- **Large shapes, not predicted.** The pair route also serves the densities
  at shapes up to 1E20, where the contracts allow 1E-5 (Beta) and 5E-5 (F).
  On the main grid:

  | Regime | Rows changed | Worst before | Worst after |
  | --- | ---: | ---: | ---: |
  | `Beta_Density`, large shape | 43 | 6.8E-7 | 4.4E-15 |
  | `F_Density`, large shape | 40 | 1.8E-6 | 3.9E-15 |
  | `F_Density`, extreme ratio | 15 | 7.1E-13 | 1.5E-15 |
  | `StudentT_Density`, large df | 28 | 1.8E-14 | 5.6E-16 |

  The registration made no claim there, and none is made now beyond these
  rows: the contracts are not changed.
- 179 observations of the main grid change, 29 of the consolidated holdout
  and two of `holdout_older`: 210 rows in all, every one of the three
  densities. 175 are closer to the reference and 35 further; none of the 35
  is above 1.1E-14, and the gates pass.
- Five dispositions of the floor grid are no longer needed and are removed:
  four `StudentT_Density` points recorded against #494 and one
  `Beta_Density` point recorded against #495. 54 inferior points remain,
  all disposed: 30 against #494, 16 against #495, one against #496, and the
  seven of the fast normal inverse.
- In the two holdouts of #175 no point is inferior to Excel above 1E-13 any
  more. What remains above 1E-13 anywhere measured is in the floor grid:
  Binomial cumulative and survival at very large shapes (#495, up to
  3.2E-13) and the Hypergeometric point of #496.
- `beta_route_model.py` now ports the Beta density through the pair route;
  it returns Excel's Double at 56 of the 57 Beta density observations it
  evaluates and is two units away at the other.

## Not done here

- The contracts of the three densities are not tightened, the large-shape
  ones included.
- The cost per call of the pair route is not measured.
- `PROB_TryBetaLogPdf` and `PROB_TryStudentTPDF` keep their code, as
  registered: the incomplete beta, the Student t inverse and central mass
  and every cumulative function keep their results.
- No assertion was added to the VBA regression module.
