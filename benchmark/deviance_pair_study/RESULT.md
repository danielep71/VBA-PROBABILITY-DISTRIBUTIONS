# #498: the Loader deviance as a pair of Doubles — result, 2026-11-26

Registration `PREREGISTRATION.md` (`87a4a9e`). The verdicts of before were
frozen from Excel run 45815867021 (`5bdfba4`, `bc06901`). The candidate is
`e9fc3b7`. Excel run 45822509492 (Excel 16.0 build 20719, 64-bit) compiled
it and passed the regression, 3,067 of 3,067 assertions. The export was
judged on the run's artifact before it was applied, then applied as
`452125b`.

**Verdict: ACCEPTED.** J1 to J4 hold, and the strict accuracy gate passes on
the main grid and the consolidated holdout with no contract changed.

    python pair_study.py --check

## The rules

| Id | Rule | Outcome |
| --- | --- | --- |
| J1 | Every holdout point within 1E-14 | held: worst 1.21E-15 (1.56E-12 before) |
| J2 | No holdout point inferior to Excel; every row a number, no library error | held: none of 180 (57 before) |
| J3 | The five points of the issue within 1E-14 | held: worst 7.0E-16 |
| J4 | No loss above 3E-14 and no more inferior points in the floor grid and the two holdouts of #175 | held: no newly inferior point; 59 against 61, 29 against 38, 16 against 26 |

By class:

| Class | Points | Worst before | Worst after | Inferior before | After |
| --- | ---: | ---: | ---: | ---: | ---: |
| `Binomial_PMF` | 46 | 1.56E-12 | 1.2E-15 | 17 | 0 |
| `Gamma_Density` | 56 | 7.7E-13 | 1.0E-15 | 18 | 0 |
| `ChiSquare_Density` | 27 | 5.3E-13 | 1.0E-15 | 9 | 0 |
| `Poisson_PMF` | 51 | 4.8E-13 | 4.7E-16 | 13 | 0 |

The five points of the issue:

| Call | Before | After | Excel |
| --- | ---: | ---: | ---: |
| `Poisson_PMF(5600, 4000)` | 4.4E-13 | 4.9E-16 | 5.1E-14 |
| `Gamma_Density(7500, 5000, 1)` | 3.0E-13 | 2.1E-16 | 8.6E-14 |
| `Gamma_Density(3500, 5000, 1)` | 1.9E-13 | 5.5E-16 | 5.9E-14 |
| `Binomial_PMF(4076, 8000, 0.4)` | 1.9E-13 | 6.8E-17 | 3.3E-14 |
| `Binomial_PMF(4466, 300000, 0.02)` | 1.1E-13 | 7.0E-16 | 5.4E-14 |

## The predictions

| | Prediction | Outcome |
| --- | --- | --- |
| 1 | J1 to J4 hold | held |
| 2 | Worst holdout error below 5E-15; above 3E-13 before | held: 1.21E-15; 1.56E-12 |
| 3 | At least 20 of 180 inferior before; none after | held: 57; none |
| 4 | The five points within 3E-15 | held: 7.0E-16 |
| 5 | Fewer inferior points in the floor grid and the two holdouts of #175, none new | held |

The model gave 1.8E-15 as the rule's worst on the holdout and 1.5E-12 as
today's; Excel measured 1.21E-15 and 1.56E-12.

## What the wave measured elsewhere

- Excel itself is off by up to 2.2E-12 on these points
  (`BINOM.DIST(523862, 900000, 0.6, FALSE)`); the library is within 1.2E-15
  there.
- 72 observations of the main grid change, 11 of the consolidated holdout
  and one of `holdout_older`; the accuracy gates pass with no contract
  changed.
- Two dispositions of the floor grid are no longer needed and are removed,
  both `Binomial_PMF` points recorded against #494. 59 inferior points
  remain, all disposed: 34 against #494, 17 against #495, one against #496,
  and the seven of the fast normal inverse.
- No point of the four functions of this rule is inferior to Excel any more
  in the two holdouts of #175. The second keeps 16 inferior points, the
  worst 7.6E-14; the first keeps 29, and one of them is above 1E-13:
  `F_Density(6, 300, 900)` at 1.3E-13. It is on the Beta log-density, which
  this rule does not touch. The registration looked at the second holdout
  only when it left the F density out; under the floor rule this point is
  open.

## An error of this wave, and its repair

The wave was dispatched with six study exports it should not have had, the
positive-ratio grids of #13 and #23 (`prs_gamma_probe` and five more). They
are frozen evidence: `positive_ratio_subnormal/post23_cutoffs.json` binds
each file by digest. The export rewrote them, `check_post23_cutoffs.py`
failed, and the six files were restored byte for byte (`0688da0`). The
evidence commit lists them as exported; that is true of the run and not of
the tree.

## After the wave

- The headers written with the rule held four measured figures;
  `check_source_thresholds.py` allows none in the source. They are removed.
  The change is to comments only: `vba_code_identity.py --from e9fc3b7`
  reports the module as carried, so the certification of the wave is carried
  under #301 and is not that of the exact bytes. A release needs a wave on
  the final source.

## Not done here

- The contracts of the four functions are not tightened.
- The cost per call of the pair route is not measured.
- The Beta, F and Student t densities, the Hypergeometric and Negative
  Binomial mass functions and the incomplete-beta prefactor do not take the
  pair. Whether they should is decided on measurements, as separate work.
