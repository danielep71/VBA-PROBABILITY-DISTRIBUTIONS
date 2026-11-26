# The incomplete-beta prefactor and the far deviance (#175)

**Status (2026-11-26): accepted.** The first candidate failed rule H2 at two
points (`RESULT_1.md`). `PREREGISTRATION_2.md` kept the rule and replaced how
it is judged, on a fresh holdout; the same source passed rules K1 to K4
(`RESULT_2.md`) and is in `src/`.

| File | Role |
| --- | --- |
| `PREREGISTRATION.md` | the defect, the mechanism, the rule, the rows, rules H1 to H4, predictions |
| `prefactor_model.py` | the rule in two binary64 ports of the source (`x87`, `double`) |
| `prefactor_study.py` | builds the holdout grid, keeps the verdicts of the source of before, judges the export |
| `RESULT_1.md` | the first candidate against its rules and predictions: rejected |
| `wave1_verdicts.json` | the verdicts of the first candidate's export, which was not applied |
| `PREREGISTRATION_2.md` | the second registration: the same rule, rules K1 to K4, a fresh holdout |
| `prefactor_study2.py` | the second holdout and its judge |
| `prefactor2_before.json` | the verdicts of the source of before on the second holdout and the floor grid |
| `RESULT_2.md` | the second wave against K1 to K4 and its predictions: accepted |
| `test_prefactor2.py` | the second design (no point shared with the first) and each rule on made-up verdicts |
| `test_prefactor.py` | the region, the far deviance, the grid against its generator, each rule on made-up verdicts |
| `../excel_function_map/prefactor_holdout_grid.csv` | 317 points, the library and Excel side by side; study export `xfp_prefactor` |

The grid lives beside `M_STATS_PROBDIST_XFM.bas`, whose evaluator fills it,
and has the format of the Excel floor grid of #493.
