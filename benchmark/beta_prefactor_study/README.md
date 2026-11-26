# The incomplete-beta prefactor and the far deviance (#175)

**Status (2026-11-26): the first candidate is rejected.** It failed rule H2
at two points (`RESULT_1.md`); the source is back to the state of before the
rule, and a second registration follows.

| File | Role |
| --- | --- |
| `PREREGISTRATION.md` | the defect, the mechanism, the rule, the rows, rules H1 to H4, predictions |
| `prefactor_model.py` | the rule in two binary64 ports of the source (`x87`, `double`) |
| `prefactor_study.py` | builds the holdout grid, keeps the verdicts of the source of before, judges the export |
| `RESULT_1.md` | the first candidate against its rules and predictions: rejected |
| `wave1_verdicts.json` | the verdicts of the first candidate's export, which was not applied |
| `test_prefactor.py` | the region, the far deviance, the grid against its generator, each rule on made-up verdicts |
| `../excel_function_map/prefactor_holdout_grid.csv` | 317 points, the library and Excel side by side; study export `xfp_prefactor` |

The grid lives beside `M_STATS_PROBDIST_XFM.bas`, whose evaluator fills it,
and has the format of the Excel floor grid of #493.
