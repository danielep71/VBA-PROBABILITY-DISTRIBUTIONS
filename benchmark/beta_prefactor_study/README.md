# The incomplete-beta prefactor and the far deviance (#175)

**Status (2026-11-26): registered.** The rule and its acceptance rules are in
`PREREGISTRATION.md`. No production source has changed yet.

| File | Role |
| --- | --- |
| `PREREGISTRATION.md` | the defect, the mechanism, the rule, the rows, rules H1 to H4, predictions |
| `prefactor_model.py` | the rule in two binary64 ports of the source (`x87`, `double`) |
| `prefactor_study.py` | builds the holdout grid, keeps the verdicts of the source of before, judges the export |
| `test_prefactor.py` | the region, the far deviance, the grid against its generator, each rule on made-up verdicts |
| `../excel_function_map/prefactor_holdout_grid.csv` | 317 points, the library and Excel side by side; study export `xfp_prefactor` |

The grid lives beside `M_STATS_PROBDIST_XFM.bas`, whose evaluator fills it,
and has the format of the Excel floor grid of #493.
