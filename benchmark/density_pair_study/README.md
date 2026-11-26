# The Beta, F and Student t densities as pairs of Doubles (#501)

**Status (2026-11-26): accepted.** The rule passed L1 to L4 in Excel
(`RESULT.md`) and is in `src/`: no point of the 114 of the holdout is
inferior to Excel, against 39 before.

After #498 ten points of these densities are still less accurate than Excel,
one above 1E-13: `F_Density(6, 300, 900)`. The causes differ by density: the
Loader deviances of the Beta log-density in one Double; in the F density also
its arguments, formed through an exponential; in the Student t a product and
a log-gamma constant. The rule carries each as a pair of Doubles, on the pair
arithmetic of #498.

| File | Role |
| --- | --- |
| `PREREGISTRATION.md` | the defect, the three mechanisms, the rule, the rows, rules L1 to L4, predictions |
| `density_before.json` | the verdicts of the source of before on the holdout and on four other grids |
| `RESULT.md` | the wave against L1 to L4 and the predictions: accepted |
| `density_model.py` | the three densities with and without the rule, in plain binary64 |
| `density_study.py` | builds the holdout grid, keeps the verdicts of the source of before, judges the export |
| `test_density.py` | the model at the points of the issue and at the edges, each rule on made-up verdicts |
| `../excel_function_map/density_pair_holdout_grid.csv` | 114 points, the library and Excel side by side; study export `xdn_pair` |
