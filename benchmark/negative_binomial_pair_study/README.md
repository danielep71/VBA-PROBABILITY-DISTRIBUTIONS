# The Negative Binomial and Geometric mass functions as pairs of Doubles (#504)

**Status (2026-11-26): registered, not yet measured.** The candidate does not
exist in VBA. The holdout is exported from the source of today first.

The Negative Binomial mass is a Binomial mass times `r / (r + x)`, with the
Binomial logarithm in one Double today. The rule takes that logarithm and the
ratio as pairs of Doubles, on the pair arithmetic of #498 and the Binomial
logarithm of #496. The Geometric mass takes the same treatment only if the
export of before shows it inferior to Excel at a holdout point.

| File | Role |
| --- | --- |
| `PREREGISTRATION.md` | the defect, the mechanism, the rule and its condition, the callers, the rows, rules N1 to N5, predictions |
| `nb_model.py` | the two masses with and without the rule, in plain binary64 |
| `nb_study.py` | builds the holdout grid, keeps the verdicts of the source of before, judges the export |
| `test_nb.py` | the model at the points of the issue, each rule on made-up verdicts |
| `../excel_function_map/negative_binomial_pair_holdout_grid.csv` | 100 points, the library and Excel side by side; study export `xnb_pair` |
