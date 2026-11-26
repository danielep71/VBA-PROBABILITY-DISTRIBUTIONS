# The Negative Binomial and Geometric mass functions as pairs of Doubles (#504)

**Status (2026-11-26): accepted.** The rule passed N1 to N5 in Excel
(`RESULT.md`) and is in `src/`: none of the 52 Negative Binomial points of
the holdout is inferior to Excel, against 10 before, and none of the 48
Geometric points, against 16. The Geometric rule was conditional on that
function being inferior to Excel before; it was.

The Negative Binomial mass is a Binomial mass times `r / (r + x)`, with the
Binomial logarithm in one Double today. The rule takes that logarithm and the
ratio as pairs of Doubles, on the pair arithmetic of #498 and the Binomial
logarithm of #496. The Geometric mass, `p (1 - p)^k`, takes `k Log(1 - p)`
as a pair.

| File | Role |
| --- | --- |
| `PREREGISTRATION.md` | the defect, the mechanism, the rule and its condition, the callers, the rows, rules N1 to N5, predictions |
| `nb_before.json` | the verdicts of the source of before on the holdout and on six other grids |
| `RESULT.md` | the wave against N1 to N5 and the predictions |
| `nb_model.py` | the two masses with and without the rule, in plain binary64 |
| `nb_study.py` | builds the holdout grid, keeps the verdicts of the source of before, judges the export |
| `test_nb.py` | the model at the points of the issue, each rule on made-up verdicts |
| `../excel_function_map/negative_binomial_pair_holdout_grid.csv` | 100 points, the library and Excel side by side; study export `xnb_pair` |
