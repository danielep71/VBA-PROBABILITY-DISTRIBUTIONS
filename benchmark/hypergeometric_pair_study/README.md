# The Hypergeometric mass function as pairs of Doubles (#496)

**Status (2026-11-26): registered.** `PREREGISTRATION.md` fixes the rule, the
holdout and the rules of judgement. No file under `src/` changes until the
holdout has been exported from the source of today.

The mass is a product of two Binomial masses over a third, at `p = n / N`.
Each logarithm is one Double today, and `p` is rounded before any of them
sees it. The rule takes `p` and the three logarithms as pairs of Doubles, on
the pair arithmetic of #498.

| File | Role |
| --- | --- |
| `PREREGISTRATION.md` | the defect, the mechanism, the rule, the rows, rules M1 to M4, predictions |
| `hyper_model.py` | the mass with and without the rule, in plain binary64 |
| `hyper_study.py` | builds the holdout grid, keeps the verdicts of the source of before, judges the export |
| `test_hyper.py` | the model at the point of the issue and at an awkward ratio, each rule on made-up verdicts |
| `../excel_function_map/hypergeometric_pair_holdout_grid.csv` | 96 points, the library and Excel side by side; study export `xhg_pair` |
