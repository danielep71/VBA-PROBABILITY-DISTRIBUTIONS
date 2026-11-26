# The Hypergeometric mass function as pairs of Doubles (#496)

**Status (2026-11-26): accepted.** The rule
passed M1 to M4 in Excel (`RESULT.md`) and is in `src/`: no point of the 69
Hypergeometric points of the holdout is inferior to Excel, against 10 before.
The registration said the cumulative and survival functions would keep their
results; they take their anchor mass from the changed procedure and 14 of
their observations moved, by at most 3.2E-16 for the worse. The maintainer
accepted the candidate with that scope; `RESULT.md` records it.

The mass is a product of two Binomial masses over a third, at `p = n / N`.
Each logarithm is one Double today, and `p` is rounded before any of them
sees it. The rule takes `p` and the three logarithms as pairs of Doubles, on
the pair arithmetic of #498.

| File | Role |
| --- | --- |
| `PREREGISTRATION.md` | the defect, the mechanism, the rule, the rows, rules M1 to M4, predictions |
| `hyper_before.json` | the verdicts of the source of before on the holdout and on five other grids |
| `RESULT.md` | the wave against M1 to M4 and the predictions, and the departure |
| `hyper_model.py` | the mass with and without the rule, in plain binary64 |
| `hyper_study.py` | builds the holdout grid, keeps the verdicts of the source of before, judges the export |
| `test_hyper.py` | the model at the point of the issue and at an awkward ratio, each rule on made-up verdicts |
| `../excel_function_map/hypergeometric_pair_holdout_grid.csv` | 96 points, the library and Excel side by side; study export `xhg_pair` |
