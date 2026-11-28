# The Student t tail beyond |t| = 16 as pairs of Doubles (#365)

**Status (2026-11-28): accepted.** The rule passed T1 to T6 in Excel
(`RESULT.md`) and is in `src/`: none of the 132 survival and cumulative
points of the holdout is inferior to Excel, against 60 before, and the worst
error went from 4.6E-13 to 9.9E-16.

For df below 1E5 and `|t|` above 16 the tail is half of an incomplete beta
whose prefactor has its logarithm, of the size of `t^2 / 2`, in one Double.
The rule takes the two beta arguments, the logarithm of the prefactor (#501)
and the continued fraction (#35) as pairs of Doubles.

| File | Role |
| --- | --- |
| `PREREGISTRATION.md` | the defect, the mechanism, the rule, the callers, the rows, rules T1 to T6, predictions; Amendment 1 |
| `tt_before.json` | the verdicts of the source of before on the holdout, the rows of the issue and seven other grids |
| `RESULT.md` | the wave against T1 to T6 and the predictions; three departures |
| `tt_model.py` | the tail with and without the rule, in plain binary64 |
| `tt_study.py` | builds the holdout grid, keeps the verdicts of the source of before, judges the export |
| `test_tt.py` | the model at the rows of the issue, each rule on made-up verdicts |
| `../excel_function_map/student_t_far_tail_holdout_grid.csv` | 154 points, the library and Excel side by side; study export `xtt_tail` |
