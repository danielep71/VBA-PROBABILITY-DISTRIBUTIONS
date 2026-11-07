# Post-#23 cutoff revalidation

Excel run [428](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/actions/runs/42421734701)
exported all four frozen arms from `f157e67ce92f1ac174d5b8a2ae32fabd10553a90`.
Excel 16.0 build 20430, 64-bit; full regression: 955/955 PASS.
The grids contain direct and log-recovery observations, including explicit
ERROR outcomes; these are evaluated by the existing analyzer, never replaced
with zero. No production source or accuracy threshold changed.

| Surface | Fitting crossover | Frozen holdout claim | Worst candidate relative error |
|---|---:|---|---:|
| Gamma density | 48 bits | PASS through 48 bits | 1.49E-13 |
| Gamma CDF | 40 bits | PASS through 40 bits | 1.49E-13 |
| Chi-square density | 48 bits | PASS through 48 bits | 1.31E-13 |
| Chi-square CDF | 48 bits | PASS through 48 bits | 1.33E-13 |

Retain the frozen 48/40/48/48-bit policies. Survival has no dispatch claim.
Gamma X/ScaleParam and Chi-square 0.5*X remain separate policies. Production
wiring and its fresh validation remain #13's work.

The two fitting arms contain 1235 and 627 rows; the two holdouts contain 1197
and 609 rows. All construction checks and all four frozen claims pass.
These historically exposed holdouts are the required post-fix replay, not a
new untouched confirmation set. The cutoffs were not tuned to these results.

Reproduce from the committed CSVs:

```sh
python benchmark/positive_ratio_subnormal/check_post23_cutoffs.py
```

`post23_cutoffs.json` binds the fitting report to its CSV bytes;
`POST23_CLAIMS.txt` retains the analyzer's four holdout claims. The original
pre-fix observations remain in Git history. The export artifact zip digest is
`ae13faafc4d094dc116092fb63649b47a8e08e3c22e2c18ce75387b13f1341d9`.
