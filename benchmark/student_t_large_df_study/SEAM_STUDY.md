# Student-t behaviour across the large-df seams (#210)

The large-df expansion serves `df >= 1E5` and `|t| <= 8`. The incomplete-beta
continued fraction serves everything else. This study measures, in Excel, what
happens at the Doubles on both sides of each seam. #34 could not establish this
before closing.

## The rows

`seam_study.py --append` added 130 rows to `tchi_large_df/tchi_large_df_grid.csv`,
with regime `seam`:

- **df seam.** For `t = 0.5, 1, 2, 4, 8`: the survival and cumulative at both
  signs, at the Double below `df = 1E5`, at `1E5`, and at the Double above it.
- **t seam.** For `df = 1E5, 1E6, 1E7, 1E8`: the same at the Double below
  `|t| = 8`, at 8, and at the Double above it.
- **Round trips.** At each of those 26 points, the inverse at the correctly
  rounded lower-tail cumulative `p = C(-t)`.

The references come from the study's oracle at 80 digits, and pass
`check_survival_references.py`. The rows were exported once, with the regression
passing 955/955 (Excel VBA Regression run 42281403598, candidate `658fd7b`).
`seam_study.py --check` writes `seam_results.json`.

## Results

| Property | Result |
| --- | --- |
| Complement: survival(t) + cumulative(t) − 1 | within **0.5 × 2^-53** at all 52 pairs |
| Symmetry: survival(t) = cumulative(−t) | **bit-identical** at all 52 pairs |
| Monotonicity across df = 1E5 | breaks by up to **3.16E-12** relative (t = 2), 5.97E-13 (t = 4), 4.19E-13 (t = 8) |
| Monotonicity across \|t\| = 8 | breaks by **8.1E-11** relative at df = 1E7; monotone at df = 1E5, 1E6, 1E8 |
| Monotonicity inside the region | within 5.4E-16, the expansion's own rounding |
| Round trips, df ≥ 1E5 and \|q\| ≤ 8 | within 4.4E-16, **except at df = 1E7 near q = −8: 1.29E-12** |
| Round trips just outside (df below 1E5) | up to 8.9E-13 (t = −2) |

Each monotonicity break comes from the continued fraction on the outside of
its seam. Its error there, up to 3.2E-12 below `df = 1E5` and 8.1E-11 above
`|t| = 8` at `df = 1E7`, is far larger than the true change across one Double.

## Consequences

- **The monotonicity breaks** are recorded as a characterized limitation,
  `StudentT.LargeDFSeamMonotonicity` in `numerical_limitations.csv`. No contract
  covers the outside of a seam.
- **The round-trip failure at df = 1E7 near q = −8 is a contract defect.** At
  `q = −8` and `−8 + 1 ulp` the true quantile is inside the inverse contract's
  domain (`|q| ≤ 8`, 5E-15), yet the result is off by 1.29E-12.
  - The continued fraction overestimates the tail just above 8.
  - That pulls the inverse's loop above 8, where the polish does not run.
  - #213 tracks it as an open silent-wrong release blocker.

  No #34 wave row had a quantile this close to 8.
