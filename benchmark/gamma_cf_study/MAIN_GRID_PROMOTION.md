# #26 main-grid promotion: predictions before the Excel wave (2026-11-09)

Written and committed before the Excel wave that observes these rows. Nothing
below may be revised after that wave's output is observed.

## Why

#26's acceptance requires main-grid regressions that cover the resolved path.
Before this change the main grid held one `Gamma_Survival` row at `X = 2`
(shape 2) and no survival row with a shape below 1 under a contract tighter
than `2.6E-10`. The 41 VBA assertions of #242 cover the path; the strict
Accuracy Gate did not.

## The rows

`promote_x2_ladder.py` promotes the twelve `x2_*` shapes of the frozen ladder
in `benchmark/positive_ratio_subnormal/gamma_q_complement.csv`, selected by
their role (`X = 2`), not by a measured error. Each shape `A` gives two rows
with one reference, `Q(A, 2)`:

- `Gamma_Survival(2, A, 1)`;
- `ChiSquare_Survival(4, 2A)`, the exact doubling.

The 24 rows enter the main grid through `benchmark/promote_grid_rows.py` with
blank observations, regime `all` and claim `rel<=5E-15`. References are computed
at the binary64 shapes, stable between 80 and 120 digits, and reproduce the
pre-fix errors recorded by the #235 study (`isolation.json`) to 1E-28.

`Gamma_Survival.all.output` is `5E-15`, so the Gamma rows are gated directly.
`ChiSquare_Survival.all.output` is `2.6E-10`; `test_promote_x2_ladder.py` holds
the Chi-square rows to `5E-15` independently, as #23 does for its promoted
Chi-square row. No contract or threshold changes.

## Model values

`wide_gamma.public`, the model that reproduces the production kernel bit for
bit in `test_public_fresh.py`, gives these survival values:

| Shape | Q (binary64) | Relative error |
| ---: | --- | ---: |
| 1E-6 | `0x1.a40d873a1dc08p-25` | 8.4E-18 |
| 3E-6 | `0x1.3b0a65ec3a85ap-23` | 2.2E-17 |
| 1E-5 | `0x1.06896666d8be2p-21` | -4.3E-17 |
| 3E-5 | `0x1.89d13fe4524d5p-20` | 3.7E-17 |
| 1E-4 | `0x1.48378fa514299p-18` | 6.8E-17 |
| 3E-4 | `0x1.ec7ab7136affcp-17` | 5.3E-17 |
| 1E-3 | `0x1.9ad92152eae8ap-15` | -3.3E-17 |
| 3E-3 | `0x1.35192c80daf91p-13` | -3.2E-17 |
| 0.01 | `0x1.04652f8e34bfap-11` | -3.5E-17 |
| 0.03 | `0x1.92bca6af2675bp-10` | 3.5E-17 |
| 0.1 | `0x1.73d6f86fc1cfbp-8` | -7.2E-17 |
| 0.5 | `0x1.74bcf82c9d860p-5` | 6.1E-17 |

The pre-fix observations in the ladder breached `5E-15` at every shape, from
3.5E-14 to 4.5E-14.

## Predictions

| Id | Prediction |
| --- | --- |
| M1 | All 24 promoted rows return a number. Each Gamma row and its Chi-square doubling return the same Double, and that Double is the model value above. Every row is within `5E-15`; the worst is 7.2E-17. |
| M2 | No observation of a row that was already in the main grid changes, and no holdout observation changes. No VBA source changes. |
| M3 | The regression suite passes all 14,886 assertions. |
| M4 | Every active contract passes: 166/166 main, 83/83 holdout. The 27-row coverage debt is unchanged. |

Any failure of M1 to M4 stops the closure of #26, and the cause is explained
before anything else changes.
