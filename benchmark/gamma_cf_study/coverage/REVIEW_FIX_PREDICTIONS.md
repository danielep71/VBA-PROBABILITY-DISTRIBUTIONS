# #242 review fixes: predictions before the Excel wave (2026-11-09)

Written and committed before the Excel wave that tests these fixes. Nothing
below may be revised after that wave's output is observed.

## The changes

All three are made in `public_integration.py`, which generates the public
adapters and extracts the frozen port's arithmetic. The frozen arithmetic is
unchanged.

1. **Reach.** `PROB_TryGammaRegularizedP` and `Q` fall back to the binary64
   series (`X < A + 1`) or continued fraction (otherwise) when the scaled pair
   reports its 100,000-iteration cap. The fallback is the pre-#26 route.
2. **Errors.** The handlers of `PROB_DD_WG_Pair`, `PROB_DD_EG1_Pair` and
   `PROB_DD_EG_Pair` re-raise every runtime error except 5, the kernel's own
   domain and cap signal.
3. **Dead code.** The extraction follows names in code only, not in comments.
   The Precision v2 entry points that a comment had pulled in
   (`PV2_Pair`, `PV2_LogGamma`, `PV2_Log`, `PV2_Log1p`, `PV2_Expm1`,
   `PV2_Exp`, `PV2_Ln2`, `PV2_Value`) are no longer generated. The extracted
   set falls from 54 procedures to 46.

## Predictions

| Id | Prediction |
| --- | --- |
| R1 | The regression suite passes all **14,886** assertions: the 14,878 of run 42529642623 plus the 8 of `Test_PublicGamma_Reach`. |
| R2 | `Test_PublicGamma_Reach` passes at `X = A = 1.2E8` and at `A = 5E11`, `X = A + 2`. Both P and Q return success with no diagnostic, P within 2E-12 and Q within 5E-13 relative of the quadrature references (the `extreme_shape` contracts). Before the fix, `wide_gamma.public` raises its cap at both points. |
| R3 | A full main export changes **no** observation, and a holdout export changes **no** observation, against `e91c535`'s exports. No committed grid point reaches the cap, so the fallback never runs there; change 2 acts only on an error path; change 3 removes unreachable code. |
| R4 | Every active contract still passes: 166/166 main, 83/83 holdout. |

Any failure of R1 to R4 stops the merge, and the cause is explained before
anything else changes.

## References

`X = A = 1.2E8`: P = 0.50001213942700714067, Q = 0.49998786057299285933.
`A = 5E11`, `X = A + 2`: P = 0.50000131644236160748, Q = 0.49999868355763839252.

Both are computed by quadrature of the normalized density over +-60 standard
deviations, split at each standard deviation, at 40 and 60 digits. The two
precisions agree to 7E-29, and P + Q - 1 is below 4E-49. The same method
matches `wide_gamma.public` at `(1E8, 1E8)` to 2.9E-17.

## Addendum, before any wave output (2026-11-09)

The first wave dispatched on `fa891a2` was cancelled before it finished, and
none of its output was read. The modules then received comment-only amendments
that describe the integrated kernel: the module method notes, the
`PROB_TryGammaSeriesP` and `PROB_TryGammaContinuedFractionQ` headers, the
`PROB_CHI_MAX_DF` note and a `CONTINUOUS` note. They are applied and reversed
by `public_integration.py`. `test_public_integration.py` checks that each one
leaves the code part of every line unchanged.

Comments cannot change behaviour, so R1 to R4 stand unchanged for the wave on
the amended source.

## Result (2026-11-09)

Excel VBA Regression run 42670065251 executed candidate `c856ddd` on Excel 16.0
build 20430, 64-bit, then exported the main grid and the holdout. The evidence
commit is `e73fe7a`.

| Id | Result |
| --- | --- |
| R1 | **Holds.** 14,886 / 14,886 assertions pass, 0 fail. |
| R2 | **Holds.** `Test_PublicGamma_Reach` is part of the passing count: both reach points return success, within 2E-12 (P) and 5E-13 (Q). |
| R3 | **Holds.** The main export changes 0 of 1,737 in-scope observations, and the holdout export 0 of 659. |
| R4 | **Holds.** The main grid has no FAIL, KNOWN LIMITATION or PENDING contract, and the holdout passes 83 of 83. |

No prediction failed, and nothing was revised after the output was read.
