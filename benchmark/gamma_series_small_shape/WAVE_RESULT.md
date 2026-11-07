# #23 Excel wave result (2026-11-07)

Excel VBA Regression run 42365758589, candidate `d22486c`, Excel 16.0 build
20430, 64-bit: the regression passed 955/955, and the same job exported the
main grid, the holdout, the #23 study grid (`gss_small_shape`) and the fitting
ladder (`prs_gamma_q_complement`). The evidence commit is `e278318`. The main
grid and the holdout did not change. `analyze_gamma_series_small_shape.py`
scores the wave; `wave_results.json` holds every number below.

## Part A: the frozen holdout, scope probe and counterexample

All 49 rows echo exactly the binary64 inputs of their references, return a
number, and meet `3E-15` relative.

| Regime | Rows | Worst relative error | At |
| --- | ---: | ---: | --- |
| `holdout` | 44 | 2.32E-15 | `ChiSquare_Cumulative(1.5, 0.00014)` |
| `scope_probe` | 3 | 5.74E-16 | `Gamma_Cumulative(1.0, 5.3, 1)` |
| `counterexample` | 2 | 1.23E-15 | `ChiSquare_Cumulative(1, 0.0002)` |

The Gamma and Chi-square routes return bit-identical results at all 23 paired
points. The design predicted the scope probe at about `9E-15` to `4E-14` before
the fix; after it, the worst is 5.7E-16.

## Part B: the fitting ladder against its pre-fix export (`5c884bc`)

| Rule of `MOVEMENT_MANIFEST.md` | Result |
| --- | --- |
| B1 `x05_s0`–`x05_s10`: series P and CDF move and meet `3E-15` | holds: all 11 moved; worst CDF error 2.04E-15, was up to 1.28E-14 |
| B2 `x05_s11` (Shape = 0.5) stays on its path | holds: unchanged, 9.6E-17 |
| B3 `xsub`: the CDF error envelope does not regress | holds: worst 6.36E-14 before, 5.39E-14 after |
| B4 `x2_s0`–`x2_s11`: public CDF and survival do not move | **falsified**: the survival moved at 11 of 12 rows |

The `x05` survival, which #24 owns, moved from up to 1.4E-8 to up to 1.6E-9.
It still breaches its contract until #24 removes `Q = 1 - P`, as the manifest
predicted.

## The falsified prediction

The manifest froze B4 because `X = 2` selects the continued-fraction path. The
prediction was wrong about that path's inputs:
`PROB_TryGammaContinuedFractionQ` forms its result with the same
`PROB_TryGammaPrefactor` as the series, so it also consumes the corrected
off-grid `PROB_StirlingError`. That is the mechanism of the manifest's
post-selection explanation, applied to a consumer it did not name.

The manifest requires such movement to be explained before #23 closes. The
explanation has four testable consequences, and the analyzer checks each one
(B4X):

- the public CDF does not move at any `X = 2` row. It holds: the survival is
  below `P`'s rounding there;
- the half-integer control `x2_s11` (Shape = 0.5, table path) does not move.
  It holds;
- only off-grid shapes move. It holds: all 11 moved rows have off-grid shapes;
- no moved survival error grows. It holds: worst 5.58E-14 before, 4.53E-14
  after.

No input, threshold or rule was changed after the export. B4 stays recorded as
falsified.

## What this settles, and what it does not

**Settled for #23:**

- the minimal Gamma and Chi-square calls meet their contracts;
- the preregistered scope probes meet `3E-15`;
- the fitting ladder and the untouched holdout pass;
- no series, main-grid or holdout result regressed;
- the continued-fraction survival improved.

**Not settled:**

- **#26.** The `X = 2` continued-fraction survival still breaches
  `Gamma_Survival.all.output` (`5E-15`) at every shape, worst 4.53E-14.
- **Promotion.** Representative #23 rows are not yet promoted into the main
  grid, so the gate does not yet watch this region.
- **#13.** The frozen `positive_ratio_subnormal` arms have not been re-run on
  this source, so the cutoffs are not yet revalidated.
