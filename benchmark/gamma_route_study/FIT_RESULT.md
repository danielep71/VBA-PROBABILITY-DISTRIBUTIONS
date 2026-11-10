# #307 repaired binary64 route: fit result (2026-11-10)

Excel VBA Regression run 42798193976, candidate `88ff6ba`, Excel 16.0 build
20430, 64-bit. The same job passed the regression (5,987/5,987) and the study
batteries (8,901/8,901), then exported the main grid, the holdout and the fit
set. The evidence commit is `52f8dd2`. `analyze_b64_fit.py --write` produced
`fit_results.json`.

## Fidelity

All 392 points agree with `route_model.py` bit for bit, in status and value,
after one correction to the model, made after this run.

The first analysis found 6 mismatches at 3 points: the seam point `X = fl(A + 1)` at `A = 1E-6`,
`1E-4` and `1E-3`.

- **What Excel did.** At those points the exact `A + 1` lies above `X`. Excel's comparison
  `X < A + 1#` nevertheless chose the continued fraction: it rounds `A + 1` to a Double
  before it compares.
- **What the model assumed.** A 64-bit intermediate, as for arithmetic that is stored.
- **The correction.** The model now compares in binary64. The series and continued-fraction
  values Excel returned at those points equal the model's.

The seam was already compared this way by the production adapter and the
pre-#242 route, so nothing else changes.

## Accuracy (exploratory)

Every tail the routing rule would return meets its governing contract at all
392 points:

- below `A + 1`, P from the series, and `Q = 1 - P` where `P <= 0.5`;
- from `A + 1`, Q from the continued fraction and `P = 1 - Q`.

The worst errors are 3.8E-15 at `A = 1E7`, inside one sd, where the bound is 2E-12 or 5E-13.
Elsewhere they are at most 3.3E-15. 88 lower-region points have `P > 0.5`; there Q stays with
the double-double kernel.

## Cost

| Routine | Fitted cost per call |
| --- | --- |
| Series | 1.5 us + 0.077 us per iteration |
| Continued fraction | 5.1 us + 0.091 us per iteration |

The worst call measured at each shape:

| A | Worst call (us) |
| ---: | ---: |
| up to 300 | 18 |
| 1,000 | 27 |
| 5,000 | 53 |
| 10,000 | 67 |
| 20,000 | 104 |
| 100,000 | 195 |
| 1,000,000 | 618 |
| 10,000,000 | 2,197 |

The double-double kernel takes 5 ms to 170 ms per call (#298).

## The shape ceiling

The frozen rule of `FIT_DESIGN.md` gives **A2 = 10,000**:

- every fit point up to 10,000 measured at most 67 us;
- the model predicts about 66 us at 10,000's worst arguments;
- at 20,000, the series at the predecessor of `A + 1` measured 104 us, over the 0.1 ms budget.

Above A2, the binary64 route is still 100 to 1,000 times faster than the
double-double kernel. The rule nevertheless leaves those shapes with the
kernel, because H6 of the frozen holdout requires every binary64-routed call
to cost under 0.2 ms. They remain slow after #307.

## Next

The routing rule is preregistered with A2 = 10,000 and its predictions for
the holdout, then implemented in the public P and Q, before the holdout is
evaluated.
