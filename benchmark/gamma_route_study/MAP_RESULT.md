# #307 binary64 route map: result (2026-11-10)

Excel VBA Regression run 42783017314, candidate `f3ff877`, Excel 16.0 build
20430, 64-bit. The same job passed the regression (5,987/5,987) and the study
batteries (8,901/8,901), then exported the main grid, the holdout (both
unchanged) and the map. The evidence commit is `f0dde69`.
`analyze_binary64_map.py --write` produced `map_results.json`.

## The pre-#242 binary64 route

203 of 282 points meet their governing contract. By shape (worst errors over
the shape's points):

```text
  A=1e-06     6/11 meet   worst P  2.19e-15  worst Q  7.45e-09   fail at X=1e-06, 0.001, 0.05, 0.5, 1
  A=1e-05     6/11 meet   worst P  1.73e-15  worst Q  1.20e-10   fail at X=1e-06, 0.001, 0.05, 0.5, 1.00001
  A=0.0001    6/11 meet   worst P  1.67e-15  worst Q  6.10e-11   fail at X=1e-06, 0.001, 0.05, 0.5, 1.0001
  A=0.001     6/11 meet   worst P  6.08e-16  worst Q  2.08e-12   fail at X=1e-06, 0.001, 0.05, 0.5, 1.001
  A=0.01      7/11 meet   worst P  7.62e-16  worst Q  2.66e-13   fail at X=0.001, 0.05, 0.5, 1.01
  A=0.05      9/11 meet   worst P  7.14e-16  worst Q  3.57e-14   fail at X=0.5, 1.05
  A=0.1      10/11 meet   worst P  4.46e-16  worst Q  1.84e-14   fail at X=1.1
  A=0.25     11/11 meet   worst P  4.50e-16  worst Q  2.68e-15
  A=0.5      11/11 meet   worst P  8.82e-16  worst Q  4.39e-15
  A=0.75     11/11 meet   worst P  7.00e-16  worst Q  2.00e-15
  A=1         8/ 8 meet   worst P  3.32e-16  worst Q  1.50e-15
  A=1.5       6/10 meet   worst P  4.45e-15  worst Q  2.15e-14   fail at X=2.5, 2.5, 2.72474, 5.17423
  A=2        10/10 meet   worst P  2.40e-16  worst Q  1.17e-15
  A=3        10/10 meet   worst P  7.14e-16  worst Q  2.29e-15
  A=5        10/10 meet   worst P  1.31e-15  worst Q  4.53e-15
  A=10       11/11 meet   worst P  1.34e-15  worst Q  3.03e-15
  A=20        8/11 meet   worst P  5.23e-15  worst Q  6.97e-15   fail at X=15.5279, 46.8328, 109.443
  A=50       10/12 meet   worst P  9.14e-15  worst Q  1.67e-14   fail at X=28.7868, 120.711
  A=100       6/12 meet   worst P  4.82e-14  worst Q  7.92e-14   fail at X=40, 70, 130, 160, 200, 300
  A=300       7/13 meet   worst P  8.53e-14  worst Q  1.72e-13   fail at X=126.795, 196.077, 317.321, 403.923, 473.205, 646.41
  A=1000      6/13 meet   worst P  2.13e-13  worst Q  5.81e-13   fail at X=683.772, 810.263, 1031.62, 1094.87, 1189.74, 1316.23, 1632.46
  A=10000     8/13 meet   worst P  1.29e-14  worst Q  1.96e-14   fail at X=9000, 9400, 10300, 10600, 11000
  A=100000    9/13 meet   worst P  1.46e-14  worst Q  1.68e-14   fail at X=98102.6, 99051.3, 100949, 106325
  A=1e+06     7/13 meet   worst P  3.69e-14  worst Q  3.70e-14   fail at X=990000, 994000, 997000, 1.003e+06, 1.01e+06, 1.02e+06
  A=1e+07     4/13 meet   worst P  9.00e-14  worst Q  8.84e-14   fail at X=9.96838e+06, 9.98103e+06, 9.99051e+06, 9.99684e+06, 1.00032e+07, 1.00095e+07, 1.0019e+07, 1.00316e+07, 1.00632e+07
```

Three regions:

- **Shapes up to 0.1, below `A + 1`.**
  - Errors reach 7.5E-9.
  - The raw series P is accurate (worst P error 2.2E-15); Q fails through the complement `Q = 1 - P`, the #24 defect.
  - The seam point `X = A + 1` fails at the same shapes.
- **Shapes 0.25 to 10.** Every point meets the contract except four at `A = 1.5`, but some margins are thin: Q at `A = 0.5` is 4.4E-15 against 5E-15.
- **Shapes from 20.** Errors grow with the shape:
  - 1.7E-14 at 50;
  - 7.9E-14 at 100;
  - 5.8E-13 at 1,000.

  They fail the general contract away from the centre, and at 1,000 also the `large_shape` contract.

## Where the error comes from

`analyze_exact_deviance.py` replays the map from shape 0.25 in Python
binary64, with the recurrences, stopping rules and small-shape backward
re-evaluation of `PROB_TryGammaSeriesP` and `PROB_TryGammaContinuedFractionQ`.
The prefactor is computed as
`x^A e^-x / Gamma(A) = A exp(-stirlerr(A) - D) / sqrt(2 pi A)`, where Loader's
deviance is `D = A log(A/X) + X - A`.

| Replay | Points meeting their contract | Worst error |
| --- | ---: | ---: |
| `D` exact, production stopping rules | 179 of 205 | 8.9E-14 |
| `D` direct in binary64, production stopping rules | 134 of 205 | 1.4E-9 |
| `D` exact, continued fraction stopped at machine epsilon | 190 of 205 | 8.9E-14 |

The binary64 route has two separate sources of error:

- **The prefactor's deviance.** This is the dominant source. An absolute error of D becomes a relative error of the result, and D reaches 20 to 50 in the tails.
- **The continued fraction's stop at `PROB_NUM_EPS` (3E-14) above shape 1.** It costs up to 2.3E-14 in Q, for example at `A = 1.5`, `X = 2.5`. At shape 1 and below, production already stops at machine epsilon and re-evaluates backwards (#26).

With both corrected, the remaining 15 failures are:

- far points of shapes from 1,000, under the general `3E-15`, where the lower series sums many terms (1.0E-14 at shape 1E5, 8.9E-14 at 1E7);
- one point at `A = 100`, `X = 300`, where Q is 5.08E-15 against 5E-15.

An earlier version of this replay stopped the continued fraction at machine
epsilon everywhere, which mixed the two sources. The review of #311 found
this. The table separates them.

## What this means for #307

Routing to the *old* binary64 route would recover speed only for shapes 0.25
to 10, with thin margins. A binary64 route with an accurate deviance could
serve the practical domain, at about the cost of binary64. This is the shape
range 0.25 to about 1,000, plus the centre of larger shapes. Double-double
would stay where it is needed:

- the region below `A + 1` at small shapes;
- the far tails of large shapes.

The next step is a deviance accurate to about double-double precision, built
from a few error-free binary64 transformations rather than the double-double
library, together with a machine-epsilon stop for the continued fraction above
shape 1. Its accuracy has to be measured first, then the routing rule
preregistered with the frozen holdout of `HOLDOUT_DESIGN.md` (branch
`study/307-routing-holdout`). No contract, threshold or production source
changed here.
