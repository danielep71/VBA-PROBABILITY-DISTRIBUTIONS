# #14: the Gamma inverse at the final scale — preregistration

Registered 2026-11-14 (Europe/Rome), **before the holdout is evaluated**.
`gamma_inverse_scale_grid.csv` is committed with its design and references
only; every observed field is blank. No production source changes in the
commit that carries this file.

## The defect

`K_STATS_Gamma_InverseCumulative` solves on the unit scale and multiplies by
`ScaleParam` afterwards. A unit-scale quantile below the Double range is
lost before the multiplication, even where the final quantile is an ordinary
number. At probability 1E-4, shape 0.01 and scale 1E100 the unit quantile is
about 5.7E-401 and the final quantile about 5.7E-301.
`K_STATS_ChiSquare_InverseCumulative` has the same structure with a final
factor of 2.

The cost benchmark of #314 shows a second effect in the same region: at
shape 0.001 an inverse takes more than a second, and is refused where the
unit quantile underflows.

## The rule

For a small unit quantile `x` the lower tail is its leading term:

    P(A, x) = x^A / Gamma(A + 1) * (1 - A x / (A + 1) + ...)
    ln x    = (ln p + ln Gamma(1 + A)) / A + x / (A + 1) + O(x^2)

The call is classified by one plain Double,

    L0 = (Log(Probability) + PROB_TryLogGamma1p(Shape)) / Shape

and, with `F` the final factor (`ScaleParam`, or 2 for Chi-square):

| Condition | Answer |
| --- | --- |
| `L0 <= -40` | `exp((ln p + ln Gamma(1 + A)) / A + ln F)`, formed once |
| otherwise | the unit-scale solver, then one multiplication by `F`, as before #14 |

At the crossover the unit quantile is `exp(-40)`, about 4.2E-18, and the
neglected term is below 4.2E-18 relative, well under one unit in the last
place.

In the log branch the two logarithms, the quotient and the sum are carried
as pairs of Doubles, with the `PROB_B64_Pair` routines the binary64 Gamma
route already has, and the result is `Exp(High) * (1 + Low)`. A plain
evaluation would inherit the size of the exponent: with `|L0|` up to 1,415
its error reaches 3.7E-13 on the development sample below. A subnormal
probability or factor takes the plain logarithm with a zero low part,
because the pair logarithm needs a normal Double.

Representability is decided at the final scale only:

- the exponential underflows: the result is zero;
- the exponential overflows: `#NUM!`;
- neither: the Double it returns, subnormal ones included.

The probability is used as given. The complement is not used in the log
branch.

### Deviation from the issue's sketch

The issue describes a kernel that returns the logarithm of the unit quantile
in every regime, with one exponentiation in the wrapper. That is narrowed:
the kernel returns the quantile at the final scale, and takes the logarithm
only in the log branch. Above the crossover a logarithm followed by an
exponential would add `|ln x|` units in the last place to every existing
result and move every existing row; the multiplication it replaces is exact
to half a unit.

## Development evidence, on revealed data only

`inv_scale_model.py` models the log branch with the pair routines of
`gamma_route_study/route_model.py`, which matched Excel bit for bit in #307.
`PROB_TryLogGamma1p` is modelled as the correctly rounded value.

| Sample | Points | Worst relative error |
| --- | ---: | ---: |
| Pair arithmetic, unit quantile from `exp(-1450)` to `exp(-40)`, final quantile normal, shapes 1E-8 to 16 | 30,000 | 3.5E-16 |
| The same formula in plain Doubles, unit quantile below the normal range | 20,000 | 3.7E-13 |

The two samples were drawn in throwaway scripts from `random.Random(14)`,
the holdout's seed, but over other ranges and in another order of draws, so
their points are not the holdout's. `test_holdout.py` uses seed 1400 and
other seam shapes.

No inverse row of the main grid or of the consolidated holdout has
`L0 <= -40`: none of the 36 Gamma and Chi-square inverse rows of the main
grid is at or below the crossover.

## The holdout

`holdout_build.py` writes 455 rows from `random.Random(14)`:

| Kind | Rows | What it holds |
| --- | ---: | --- |
| `underflowed` | 120 | Unit quantile below the smallest normal Double, down to `exp(-1415)`; the scale brings the final quantile into the normal range. The first row is the counterexample above |
| `log_region` | 120 | Unit quantile between the smallest normal Double and `exp(-40.5)`; half at scale 1 |
| `seam` | 130 | Ten shapes from 1E-4 to 15; at each, the last probability the model answers in log space and its neighbours at 1, 2, 4, 8, 16 and 32 units in the last place on either side |
| `chisquare` | 60 | Scale exactly 2: unit quantiles in `[MIN/2, MIN)`, in the log region, and just above the crossover |
| `limits` | 25 | Final quantile below every Double (10), subnormal (10), and at shape 1 one just inside the range and four above it |

The seeded points are evaluated neither by the model nor by Excel before
the verdict. The seam is located from the model's predicate, by bisection on
the bit pattern of the probability.

At every row `M_STATS_PROBDIST_GIS.bas` records, as bit patterns:
`Gamma_InverseCumulative(p, A, S)`; `ChiSquare_InverseCumulative(p, 2A)`
where `S = 2`; `Gamma_Cumulative` at the returned quantile; and `L0` as
Excel computes it. The inputs travel as bit patterns too.

References are the exact quantile at the final scale to 40 digits: the
fixed point of the lower series with four terms at 60 digits, or
`-ln(1 - p) * S` at shape 1.

## Acceptance rules

`holdout_analyze.py` implements them. The bound is **1E-12 relative**, the
inverse guarantee of #327.

| Id | Rule |
| --- | --- |
| G1 | Where the reference is a normal Double, the Gamma inverse returns a number within 1E-12 relative of it |
| G2 | A reference below every Double returns the exact zero; one above every Double returns `#NUM!`; a subnormal one is within one unit of the subnormal spacing plus 1E-12 relative |
| G3 | Within a seam cluster, over increasing probability, the quantile does not fall by more than 1E-12 times the sum of the two references |
| G4 | At scale 2 the Chi-square inverse returns the same Double, or the same error, as the Gamma inverse |
| G5 | Where the quantile is a normal Double, `Gamma_Cumulative` at it is within `5E-13 + A * 1E-12` relative of the probability: the loosest forward contract of the region (#13) plus the inverse bound carried through the tail |
| G6 | No cell holds a runtime fault or a failed kernel |
| G7 | Every seam cluster holds both branches, by the `L0` Excel exported |

G7 is a validity rule for the probes: the seam is built from a model whose
`LogGamma1p` may differ from production's by a few units in the last place.
If a cluster falls on one side only, the crossover was not probed there and
the holdout cannot be accepted.

## Contracts, fixed before the export

Regime `scaled_quantile_from_underflowed_unit`, on the main grid:

| Contract | Measure | Target |
| --- | --- | ---: |
| `Gamma_InverseCumulative` | quantile relative error | 1E-14 |
| `Gamma_InverseCumulative` | tail-probability residual, relative | 1E-14 |
| `ChiSquare_InverseCumulative` | quantile relative error | 1E-14 |
| `ChiSquare_InverseCumulative` | tail-probability residual, relative | 1E-14 |

The guarantee of each is that of its tier under #327. Their rows are added
to the main grid with blank observations in the candidate commit.

## Predictions, fixed before evaluation

1. G1 to G7 hold on all 455 rows.
2. The worst relative error of the rows Excel answers in log space is below
   1E-14. Reported, not gating.
3. No existing observation of the main grid or of the consolidated holdout
   changes. The assertions of the Excel regression that exist today pass
   unchanged.
4. A call answered in log space takes microseconds, where the solver took
   more than a second at shape 0.001.

## Rejection policy

Any failure of G1 to G7 rejects this rule. The crossover, the bound and the
probes are not retuned against the revealed holdout, and the unit-scale
solver stays in production alone. A later candidate needs its own
registration and its own untouched evidence.

## Order of work

1. This file, the model, the holdout with blank observations, the exporter,
   the analyzer and the tests are committed and pushed.
2. The kernel and the two wrappers change; the contracts, their rows and the
   Excel regressions are added.
3. One Excel wave exports the main grid, the consolidated holdout and this
   holdout. `holdout_analyze.py` gives the verdict.
