# #14: the Gamma inverse at the final scale — second preregistration

Registered 2026-11-14 (Europe/Rome), **before the second holdout is
evaluated**. `gamma_inverse_scale_grid2.csv` is committed with its design
and references only; every observed field is blank. No source under `src/`
changes in the commit that carries this file: the branch still holds the
first candidate, which the commit after this one replaces.

This file is an addition. `PREREGISTRATION.md`, its Addendum 1 and
`RESULT_1.md` are unchanged, and the first rule stays rejected.

## Why a second candidate can be registered

The first holdout was rejected on G7 alone (`RESULT_1.md`): at five shapes
the seam probes did not reach the crossover in Excel, so the crossover was
not tested there. G1 to G6 held on every row. Two things were learned, and
both come from one cause, the accuracy of `PROB_TryLogGamma1p` for shapes
between 0.1 and 2:

- its value differs from the correctly rounded one by enough to move the
  crossover past probes thirty-two units in the last place away;
- it limits the log branch to about 2.3E-14, where the pair arithmetic
  reaches 2E-16.

The second candidate removes the cause instead of widening a tolerance.

## The rule

As in `PREREGISTRATION.md` and Addendum 1, with one change: **the source of
`LogGamma(1 + A)`**, in the classifying value `L0` and in the log branch
alike.

| Shape | `LogGamma(1 + A)` |
| --- | --- |
| `A <= 1` | `PROB_GammaSmallLogGamma1p(A)`: the degree-32 polynomial for `LogGamma(1 + A) / (A (A - 1))` of #26, with exact zeros at 0 and 1 |
| `1 < A <= 2` | `PROB_GammaSmallLogGamma1p(A - 1) + Log(A)`, from `Gamma(1 + A) = A Gamma(A)`; `A - 1` is exact |
| `A > 2` | `PROB_TryLogGamma1p(A)`, as in the first candidate |

The crossover `L0 <= -40`, the pair arithmetic, the exact rescaling by
`2^128` outside the pair logarithm's range, the unit-scale solver above the
crossover and the representability policy at the final scale are unchanged.

`gamma_route_study/route_model.py` reproduces the polynomial bit for bit
(#307), so the model's `L0` is production's for shapes up to 1, and differs
from it only by the rounding of one `Log` between 1 and 2.

## Development evidence, on revealed data only

`inv_scale_model.py` (`log_gamma1p_2`, `candidate_2`) and `test_holdout.py`.

| Sample | Worst relative error |
| --- | ---: |
| The polynomial source against `LogGamma(1 + A)`, 20,000 shapes from 1E-8 to 2 | 4.0E-16 |
| The first holdout's 350 rows the second model answers in log space | 2.5E-16 |

The first holdout is revealed, so it is development data here. Its Excel
observations are those of the first candidate and are not reused as evidence
for the second.

## The holdout

`holdout_build.py --holdout 2` writes 525 rows from `random.Random(2014)`.
No row has the probability, shape and scale of a row of the first holdout.

| Kind | Rows | What it holds |
| --- | ---: | --- |
| `underflowed` | 120 | As in the first holdout, from the new seed |
| `log_region` | 120 | As in the first holdout, from the new seed |
| `seam` | 170 | Ten shapes, none of them a seam shape of the first holdout: 3E-4, 0.005, 0.03, 0.2, 0.7, 0.95, 1.5, 2, 4 and 12. At each, the last probability the second model answers in log space and its neighbours at 1, 4, 16, 64, 256, 1,024, 4,096 and 16,384 units in the last place on either side |
| `chisquare` | 60 | As in the first holdout, from the new seed |
| `limits` | 25 | As in the first holdout, from the new seed, with other shape-1 rows |
| `shape_boundary` | 30 | The shapes where the source of `LogGamma(1 + A)` changes: the predecessor of 1, 1, its successor, the predecessor of 2, 2 and its successor, five points each in the log region |

The wider seam is the repair of the first holdout's probes. At shape 12 the
outermost probes are 1.9E-13 from the crossover in `L0`, where the first
holdout's reached less than 7.1E-15; at smaller shapes they reach further.

The seeded points are evaluated neither by the model nor by Excel before
the verdict. References are computed as for the first holdout.

`M_STATS_PROBDIST_GIS2.bas` records the same four columns. `L0` is read from
production through `PROB_TryGammaInvLogUnit`, the function the kernel
itself classifies with, so the exporter cannot disagree with the kernel
about the branch. The exporter is added with the candidate, since it calls
that function.

## Acceptance rules

G1 to G7 of `PREREGISTRATION.md`, unchanged, on the 525 rows:
`holdout_analyze.py --holdout 2`. The bound stays 1E-12 relative, the
round-trip bound `5E-13 + A * 1E-12`. The `shape_boundary` rows are scored
by G1, G5 and G6 like any other.

## Contracts

The four contracts of the regime `scaled_quantile_from_underflowed_unit` and
their 13 main-grid rows stay as registered, with their target of 1E-14. The
rows are exported again with the second candidate.

## Predictions, fixed before evaluation

1. G1 to G7 hold on all 525 rows.
2. The worst relative error of the rows Excel answers in log space is below
   5E-15, and below 1E-15 for shapes up to 2. Reported, not gating.
3. Against the main grid as exported from the first candidate, only the 13
   rows of the new regime may change; the four contracts meet their target
   of 1E-14. No other observation of the main grid or of the consolidated
   holdout changes.
4. The 1,933 assertions of the Excel regression pass unchanged.
5. `ChiSquare_InverseCumulative` equals `Gamma_InverseCumulative` at scale 2
   bit for bit on all 60 rows, as in the first holdout.

## Rejection policy

Any failure of G1 to G7 rejects this rule. Nothing is retuned against the
revealed holdout, and the unit-scale solver stays in production alone. A
later candidate needs its own registration and its own untouched evidence.

## Order of work

1. This file, `RESULT_1.md`, the second model, the second holdout with blank
   observations and the tests are committed and pushed.
2. The kernel takes the new source of `LogGamma(1 + A)` and exposes its
   classifying value; the second exporter is added.
3. One Excel wave exports the main grid, the consolidated holdout and the
   second holdout. `holdout_analyze.py --holdout 2` gives the verdict.
