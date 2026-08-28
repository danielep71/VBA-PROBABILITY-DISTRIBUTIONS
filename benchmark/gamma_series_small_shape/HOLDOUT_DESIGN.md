# Preregistered holdout and scope probe for #23

This design is frozen before production source changes. The points below were
chosen from source structure and from the already-measured `X = 0.5` ladder;
no post-fix output has been observed.

## Independent small-shape holdout

The holdout uses public inputs disjoint from the fitting ladder.

- kernel shapes:
  `2E-6, 7E-6, 2E-5, 7E-5, 2E-4, 7E-4, 2E-3, 7E-3, 2E-2, 7E-2, 2E-1`;
- standardized arguments: `0.125` and `0.75`;
- Gamma route: `X = standardized argument`, `ScaleParam = 1`;
- Chi-square route: public `X = 2 * standardized argument`,
  `DegreesFreedom = 2 * kernel shape`.

Both transformations are exact at these binary64 inputs, so disagreement
between the two public routes cannot be attributed to #13's positive-ratio
loss. References must be evaluated at the echoed binary64 public inputs.

The existing frozen thresholds apply: `3E-15` relative for each cumulative
surface. The holdout does not choose a new threshold or a new dispatch point.

## Candidate-boundary seam

If the production fix uses a shape boundary, test at each standardized argument
`0.125, 0.5, 0.75, 1.0` and at:

- the largest binary64 value below the boundary;
- the boundary itself;
- the smallest binary64 value above the boundary.

Neighbours must be constructed arithmetically from a power-of-two ULP, never
from long decimal literals. No seam jump may exceed the governing CDF contract.

## Real-Excel scope probe

These Gamma CDF points test whether #23 is truly confined below shape `0.5` or
is the visible edge of an off-grid prefactor defect:

| Public X | Shape | Scale | Reference P |
| ---: | ---: | ---: | ---: |
| 0.5 | 0.6 | 1 | 0.6189010171407936557816432 |
| 1.0 | 5.3 | 1 | 0.002157194707472261441497618 |
| 10.8 | 10.3 | 1 | 0.6014210971218611799875902 |

The binary64 source mirror predicts relative errors of approximately
`9E-15`, `2.5E-14`, and `4E-14`, respectively. Those figures are hypotheses.
Only a real-VBA export may confirm them and determine whether #23 must expand.

## Independence rule

The listed holdout inputs, thresholds, and acceptance rules must not change
after post-fix output is observed. A rejected holdout rejects the candidate;
it does not become a second fitting set.
