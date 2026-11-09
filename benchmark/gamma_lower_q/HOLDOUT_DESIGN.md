# Preregistered #24 holdout: lower-region Gamma Q (2026-11-09)

This design is frozen and committed before any output of the candidate is
computed at its points, in Python or in Excel. The inputs are in `design.py`;
nothing below may change after the holdout is evaluated. A rejected holdout
rejects the candidate; it does not become a fitting set.

## What it tests

#24: below the branch seam `X = A + 1`, survival was formed as `Q = 1 - P`, so
the absolute error of P became the absolute error of Q, and the relative error
grew by `P / Q`. That ratio is about 1.8E6 at shape 1E-6 and `X = 0.5`. The
candidate is the current `release/1.0.0` source (`279f15c`): #242 computes
lower-region Q from log P before rounding. The holdout targets the region where
the amplification was largest, and the seam.

## Inputs

- **Kernel shapes A:** `1.7E-7, 4.5E-6, 6.5E-5, 8.5E-4, 3.7E-3, 4.3E-2, 0.465`.
- **Standardized arguments X per shape:**
  - interior: `0.05`, `0.25`, `0.6`, all below `A + 1`. At `A = 1.7E-7`,
    `X = 0.6` the old amplification `P / Q` is above 1E6;
  - seam: the binary64 predecessor of `A + 1`, `A + 1` itself and its successor,
    built with `math.nextafter`, never from decimal literals.
- **Routes, for every (A, X):**
  - `Gamma_Cumulative(X, A, 1)` and `Gamma_Survival(X, A, 1)`;
  - `ChiSquare_Cumulative(2X, 2A)` and `ChiSquare_Survival(2X, 2A)`.
  
  The doubling is exact in binary64, so both routes reach the same kernel inputs.

That is 7 shapes, 6 arguments and 4 routes: 168 rows.

### Disjointness

Shapes were chosen before any evaluation and kept only if `check_disjoint.py`
finds neither `A` nor `2A` in any tracked CSV, JSON, Markdown, VBA or Python
file. The scan rejected these first choices: `1.5E-7`, `2.5E-3`, `4.5E-2`,
`0.35`, `0.37`, `0.43`, `0.41`, `0.47`, `0.53`, `0.29`, `0.31`, `0.385`,
`0.415`, `0.445`.

## References

P and Q are computed separately with mpmath (`gammainc`, lower and upper,
regularized) at the binary64 inputs, at 80 and 120 digits. Both readings must
agree to 1E-60 relative, or the row is not evaluated. Q is never formed as
`1 - P` in the oracle.

## Acceptance rules

Thresholds are the frozen contracts. No new threshold is chosen.

| Id | Rule |
| --- | --- |
| H1 | Every Gamma and Chi-square survival row is within `5E-15` relative (`Gamma_Survival.all.output`). Every cumulative row is within `3E-15` relative (`Gamma_Cumulative.all.output`). The Chi-square rows are held to the same bounds, although their registry contracts are broader. Every row returns a number, and the echoed inputs equal the design inputs. |
| H2 | At every (A, X), the Gamma and Chi-square routes return the same Double, for P and for Q. |
| H3 | For each shape, over the six arguments in increasing order, the CDF never decreases and the survival never increases. No allowance. |
| H4 | At every (A, X), `abs(P + Q - 1) <= 3E-15 * P + 5E-15 * Q`, evaluated exactly from the two returned Doubles. |

Any failure rejects the candidate for #24.

## Not in scope

The binary64 fallback at the double-double pair's iteration cap, the only place
where the lower branch still forms `1 - P`, occurs only for `A` above about
1.07E8. There `P <= P(A, A + 1)`, which is about 0.5, so the complement does
not amplify P's relative error. #24 treats it as a documented exception, with
its own Excel assertion; it is not part of this holdout.
