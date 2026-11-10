# #307 routing holdout: design (2026-11-10)

Frozen and pushed before any output of the binary64 route map (run 42783017314)
was read. The inputs are in `holdout_design.py`; nothing below may change
after they are evaluated. The routing rule is derived from the map alone. A
rejected holdout rejects the rule; it does not become a fitting set.

## What it tests

A routing rule for public Gamma P and Q:

- binary64 (the lower series below `X = A + 1`, the continued fraction from it) where the map shows that it meets the contracts;
- the double-double kernel elsewhere.

The holdout tests the rule's accuracy, its continuity and its cost on points
the rule never saw.

## Inputs

- **Kernel shapes A:** `3.7E-6, 7.7E-5, 2.9E-3, 0.0613, 0.383, 0.871, 1.33, 2.71, 6.17, 13.7, 37.9, 171.3, 743, 4111, 27191, 314159, 2718281`. `check_disjoint.py` finds neither A nor 2A in any tracked file.
- **Arguments X:**
  - for A below 1: `2.3E-5, 0.0371, 0.29, 1.13, 3.7, 17.3`;
  - for A from 1: `X = A + k sqrt(A)` with `k = -7.3, -2.2, -0.6, 0, 0.7, 2.3, 7.9, 13.1` (positive X only).
- **Total:** 117 points.
- **Routes, at every point:** `Gamma_Cumulative(X, A, 1)`, `Gamma_Survival(X, A, 1)`, `ChiSquare_Cumulative(2X, 2A)` and `ChiSquare_Survival(2X, 2A)`.
- **Boundary probes:** for every boundary of the frozen rule, the binary64 predecessor, the boundary and the successor, on each side the rule distinguishes, at the holdout shapes it crosses. They are constructed mechanically from the rule once it is frozen, before evaluation, and they are not chosen by looking at outputs.

## References

P and Q at the binary64 inputs, computed separately at high precision, never
one as `1 - other`, and stable between two precisions. This is the method of
the map, and of the #24 holdout.

## Acceptance rules

The thresholds are the frozen contracts; no new threshold is chosen.

| Id | Rule |
| --- | --- |
| H1 | Every row returns a number within its governing contract: `all` P `3E-15`, Q `5E-15`; `large_shape` `2E-13`; `extreme_shape` P `2E-12`, Q `5E-13`. The Chi-square rows are held to the Gamma bounds. |
| H2 | The Gamma and Chi-square routes return the same Double at every point. |
| H3 | Per shape, over increasing X, including the boundary probes, the CDF never decreases and the survival never increases. |
| H4 | `abs(P + Q - 1) <= bound_P * P + bound_Q * Q` at every point. |
| H5 | At every boundary probe triple, the values on the two sides differ by no more than the governing contract allows. |
| H6 | Measured by the #298 benchmark on the same runner and in the same run, every point the rule routes to binary64 costs under 0.2 ms per call. |

Any failure rejects the rule.
