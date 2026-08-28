# Predicted source movement for #23

Frozen before the production source edit.

## Existing real-VBA evidence

For `benchmark/positive_ratio_subnormal/gamma_q_complement.csv`:

- `x05_s0` through `x05_s10`:
  - raw `series_p` and public `cdf` must move and must satisfy the existing
    `3E-15` relative CDF contract;
  - public `survival` is expected to move materially but remain non-compliant
    until #24 removes `Q = 1 - P`;
- `x05_s11` (`Shape = 0.5`) is expected to remain on the existing path unless
  real-Excel scope evidence justifies a broader boundary;
- `xsub_s0` through `xsub_s10` may move because they also enter the series
  branch at a small shape; the worst-case error envelope must not regress;
- public `cdf` and `survival` for `x2_s0` through `x2_s11` must not move,
  because `X = 2` selects the continued-fraction path. The raw `series_p`
  diagnostic violates that routine's precondition at those points and is not a
  public movement constraint.

## Main accuracy grid

No committed main-grid row currently reaches the demonstrated
`Shape < 0.5`, ordinary-argument defect. Therefore a no-diff result on the
current grid is predicted and is not acceptance evidence. Representative
post-fix rows must be promoted before closure.

## Unrelated regions

No movement is expected in:

- incomplete-beta functions;
- normal/lognormal functions;
- Student t and F functions;
- discrete functions other than shared incomplete-gamma consumers reached by
  a changed kernel regime;
- Gamma/Chi-square large-shape Loader regimes;
- the continued-fraction-Q fitting ladder tracked by #26.

Any movement outside these predictions must be explained before #23 closes.

## Post-selection explanation

Added with the production source change, after the predictions above were
frozen. The preregistered scope probe showed that a small-shape-only dispatch
would leave the same source defect reachable at ordinary off-grid shapes. The
selected fix therefore changes the off-grid `PROB_StirlingError` construction
for `N <= 15` rather than adding a Gamma-only shape boundary.

Consequently, off-grid Stirling-error observations and Gamma/Chi-square
surfaces that consume them may move even when `Shape >= 0.5`. Exact
half-integers, integer-count discrete paths, and the `N > 15` asymptotic branch
remain unchanged. This is an explained expansion from the original prediction,
not a post-hoc change to the fitting or holdout inputs.
