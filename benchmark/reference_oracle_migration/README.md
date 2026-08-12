# Reference-oracle migration

One-off audit of a change to how `generate_reference_values.py` constructs
log-Gamma references:

```python
mp.log(mp.gamma(z))   ->   mp.loggamma(z)
```

applied to the four helpers whose requested output is naturally expressed
through logarithmic Gamma combinations: `_loggamma`, `_loggamma_halfdiff`,
`_stirling_error` and `_logchoose`. This folder is the provenance for why a
reference implementation changed even though no committed number did.

## Result

| | |
|---|---|
| audited references | **112** |
| `UNCHANGED_AT_GRID_PRECISION` | **112** |
| `REFERENCE_IMPROVEMENT` | 0 |
| `UNEXPECTED` | 0 |
| oracle self-disagreement (200 digits, two formulations) | none |

`LogGamma` 40, `LogGammaHalfDiff` 30, `StirlingError` 12, `LogChoose` 30.

**Not one committed reference string moves.** That is a stronger result than an
improvement would have been: the old construction was adequate for every point
already in the grid. The migration is required *prospectively*, to make the
forthcoming small-positive and subnormal evidence reproducible at all — not to
revise historical numbers.

## Why the change is needed anyway

`mp.log(mp.gamma(z))` builds `Gamma(z)` as an intermediate. Two consequences the
current grid happens not to reach:

- **Overflow.** `Gamma(z)` diverges as `z -> 0`. The committed `LogGamma` grid
  starts at `1E-8`, where `Gamma(1E-8)` is about `1E+8` and nothing overflows.
  The regime work adds points down to the smallest positive subnormal, where the
  intermediate is unrepresentable while `Log(Gamma(z))` is an ordinary number
  near 744.
- **Cancellation.** `Log(Gamma(z))` is zero at `z = 1` and `z = 2`. Building
  `Gamma(z)` first and taking its logarithm loses significance approaching those
  zeros. The grid's logspace steps never land near them; the `near_zero` regime
  is defined by them.

`mp.loggamma` computes the contracted quantity directly and has neither problem.

## Precision

`with mp.workdps(max(mp.mp.dps, 110))`, with a unary `+` rounding back to the
generator's active precision on leaving the context.

This is a fixed margin, not an adaptive scheme. Exponent-adaptive precision is
needed when a reference formula literally constructs `1 + x` for tiny `x`, as the
`LogGamma1p` study reference does — an `mpf` separates exponent from mantissa, so
`loggamma(4.94E-324)` needs no extra digits merely because its exponent is -324.

## Method

`migrate_reference_oracle.py` loads the old and new generators side by side at
the generator's own precision (50 dps) and calls `build_rows()` on each:

1. **Structure is asserted unchanged** — `function`, `vba_kernel`, all four
   arguments and `regime` identical across all 1346 generated rows.
2. **Scope is asserted** — zero references move outside the four helpers.
3. **Both values are scored against an independent 200-digit oracle**, computed
   two ways (`mp.loggamma` and `mp.log(mp.gamma)`); the two formulations must
   agree, which validates the oracle rather than assuming it.
4. Each row is classified `UNCHANGED_AT_GRID_PRECISION`,
   `REFERENCE_IMPROVEMENT` or `UNEXPECTED`.

## Reading `reference_migration.csv`

`old_abs_error` and `new_abs_error` are distances from the 200-digit oracle, so
they include the truncation of the committed 25-significant-digit reference
string. At `LogGamma(1E+50)` the value is about `1.1E+52`, so a 25-digit string
carries roughly `1E+32` of absolute truncation; the relative column is the
meaningful one there. Both columns are affected identically, so the comparison
between old and new is unaffected.

Relative columns are blank where the oracle is within `1E-8` of zero, since a
relative error against a value approaching zero is not an acceptance metric.

`delta_old_new` is zero on every row: the two reference strings are identical.

## Scope boundary

Deliberately **not** a global replacement of `mp.gamma()` in the generator.
Weibull moments, the Student t normalising constant and every `gammainc` call
legitimately use `mp.gamma` because their reference quantity is Gamma itself,
not its logarithm. Those are untouched.

## Not run here

The generator is **not** re-run against `probability_accuracy_grid.csv`. It emits
`observed_vba` empty by design, so a full regeneration would discard all 2030
committed Excel observations — and separately, the committed generator no longer
reproduces the committed grid (563 grid-only rows, 238 generator-only rows).
That divergence is tracked as its own issue. Since all 112 references are
identical, this migration requires no grid change at all.
