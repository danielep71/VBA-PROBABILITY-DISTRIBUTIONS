# #307 binary64 route map: design (2026-11-10)

Frozen before the Excel export that fills it. The inputs are in `design.py`.
The map is a **fitting set**: it is where a routing rule will be derived, so
it cannot also validate that rule. The rule will be frozen afterwards and
tested on new, untouched points, with seam and monotonicity checks.

## Why

Since #242 every public Gamma P and Q call runs the double-double kernel, at 5 ms
to 170 ms per call (#298). The binary64 routines it replaced still exist in
production, `PROB_TryGammaSeriesP` and `PROB_TryGammaContinuedFractionQ`, and
take 2 to 39 microseconds. Analysis of the bit-exact Python model shows that a
double-double call costs 457 to 11,924 double-double operations. Optimizing the
VBA would gain a few times at most; only routing can gain about a thousand
times. Routing means using binary64 wherever it meets the contracts, and
double-double only where it does not.

The pre-#242 failures were measured in regions:

- lower-region Q at small shapes (#24);
- `X = 2` at small shapes (#26);
- some seams;
- the deep tail at `A = 100`;
- one inverse at `A = 0.01`.

This map measures the binary64 route itself, in Excel, across the domain.

## Inputs

- **Shapes A:** 25 values from 1E-6 to 1E7, denser near 1.
- **Arguments X:**
  - for A below 1: `1E-6, 1E-3, 0.05, 0.5, 2, 5, 10, 30`;
  - for A from 1: `X = A + k sqrt(A)` with `k = -10, -6, -3, -1, 0, 1, 3, 6, 10, 20` (positive X only);
  - for every A: the binary64 predecessor of `A + 1`, `A + 1` and its successor.
- **Total:** 282 points.

## What Excel records

At every point the exporter `M_STATS_PROBDIST_GBR.bas` calls both binary64
routines and records their status and value. It also echoes the parsed inputs:

- `PROB_TryGammaSeriesP` (lower series);
- `PROB_TryGammaContinuedFractionQ` (upper continued fraction).

It does not call the public functions or the double-double kernel.

## References

P and Q at the binary64 inputs, from `benchmark/_igamma.py`:

- the tail on the converging side is computed directly;
- the other tail is its complement at high precision;
- both are evaluated at 40 and 60 digits, and the two readings must agree to 1E-30 relative.

## How it is scored

`analyze_binary64_map.py` forms the value the pre-#242 adapter returned:

- below `A + 1`, P from the series and `Q = 1 - P` in binary64;
- from `A + 1`, Q from the continued fraction and `P = 1 - Q`.

It also scores each routine alone wherever it converged. Errors are relative
to the references. Each point is classed against the contract that governs
it:

| Regime | Applies to | P bound | Q bound |
| --- | --- | ---: | ---: |
| `all` | every point | `3E-15` | `5E-15` |
| `large_shape` | shape 1E2 to 1E6, X within 1 sd of the mean | `2E-13` | `2E-13` |
| `extreme_shape` | shape above 1E6, X within 1 sd of the mean | `2E-12` | `5E-13` |

A reference below 1E-300 is reported apart, as underflow, not scored as a
relative error.

## Use

The map yields an exploratory description: where the binary64 route meets
the governing contract, and with what margin. A routing rule derived from it
will be preregistered with its own untouched holdout. No contract or threshold
changes because of this map.
