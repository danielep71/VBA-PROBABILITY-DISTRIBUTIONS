# CDF large-shape study

Measures the accuracy of the regularized incomplete **gamma** and **beta**
kernels — and therefore of every cumulative and survival probability built on
them — at shapes far beyond the frozen contract grid.

## Why this study exists

CR-P1-02 found that both kernels formed their prefactor by literal subtraction:

```
-X + A*Log(X) - LogGamma(A)            (incomplete gamma)
A*Log(X) + B*Log(Y) - LogBeta(A, B)    (incomplete beta)
```

Each subtracts two quantities of size `A*Log(A)` to leave a modest logarithm.
`PROB_LogGamma` carries a *relative* error contract, so the absolute error of
the difference grows with the shape: measured at 2.0E-03 by A = 1E12 and e^46 by
A = 1E16. Every probability built on the kernel inherits that error, silently.

The kernels now route through the stable Loader log-density helpers, with a
measured regime dispatch for the beta factor (`PROB_IBETA_LOADER_MIN_SHAPE`)
because below the crossover the literal form is at least as accurate and is what
the frozen tiny/unbalanced contracts were validated against.

The frozen contracts reached only shape 3 for Gamma and 1E5 for Beta, so the
repaired regime had no coverage. This grid supplies it.

## References

`generate_cdf_large_shape.py` computes every reference with mpmath at
`40 + log10(A)` digits, using the route that converges — the ascending series
below `A + 1`, the Lentz continued fraction above it — and then **recomputes 15
digits higher and requires agreement to 32 significant digits**. A value that has
genuinely converged does not move; one that hit an iteration cap or ran out of
digits does. Points that fail the check are dropped rather than trusted, and the
generator aborts before writing if the worst disagreement exceeds tolerance. The
survival is computed directly on the side where the continued fraction
converges, never as `1 - CDF` in the tail. Beta references are cross-checked
against the symmetry `I_x(a,b) = 1 - I_{1-x}(b,a)`.

**Amendment, 2026-11-05 (#159).** Until this date the references were computed
at a fixed 30 digits and checked at 45. The log prefactor subtracts terms of size
`A·log(A)`, so at shapes of 1E6 and above the 30-digit value kept only 17 to 24
significant digits, and the survival was formed as `1 - CDF`. The generator now
works at the precision above, and the committed grid was re-referenced with
`python generate_cdf_large_shape.py --merge-references --write`, which changes
reference cells only: all 96 references moved, by at most 4.9E-18 relative, and
no observation changed. The old values are in `../study_reference_rebase.csv`.
At the new precision the check also accepts 24 points the old check dropped (all
Gamma points at A = 1E12, three at 1E10, the centre at 1E8, and the Beta centre
at 1E8 and above). They are not in the grid: adding a point needs an Excel
export. The merge lists them.

## Reachability, and why some points are expected to fail

The lower-tail series needs O(sqrt(A)) terms: about 6,100 at A = 1E6 and 57,600
at A = 1E8, against a cap of `PROB_GAMMA_MAX_ITER = 100000`. Somewhere above
A = 1E8 the series therefore reaches the cap and the kernel returns `#NUM!`
rather than a partial sum. That is the Try-contract behaving correctly, and an
`ERROR` cell in this grid is a measurement, not a defect: it records where the
validated domain stops. The upper-tail continued fraction is far cheaper (about
280 iterations even at A = 1E12) and reaches much further.

## Files

| file | role |
| --- | --- |
| `generate_cdf_large_shape.py` | builds the grid with self-checked mpmath references |
| `cdf_large_shape_grid.csv` | the grid; `observed_vba` filled by the macro |
| `M_STATS_PROBDIST_CDFLS.bas` | `Export_CdfLargeShape` writes `observed_vba` from Excel |
| `analyze_cdf_large_shape.py` | reports worst relative error per family and shape |

## Procedure

1. `python generate_cdf_large_shape.py`
2. Import `M_STATS_PROBDIST_CDFLS.bas`, then run `Export_CdfLargeShape`
3. `python analyze_cdf_large_shape.py`
4. Promote the measured points into the main grid and freeze contracts
