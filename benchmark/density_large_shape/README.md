# Large-shape density characterization study (CR-P1-01)

## Why this exists

The continuous density kernels evaluate the log-density in the cancellation-prone
normalization form:

- Gamma: `(Shape-1)*Log(X/Scale) - X/Scale - Log(Scale) - PROB_LogGamma(Shape)`
- Beta:  `(a-1)*Log(X) + (b-1)*Log1p(-X) - PROB_LogBeta(a,b)`
- F / Chi-square: the same `PROB_LogGamma` / `PROB_LogBeta` structure.

At large shape near the mode each term is `~ a log a` while the true log-density
is only `~ log a`, so the leading terms cancel and expose absolute error from the
separately rounded components. Because densities are not enveloped, the public
functions return a finite, silently wrong value in the accepted domain. This
study measures the error across shape/df from `1E2` to `1E20` so the boundary is
committed evidence rather than an assertion.

## References (self-checked, and they validate the fix)

Each reference log-density is computed TWO ways at 60 digits: the direct form and
the reviewer's Loader deviance form (15.2 / 15.3). They must agree to `1E-40`, so
the grid simultaneously (a) certifies the reference and (b) proves the Loader
formulas are algebraically exact. A row whose forms disagree, or whose density
under/overflows Double, is dropped rather than shipped. The stored reference is
the density value; abscissae are rounded to the exact Double the kernel receives.

## Sample structure (reviewer 15.5)

Shapes `1E2, 1E4, 1E6, 1E8, 1E10, 1E12, 1E16, 1E20`; for each family the points
sit in the representable band around the mode (mode +/- z*sd, not deep tail):

- Gamma: `X/Scale` at and a few sd around Shape.
- Chi-square: `X` near `df` (reuses the Gamma helper with Shape=df/2, Scale=2).
- Beta: balanced `a=b`, `x` near `0.5`; unbalanced `a=Shape, b=Shape/100`, `x`
  near the mode `a/(a+b)`.
- F: balanced `df1=df2`, `x` near `1`; unbalanced `df1=Shape, df2=Shape/100`,
  `x` straddling `df2/df1` so both `r<1` and `r>1` log-ratio branches are hit.

## Running it

```
python generate_density_large_shape.py   # -> density_large_shape_grid.csv (observed empty)
# import M_STATS_PROBDIST_DENS.bas, run Export_DensityLargeShape, pick the grid
python analyze_density_large_shape.py     # naive vs stable-Loader, per (function, shape)
```

## How to read the result

The analyzer reports, per `(function, shape)`:

- **naive absLog / relDens** - the current kernel's worst absolute log-density and
  relative density error (from the exported VBA). This is the defect: expect it to
  climb past `~1E-6` well before `1E12` and become catastrophic (`O(1E8)x` density
  error) by `1E16`.
- **loader absLog** - what a stable Loader-in-Double kernel achieves on the same
  points (StirlingError + stable deviance). It stays `<= ~1E-6` across the entire
  range; the gap between the two columns is exactly what the fix recovers.

### Implementation note the study establishes

Gamma / Chi-square need a single stable deviance `bd0(a, y)`. **Beta and F have TWO
large shapes, so their deviance must be the Loader decomposition**
`D = bd0(a, n*x) + bd0(b, n*y)` with each `bd0` evaluated by the stable
`(1+u)*log1p(u) - u` form (`u = (k-m)/m`). Computing the raw
`a*log(a/(n*x)) + b*log(b/(n*y))` re-introduces the cancellation and degrades to
`O(1E4)` at `1E20` even in the Loader arrangement - measured here. The repo's
existing `PROB_DS_TryDeviancePart` (used by the discrete PMFs) is this stable
`bd0`; the fix reuses it rather than adding a new routine.
