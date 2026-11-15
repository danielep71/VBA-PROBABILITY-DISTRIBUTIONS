# Amendment 1 — the cumulative on the integrated kernel (#13)

Registered 2026-11-13 (Europe/Rome), **before the second holdout is
exported**. The committed `gamma_holdout2.csv` and `chisquare_holdout2.csv`
hold their key columns only; every measurement field is blank.

This file is an addition. `HOLDOUT_DESIGN.md`, `POST23_RESULT.md`,
`MOVEMENT_MANIFEST.md` and the committed arms are unchanged. The registered
policy stays in force for the two densities and for a stored zero, and both
are in production. This amendment replaces only the precision dispatch of the
two cumulatives.

## What changed

`POST_INTEGRATION_RESULT.md` records it. The incomplete-gamma kernel was
replaced after the cutoffs were replayed. The current kernel is exact for the
argument it receives, so the cumulative loses accuracy only when the
standardized argument was rounded. Retained bits alone no longer decide which
path is better, and the Chi-square cumulative claim at 48 bits was rejected.

## Hypothesis

For a positive standardized argument that is stored as a subnormal:

1. If the transform was exact, the current path is at least as accurate as
   the log-domain candidate, at every retained-bit count.
2. If the transform rounded, the candidate is more accurate at 40 retained
   bits or fewer, in both families.

Fitting evidence, on the current kernel: the two fitting arms in
`post_integration/` (rounded: the candidate wins at every bucket from 40
down, the current path at 48 and 52, and 44 is mixed between the families),
and the two first holdouts, which this measurement has now exposed.

## The rule to be wired if the claims hold

Recover the cumulative when all of these hold; otherwise keep the direct path:

| Family | Condition |
| --- | --- |
| Gamma | `0 < StandardX < PROB_MIN_NORMAL / 4096#` and `ScaleParam <> 1` |
| Chi-square | `0 < StandardX < PROB_MIN_NORMAL / 4096#` and `2 * StandardX <> X` |

The comparison with the cutoff stays strict. A stored zero is recovered
always, as today.

The Chi-square test is exact: doubling a Double is exact, so `2 * StandardX`
equals `X` exactly when the halving did not round. The Gamma test is a proxy:
a quotient can be exact with a scale other than one, and such an argument
will be recovered although the direct path would be better. Claim C4 bounds
what that costs. An exact test of the quotient needs the product to full
precision, which the production arithmetic does not have.

## Registered claims

On the second holdout, cumulative only, over rows whose true result is a
normal Double. "Worst" is the largest relative error in one bucket, over all
shapes. Classes are decided from the echoed operands as exact rationals.

| Claim | Rows | Statement |
| --- | --- | --- |
| C1 | rounded, 40 bits or fewer | the candidate's worst error is below the current path's, at every bucket |
| C2 | exact, every bucket | the current path's worst error is not above the candidate's |
| C3 | rounded, more than 40 bits | the current path's worst error is at most 5E-13 |
| C4 | Gamma, exact with scale other than one, 40 bits or fewer | the candidate's worst error is at most 3E-13 |

The two bounds are stated from analysis, before any second-holdout value:

- C3: with `k` retained bits, rounding moves the argument by at most half a
  subnormal ULP, a relative `2^-k`. The lower tail scales as the argument to
  the power of the shape, and every shape here is below one, so the result
  moves by less than `2^-41`, about 4.5E-13, at 41 bits and above.
- C4: the candidate forms `Shape * (Log(X) - Log(ScaleParam))`. Two
  logarithms of magnitude up to 745, their difference and the product each
  round once; with shapes below one the sum of those bounds is about 2.4E-13.

A bucket without a comparable value makes its claim pending, not passed. A
rejected or pending claim returns this amendment to design; no cutoff is
taken from the second holdout itself.

## The second holdout

Exporters `M_STATS_PROBDIST_GAMMAHOLD2.bas` and
`M_STATS_PROBDIST_CHISQHOLD2.bas`; analyzer `analyze_amendment1.py`.

- **Buckets:** 51, 47, 43, 42, 39, 35, 31, 27, 23, 19, 15, 11, 7. None was
  used by the fitting arms (52 down to 1 in steps of four) or by the first
  holdouts (50, 46, 41, 37, 33, 29, 25, 21, 17, 13, 9, 5, 3).
- **Gamma shapes:** EulerGamma/100, 1/(1000 pi), sqrt3/100, ln3/10, e/10,
  1/sqrt3, pi/4, ln 2.5. None was used before. All are below one, so the
  lower tail is a normal Double at every bucket.
- **Chi-square degrees of freedom:** twice those shapes.
- **Gamma constructions**, with `m = 2^-1074`, `N = 2^(k-1)`, `W = 2^300`:

| Class | X | Scale | Exact quotient |
| --- | --- | --- | --- |
| `exact_unit` | `N m` | 1 | `N m` |
| `exact_scaled` | `4N m` | 4 | `N m` |
| `rounded` | `(5N + 2) m` | 5 | `N m + 2m/5` |
| `rounded` | `(7N + 3) m` | 7 | `N m + 3m/7` |
| `rounded` | `(6N + 1) m` | 6 | `N m + m/6` |
| `rounded` | `(3N + 1) m W` | `3 W` | `N m + m/3` |

  The scales 5, 7 and 6 and their offsets are new. The last row has a normal
  X and a large scale, the form in which the defect is reached from ordinary
  operands.
- **Chi-square constructions:** `X = 2N m` (exact) and `X = (2N + 1) m`
  (rounded to the midpoint, the only rounding a halving can produce).

Size: 624 Gamma rows and 208 Chi-square rows.

## Order of work

1. This registration, the exporters, the analyzer and its fixtures are
   committed and pushed with the blank baselines.
2. The second holdout is exported on that commit. The current path it
   measures must be the unwired cumulative, so the export precedes any wiring.
3. `analyze_amendment1.py` gives the verdict. Only if every claim holds is the
   rule wired, with its own Excel wave and the zero-movement check.
4. The six contract rows of #13 follow.

## Addendum — the Gamma exactness test (2026-11-13)

Added after the result, when the rule was wired. The claims, the cutoff and
the second holdout are unchanged; only the production test of "the transform
rounded" for Gamma is.

The registered proxy, `ScaleParam <> 1`, was wired first. Excel run
43507315705 on `8e3aca6` failed three bit-level assertions of the public
Gamma integration (PF 000, PF 014, PF 028): they call the cumulative with
scale 2 and an exact quotient, and the proxy sent them to the recovery. That
is the cost claim C4 bounded, but those assertions hold the correctly rounded
result and must keep holding.

Production therefore tests the division itself. With
`Product = StandardX * ScaleParam`, stored as a Double, the transform is
taken as exact when `Product = X` and one of these holds:

- `X` is a normal Double. Then the stored quotient is within one part in
  `2^53` of the exact ratio, which the direct path cannot distinguish from
  exact.
- `ScaleParam` is an integer. Then the product of a subnormal quotient is
  exact, so equality means the division was exact.

A subnormal `X` with a fractional scale is treated as rounded: there a
rounded quotient can multiply back to `X` (3.2 ULP stored as 3, and 3 times
2.5 rounds to 8).

The test only narrows the registered rule towards claim C2: it moves
arguments with an exact transform from the recovery to the direct path, which
C2 found better at every bucket. `analyze_amendment1.py` mirrors it
(`wired_path`) and requires, on every row of the second holdout, the direct
path for an exact transform and the recovery for a rounded one at 40 bits or
fewer. All 832 rows agree. A random search over 390,000 quotients below the
cutoff found no argument taken as exact with a transform error above
1.2E-16.

The Chi-square test, `2 * StandardX <> X`, is unchanged.
