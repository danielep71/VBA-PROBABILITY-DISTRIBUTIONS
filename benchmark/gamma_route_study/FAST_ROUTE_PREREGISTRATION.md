# #314: the fast Gamma route — preregistration

Registered 2026-11-14 (Europe/Rome), **before the routing holdout is
evaluated**. `holdout_route_grid.csv` is committed with its design, probes and
references only; every observed field is blank. No production source changes
in the commit that carries this file.

This file is an addition. `HOLDOUT_DESIGN.md`, `ROUTING_RULE.md` and the
records of the rejected rule are unchanged, and that rule stays rejected.

## Why a new rule can be registered

The rule of `ROUTING_RULE.md` was rejected for two reasons, neither of them
accuracy (`PHASE_B_PREFLIGHT.md`, `TRANSITION_DEVELOPMENT.md`):

- strict monotonicity between adjacent Doubles: the binary64 route runs
  backwards by one or two ulps at a few places, across its seam and inside
  each routine;
- a limit of 0.2 ms per call, which excluded every shape above 10,000.

Two decisions since then remove both grounds:

- **#327** holds results to a guarantee tier, `max(target, 1E-13)` for forward
  relative error, and gates monotonicity as an invariant with an allowance: a
  sequence may not run backwards by more than the two rows' error allowances.
- **The maintainer decided on 2026-11-14** that the fast route is registered
  under that invariant, not under strict monotonicity, and that the public
  bit-level assertions are realigned to the guarantee tier.

## The rule

Two new entry points, `PROB_TryGammaRoutedP` and `PROB_TryGammaRoutedQ`,
replace `PROB_TryGammaRegularizedP` and `_Q` at their callers: the Gamma,
Chi-square and Poisson cumulatives and survivals, and the Gamma inverse
solver. The double-double entry points themselves do not change, and the
generated public Gamma block stays as the Gamma study left it.

For finite `A > 0` and `X > 0`, with `fl(A + 1)` formed as a Double before
the comparison, as the fit wave established:

| Region | Binary64 routine | P returns | Q returns |
| --- | --- | --- | --- |
| `X < fl(A + 1)` | lower series, value `S` | `S` when `1E-290 <= S <= 1` | `1 - S` when `0 <= S <= 0.9` |
| otherwise | continued fraction, value `C` | `1 - C` when `0 <= C <= 0.5` | `C` when `1E-290 <= C <= 0.5` |

Everything else is delegated to the double-double entry point, with its
existing cap fallback and error policy: any other value, a routine that
declines or raises, a non-finite or non-positive argument, and `X = 0`. A
declined routine's diagnostic is cleared before the delegation.

Three constants, fixed here:

- **Floor, 1E-290.** A directly returned tail must be at least this, so a
  result near the bottom of the Double range keeps the kernel's behaviour. A
  complemented tail has no floor: `1 - value` is then exactly one.
- **Lower complement, 0.9.** Below the seam, `Q = 1 - S` loses relative
  accuracy as `S` approaches one. With the series within 4E-15 relative, the
  complement is within 3.6E-14 while `S <= 0.9`. Above it, Q is the #24
  regime and stays with the kernel.
- **Upper complement, 0.5.** From the seam the continued fraction returns the
  smaller tail; a larger value is not trusted.

There is **no shape ceiling**. The binary64 routines are those of
`b64_route_vba.py`, unchanged: they decline where their arithmetic would
leave the normal range and at their cap of 100,000 iterations, and a decline
delegates.

## Development evidence, on revealed data only

`replay_fast_route.py` and `fast_route_development.json`. The model is
`fast_route.py` over `route_model.py`, which matched Excel bit for bit at all
392 fit points.

| Set | Scored values | Beyond their bound | Worst relative error |
| --- | ---: | ---: | ---: |
| #307 fit grid | 779 | 0 | 3.8E-15 |
| Main grid: Gamma, Chi-square, Poisson P and Q | 163 | 0 | 4.0E-15 |
| Consolidated holdout: the same functions | 6 | 0 | 6.6E-16 |
| Seeded random sweep, shapes 1E-300 to 1E5 | 2,904 | 0 | 2.4E-15 |

The bound is each row's guarantee for the grids and 1E-13 elsewhere.

## The holdout

The 117 points of `holdout_design.py`, their routes and the method of their
references are unchanged. `holdout_route_build.py` adds the boundary probes
that `HOLDOUT_DESIGN.md` left to be constructed from the rule:

- **Transitions.** For every holdout shape, every change of the rule's pair
  (source of P, source of Q) along X: the last Double of the old state, the
  first of the new one and its successor. They are located from the rule's
  predicate alone, on a ladder of powers of two from `2^-1074` to `2^1020` in
  steps of `2^6`, with `fl(A + 1)` and the shape's holdout points added, each
  differing interval bisected on the integer bit pattern. This finds the
  seam, both complement limits, the floor on either tail and the decline
  regions. A change that begins and ends between two ladder points is not
  found.
- **Zero.** `X = 0` at every holdout shape.
- **Shape probes.** At the internal boundary `A = 1`: its predecessor, 1 and
  its successor, at every small-shape argument of the design and at their own
  seam triples. These are boundary probes, not independent shapes.

Size: 371 rows (117 holdout, 210 transition, 17 zero, 27 shape probes). At
every row the exporter `M_STATS_PROBDIST_GRH.bas` records
`Gamma_Cumulative(X, A, 1)`, `Gamma_Survival(X, A, 1)`,
`ChiSquare_Cumulative(2X, 2A)`, `ChiSquare_Survival(2X, 2A)` and the two
double-double entry points, each as the 16 hexadecimal digits of its Double.
The inputs travel as bit patterns too, so Excel parses no decimal.

## Acceptance rules

They amend H1 to H6 of `HOLDOUT_DESIGN.md` as stated; `holdout_route_analyze.py`
implements them. One bound serves every row: **1E-13 relative**, the forward
floor of #327. It is at least as strict as the guarantee of every contract
these functions have.

| Id | Rule | Amends |
| --- | --- | --- |
| R1 | Every public tail whose reference is at least 1E-290 is within 1E-13 relative of it. Where the reference is an exact 0 or 1, the value is that Double. | H1: the guarantee tier of #327, in place of the old targets |
| R2 | The Gamma and the Chi-square route return the same Double at every row. | H2, unchanged |
| R3 | Per shape, over increasing X and including the probes, the CDF does not fall and the survival does not rise by more than 1E-13 times the sum of the two references. | H3: the invariant of #327, in place of strict monotonicity |
| R4 | `abs(P + Q - 1) <= 1E-13 * (P + Q)` at every row. | H4, with the bound above |
| R5 | A tail whose reference is below 1E-290 is returned by the kernel: the public value equals the double-double entry point's, bit for bit. | new: the floor of the rule |
| R6 | Every public route returns a number at every row. | part of H1 |

H5 is implied by R1: two values within their allowances cannot differ across
a boundary by more than the allowances permit. H6, the limit of 0.2 ms per
call, is withdrawn as a rejection rule by #314; cost is measured with the
#298 benchmark and published.

A reference within one percent below 1E-290 is scored by neither R1 nor R5:
there either path may answer.

## Predictions, fixed before evaluation

1. R1 to R6 hold on all 371 rows. The worst scored error is below 1E-14.
2. Wherever the model does not delegate, the public value equals the model's
   bit for bit. This is reported, not gating: the references decide.
3. In the main grid, 87 of the 163 scored Gamma, Chi-square and Poisson
   cumulative and survival rows change, by at most 23 ulps, and 3 of the 6 in
   the consolidated holdout, by at most 4. Rows of their inverses may change.
   Every changed row stays within its guarantee. No row of any other function
   changes, and no row of the positive-ratio regimes of #13 changes.
4. A call the rule answers in binary64 costs microseconds up to shape 10,000,
   about 0.2 ms at 1E5 and about 2 ms at 1E7. Lower-region Q with `S > 0.9`,
   and tails below 1E-290, keep the kernel's cost.

## Rejection policy

Any failure of R1 to R6 rejects this rule. The constants, the predicate, the
bound and the probes are not retuned against the revealed holdout, and the
double-double route stays in production. A later candidate needs its own
registration and its own untouched evidence.

## Order of work

1. This file, the model, the replay, the holdout grid with blank
   observations, the exporter and the analyzer are committed and pushed.
2. The routines are regenerated into production, the routed entry points are
   added and the callers repointed. The public bit-level assertions and the
   strict seam checks are realigned to the guarantee tier; the bit-level
   assertions of the double-double entry points stay.
3. One Excel wave exports the main grid, the consolidated holdout and the
   routing holdout. `holdout_route_analyze.py` gives the verdict.
4. Cost is measured and recorded in the README.
