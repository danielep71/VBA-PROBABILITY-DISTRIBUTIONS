# Predicted movement manifest — ICR-P1-01A (#13)

**Committed before the source change.** Every figure below is derived from the
committed grid at `5c884bc` and the dispatch conditions the patch will
introduce. Nothing here may be revised after measuring; the point is to be
wrong in public if the reasoning is wrong.

This exists because of #18. There I predicted 104 candidate rows, 7 moved, and
the single row that did move came from the caller I had set aside as "protected
by its clamp". `CONTRIBUTING.md` now carries the rule that produced: *a caller
protected from an incorrect result by clamping, fallback or error handling is
not necessarily unaffected by the change.*

## The prediction

**Zero rows move. Any movement at all is unexplained and halts the change.**

That is a deliberately brittle prediction. It is falsified by a single moved
observation anywhere in the 2,091-row grid.

## Dispatch conditions

Holdout-confirmed, per family — the two families do **not** share a cumulative
boundary:

| surface | family | route to log path when | cutoff |
| --- | --- | --- | --- |
| density | Gamma | `StandardX < 2^-1026` | `PROB_MIN_NORMAL / 16` |
| cumulative | Gamma | `StandardX < 2^-1034` | `PROB_MIN_NORMAL / 4096` |
| density | Chi-square | `StandardX < 2^-1026` | `PROB_MIN_NORMAL / 16` |
| cumulative | Chi-square | `StandardX < 2^-1026` | `PROB_MIN_NORMAL / 16` |

Both constants derive from `PROB_MIN_NORMAL` by exact division by a power of
two, so **no subnormal literal is ever parsed** — a source literal cannot be
relied on to denote one, and this study has already been bitten by a
hand-written constant that was one ULP wrong.

Survival is **not** in this table. It was removed from the positive-ratio
dispatch decision: its A1/A2 observations turned out to measure a shared
incomplete-gamma series defect and a `Q = 1 - P` amplification, both reachable
at `ScaleParam = 1` where the transform is a no-op. Neither belongs to #13.

Hard underflow — a positive mathematical ratio storing as zero — routes to the
log path unconditionally, independent of any cutoff.

## Reachability, traced not assumed

The changed kernels are `PROB_TryGammaRegularizedP`, `PROB_TryGammaRegularizedQ`
and `PROB_TryGammaLogPdf`. Transitive closure to the public surface gives
**eleven** functions, not the four the cutoffs were measured on:

```
ChiSquare_Cumulative     Gamma_Cumulative          Poisson_Cumulative
ChiSquare_Density        Gamma_Density             Poisson_InverseCumulative
ChiSquare_InverseCumulative  Gamma_InverseCumulative   Poisson_Survival
ChiSquare_Survival       Gamma_Survival
```

The three Poisson tails and both inverses reach the same kernels. Naming them
here is the whole point of the exercise: they are exactly the class of caller
that #18 taught me not to leave out of a prediction.

## Why zero

| function | rows | smallest StandardX | smallest arg1 |
| --- | --- | --- | --- |
| ChiSquare_Cumulative | 34 | 0.25 | 0.5 |
| ChiSquare_Density | 32 | 0.25 | 0.5 |
| ChiSquare_InverseCumulative | 34 | — | 0.001 |
| ChiSquare_Survival | 34 | — | 0.5 |
| Gamma_Cumulative | 14 | 0.5 | 0.5 |
| Gamma_Density | 27 | 0.5 | 0.5 |
| Gamma_InverseCumulative | 2 | — | 0.5 |
| Gamma_Survival | 14 | — | 0.5 |
| Poisson_Cumulative | 8 | — | 3 |
| Poisson_InverseCumulative | 12 | — | 0.05 |
| Poisson_Survival | 8 | — | 3 |

219 rows across all eleven reachable surfaces. The smallest standardized
argument anywhere among them is **0.25**, against a nearest cutoff of
**1.391E-309** — a ratio of about 1.8E+308. No argument on any reachable
surface is below 1E-300.

The remaining 1,872 grid rows belong to functions that do not reach these
kernels at all: the discrete families other than Poisson, the normal and
lognormal families, Beta, F, Student-t, Weibull, Exponential, Uniform, and the
special-function surfaces.

## What this says about the fix

The patch is **branch-additive**. It adds a route taken only below the cutoffs
and leaves every existing path byte-identical above them. If any row moves,
the change was not additive and the reasoning here is wrong.

It also says something uncomfortable and worth stating: **the region being
repaired has no grid coverage whatsoever.** That is the fifth instance of the
same theme in this codebase — the Stirling reciprocal branch, the Expm1
subnormal window, the `envelope_domain` inverse rows, the Exponential/Weibull
far tail, and now this. Each time the gate was green because nothing looked.

Which is why the contract rows in #13's task list are not bookkeeping. They are
the only thing that will make this region visible to the gate afterwards, and
they should be promoted in the same change as the fix rather than deferred.

## Verification protocol

1. Full Excel re-export after the source change.
2. Diff the grid against `5c884bc`. **Expect zero moved observations.**
3. If anything moved: stop, do not proceed to contract promotion, and explain
   the movement before continuing. Do not amend this manifest.
4. Assert zero movement in the discrete families explicitly, Poisson included —
   Poisson reaches the kernels and its absence from the moved set is a
   prediction, not a given.
5. Both hosted workflows green. This is the first change in the #13 chain to
   touch `src/**`, so **Excel VBA Regression will finally fire** — every
   preceding commit was `benchmark/**` only and never triggered it.
6. Only then promote the `positive_ratio_subnormal` contract rows.
