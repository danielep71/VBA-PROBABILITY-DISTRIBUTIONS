# Positive-ratio subnormal study — Phase A, Gamma and Chi-square arms

Measurement only for ICR-P1-01A (#13). **No cutoff is frozen, no contract is
proposed, no source is changed, and nothing here transfers to Chi-square.**

Every number below is produced by
`analyze_positive_ratio_subnormal.py gamma_probe.csv`. Nothing is computed by
hand.

## Evidence integrity

- **1,235 observations** — 65 constructions across 19 surface × shape slices.
- **Schema migration verified.** The study was re-based mid-flight when shape
  became part of evidence identity. All 195 legacy shape-0.5 rows reproduce
  **bit-for-bit** across `echo_x`, `stored_standardx`, `log_standardx`,
  `current_value` and `candidate_value` — zero input changes, zero result
  changes. The revision moved no evidence.
- **An oracle defect was found and corrected.** The density reference omitted
  the `1/ScaleParam` Jacobian, producing a constant relative error of
  `1 - 1/scale` at every point sharing a scale. It looked exactly like a
  catastrophic kernel defect at all fifteen buckets. It is now checked at
  startup on every run, and covered by a metamorphic test spanning two shapes
  and four scales.
- **21/21 analyzer unit tests**, including five negative cases that must abort
  the run.

## Construction validation

Landmarks are `q = 2^(k-1) * 2^-1074`, exactly representable, carrying exactly
`k` significand bits. Transform stresses have exact quotients at 1/3, 1/2 and
2/3 of a subnormal ULP above the landmark. Hard underflow is a separate class.
Decimal twins are diagnostic only.

Nothing authoritative is written as a subnormal literal: a VBA source literal
cannot be relied on to denote one, and a mis-parsed literal would silently test
a different number. This was not hypothetical — the first twin comparison
failed on a hand-written constant that was one ULP wrong.

`transform_relative_error` tracks `2^-k` across the whole ladder, from 2.22E-16
at 52 bits to 0.333 at 1 bit. Nothing in the exporter computes that column, so
its agreement is independent evidence that the exact-rational reconstruction is
correct.

## Finding 1 — hard underflow. **Confirmed, not provisional.**

There are valid public inputs for which the mathematical ratio is positive, the
stored standardized argument is zero, and the current implementation is wrong:

| input | surface | current | candidate | true |
| --- | --- | --- | --- | --- |
| `X=1E-300, Scale=1E+300` | density | **#VALUE!** | 0.5642 | 0.5642 |
| `X=1E-300, Scale=1E+300` | cumulative | **0** | 1.128E-300 | 1.128E-300 |
| `X=1E-200, Scale=1E+200` | density | **#VALUE!** | 0.5642 | 0.5642 |
| `X=1E-200, Scale=1E+200` | cumulative | **0** | 1.128E-200 | 1.128E-200 |
| `X=2^-1074, Scale=4` | density | **#VALUE!** | 1.269E+161 | 1.269E+161 |
| `X=2^-1074, Scale=4` | cumulative | **0** | 1.254E-162 | 1.254E-162 |

57 observations in this class. The `1E-300 / 1E+300` case is the strongest:
**neither operand is anywhere near the subnormal domain**, and the true density
is 0.5642 — an unremarkable number. Only the quotient is extreme.

This is a **silent wrong answer for the CDF** — a positive probability returned
as exactly zero — and a **spurious error for the density**. No crossover
analysis can dissolve this requirement: the log path is mandatory for
information recovery whatever the dispatch boundary turns out to be.

## Finding 2 — density. Provisional crossover at 48 bits.

455 rows, **all `normal_output`**, so the public and algorithmic envelopes
coincide. Worst case across seven shapes, stress points only:

| bits | worst current | binds | worst candidate | binds | better |
| --- | --- | --- | --- | --- | --- |
| 52 | 1.14E-13 | s0.1 | 1.14E-13 | s0.1 | current |
| 48 | 1.48E-13 | s0.75 | 8.27E-14 | s0.1 | candidate |
| 40 | 1.02E-12 | s1.0 | 1.34E-13 | s0.001 | candidate |
| 32 | 2.33E-10 | s1.0 | 1.28E-13 | s0.1 | candidate |
| 16 | 1.53E-05 | s1.0 | 5.27E-14 | s0.1 | candidate |
| 1 | 0.333 | s1.0 | 5.32E-14 | s0.1 | candidate |

This is the textbook signature the study was built to detect: landmarks tie
because they carry no transform error, stress points separate the
implementations, current degrades as `2^-k`, and the candidate holds a flat
5E-14 to 1.4E-13 log-domain floor.

The worst current error binds at `s1.0` at every bucket below 44 — the
structural control, where the `z^(Shape-1)` factor vanishes. That the control
is the binding case is itself informative.

**48 is not a production cutoff.** It is a single-arm, fitting-set,
seven-shape result with no holdout.

## Finding 3 — cumulative. Confounded; algorithmic crossover at 40 bits.

455 rows: **389 `normal_output`, 62 `subnormal_output`, 4
`underflowed_output`**.

| envelope | crossover | binds |
| --- | --- | --- |
| **public** — all rows | **none** | — |
| **algorithmic** — `normal_output` only | **40 bits and below** | s0.75 |

The two disagree, and the reason is structural rather than numerical. At
`Shape = 1` the CDF is `P(1,z) = 1 - e^-z ≈ z`, so when `z` is subnormal the
true output is subnormal too — at bucket 32 it is 1.061E-314. The returned
Double must round back onto the same lattice the input lost, so **no branch can
recover anything there**, and both paths show the same relative error as the
rounded standardized argument. Those rows bound the public envelope and hide a
three-order separation elsewhere: at 32 bits the algorithmic figures are
1.75E-10 against 1.17E-13.

Reading the collapsed envelope alone gives "the log path never dominates,"
which is false. It is an artefact of pooling output-limited rows with
algorithmically limited ones.

The output-limited rows are **retained, classified and counted**, not
discarded. They describe what a Double-returning UDF can actually deliver, and
may belong in `numerical_limitations.csv` eventually. They must not choose a
dispatch boundary.

## Finding 4 — survival. Two superimposed mechanisms; no dispatch conclusion.

At `Shape = 0.5` the surface is **`non_discriminating_saturation`**: with `z`
near 1E-308 the lower tail is about 1E-154, so `Q = 1 - P` rounds to exactly 1
across the entire ladder. Both branches return 1 and both are right. That slice
is retained as characterisation and excluded from fitting — it is a real result
about where the public value already saturates, not a failed measurement.

The analyzer mechanically reports a crossover at **52 bits, binding at
`s0.0001`**. That figure is an **algorithm comparison result, not evidence for
the positive-ratio defect**, and must not be read as a #13 dispatch crossover —
see the decomposition below. The four unsaturated slices give a candidate at
**1E-16 to 2E-16 at every bucket** — essentially correctly rounded — against a current path reaching
3.73E-04. The worst current always binds at `s0.0001`, the smallest shape,
consistent with the sensitivity of `P` to quotient rounding scaling as `a`.

**But the mechanism is not what the study was measuring.** Landmarks carry zero
transform error, so any error there is downstream. Ratio of worst stress error
to landmark error, current path:

| shape | 52 bits | 32 bits | 8 bits | 1 bit |
| --- | --- | --- | --- | --- |
| s0.0001 | **1.0** | 3.1 | 1.2E+07 | 5.2E+08 |
| s0.001 | **1.0** | 12.7 | 1.4E+09 | 8.5E+10 |

At 52 bits the ratio is exactly 1.0: the 3.87E-13 error is present with **no
transform error at all**, which is precisely why the mechanically-reported
52-bit crossover cannot be a positive-ratio dispatch boundary — at that bucket
there is no positive-ratio loss to dispatch on. Quotient rounding contributes nothing there. Below
about 32 bits the transform takes over and eventually dominates by ten orders.

So the survival evidence contains **two distinct effects**: a tiny-shape
downstream incomplete-gamma error already present at zero transform error, and
the positive-ratio representation loss this study was designed to measure. One
candidate improving both is not evidence they are one defect, and #13 should
not absorb the first simply because the same rewrite happens to fix it.

Note also that the candidate sits far below the ~1E-12 log-subtraction floor
#13 anticipates. **That floor does not bind here** — another hypothesis this
arm did not confirm.

## Output classes and migration proof

```
Density      455 normal
Cumulative   389 normal / 62 subnormal / 4 underflowed
Survival     325 normal
```

```
schema migration: 195/195 Shape=0.5 legacy rows matched
0 changes in:
  inputs
  stored quotient
  log coordinate
  current result
  candidate result
```

## Against the hypotheses in #13

| surface | #13 hypothesis | measured (Gamma, provisional) |
| --- | --- | --- |
| Cumulative | 44 bits, `2^-1030` | 40 bits algorithmic; public confounded |
| Survival | 36 bits, `2^-1038` | **no dispatch conclusion**; the 52-bit figure is an algorithm comparison, not a transform-driven crossover |
| Density | 52 bits, whole band | 48 bits, transform-driven |

All three differ. The single-axis "bits retained → crossover" model is
incomplete: **output representability is a second axis**, and it inverts the
cumulative result on its own.

## What Phase A1 does not establish

- no frozen cutoffs, on any surface
- no contract rows and no thresholds
- no source change
- **no Chi-square inference whatsoever** — that arm has a different transform
  (`0.5 * X`, inline, no `Try` wrapper), a binary rather than graded rounding
  structure, and a far narrower reachable region
- no holdout: every shape here is fitting evidence

---

# Phase A2 — Chi-square arm

**627 observations**, 33 constructions across 19 surface × df slices.
Construction check passed; both decimal twins matched their constructed
partners.

## Why this is a separate arm, not a re-run

Chi-square never calls `PROB_TryDivide`. All three surfaces compute `0.5 * X`
inline with no `Try` wrapper. Two consequences the evidence design had to
respect:

- **Rounding is binary, not graded.** Halving a binary64 is either exact or
  lands on a midpoint, so only two evidence classes exist. The 1/3 and 2/3 ULP
  stresses used in the Gamma arm are not omitted for convenience — they are
  physically unconstructible, and the analyzer has a test asserting so.
- **Reachability is far narrower.** One free parameter, so the only route to a
  tiny standardized argument is a tiny `X`, and hard underflow needs
  `X < 2^-1073`. Gamma reaches the same state from `1E-300 / 1E+300` — two
  entirely ordinary operands.

`PROB_TF_ValidateDF` also carries a lower guard Gamma lacks: it rejects a df
whose half underflows. Gamma's shape guard has `X > 0` and an upper bound only.

The df slices are chosen so the kernel shapes (`df / 2`) match the Gamma arm
exactly. That is what makes the two families comparable at all.

## Result — the two families do NOT share a cumulative boundary

| surface | Gamma | Chi-square | same? |
| --- | --- | --- | --- |
| density | 48 bits, binds s0.1 | 48 bits, binds s0.002 | boundary yes, binding shape no |
| cumulative, public | none | none | yes, both output-confounded |
| **cumulative, algorithmic** | **40 bits**, s0.75 | **48 bits**, s1.0 | **no** |
| survival | 52 bits, s0.0001 | 52 bits, s0.0002 | yes |

Chi-square shape ids are degrees of freedom: `s0.002` is `df = 0.002`, kernel
shape 0.001; `s1.0` is `df = 1`, kernel shape 0.5.

**This is the finding Phase A existed to produce.** Sharing a downstream kernel
does not imply sharing a dispatch policy — a single cumulative cutoff would
route one family wrongly. Both figures remain provisional.

## Structural confirmations

Two effects discovered in the Gamma arm reproduce independently here, which is
evidence the analyzer measures what it claims:

- **The output-representability confound recurs.** Chi-square cumulative has
  198 `normal_output`, 32 `subnormal_output` and 1 `underflowed_output` rows.
  The public envelope again reports no crossover; the algorithmic envelope over
  normally-representable-output rows gives 48 bits.
- **Survival saturation recurs at the matching shape.** `df = 1`, kernel shape
  0.5, is flagged `NON_DISCRIMINATING_SATURATION` — exactly the Gamma
  Shape = 0.5 result, reached through a different transform.

## A limit of the candidate itself

48 of 627 candidate evaluations return ERROR, all on the density surface at
tiny shapes. There `(a - 1) * Log(z)` reaches about +714 and `PROB_TryExp`
overflows: the candidate has its own output-representability limit, and a log
formulation does not make an unrepresentable density representable.

No bucket lost every candidate, so the reported envelopes stand. But this
exposed a defect in the analyzer, fixed in the same commit: `envelope()`
initialised the worst candidate error to `0.0`, so a bucket where the candidate
never answered would have appeared to win with a perfect score. **Failing to
answer is not the same as answering accurately.** It is now `None`, reported as
`n/a`, and `crossover_of` refuses to call such a bucket a crossover.

## Next

1. Understand why the cumulative boundaries differ between families — 40 versus
   48 bits — before either becomes a cutoff. The candidate formula is identical
   in both arms, so the difference must come from the transform structure
   (graded versus binary) or from the shape grid.
2. Independent shape holdout for both families. Everything measured so far is
   fitting evidence.
3. Isolate the tiny-shape survival mechanism from the transform mechanism, and
   decide whether it is a second finding rather than part of #13.
4. Only then, a provisional dispatch table — per family, not shared.

## Files

| file | purpose |
| --- | --- |
| `M_STATS_PROBDIST_GAMMAPROBE.bas` | exporter; run `Probe_GammaPositiveRatio` |
| `gamma_probe.csv` | 1,235 observations |
| `gamma_probe_shape05_pre_schema.csv` | pre-revision export, kept for the migration check |
| `M_STATS_PROBDIST_CHISQPROBE.bas` | exporter; run `Probe_ChiSquarePositiveRatio` |
| `chisquare_probe.csv` | 627 observations |
| `analyze_positive_ratio_subnormal.py` | schema, oracle and envelopes |
| `test_analyze_positive_ratio_subnormal.py` | 21 tests, including five negative |
