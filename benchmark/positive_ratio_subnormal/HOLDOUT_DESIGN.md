# Holdout design specification — Phase A, #13

**Frozen before any holdout result exists.** No holdout output has been
generated, inspected, or analysed at the time of this commit. The exporters are
mechanical once this is fixed; nothing below may be revised in response to a
measured error.

That statement is the point of the document. A holdout chosen after seeing
which points embarrass a candidate is not a holdout, and there is no way to
demonstrate the absence of that after the fact. So the specification is
committed first and the evidence is committed second.

## What the holdout is for

Every figure in Phase A1 and A2 is **fitting evidence**. In this repository
every strict threshold has been rejected on holdout evidence, and an
independent holdout has found a worse point than its fitting grid on every
occasion it has been run.

**The holdout's job is to test the provisional fitting claims, not to discover
a better cutoff.** The claims under test:

| claim | from |
| --- | --- |
| Gamma density — candidate dominates at ≤ 48 bits | A1 |
| Gamma cumulative — candidate dominates at ≤ 40 bits | A1 |
| Chi-square density — candidate dominates at ≤ 48 bits | A2 |
| Chi-square cumulative — candidate dominates at ≤ 48 bits | A2 |
| Survival — no dispatch claim | A1, A2 |

If the holdout indicates 46 or 41 instead of 48 or 40, **that is a rejection of
the fitted cutoff, not an invitation to retune**. The response is to return to
design, revise the hypothesis, and if necessary construct a further untouched
confirmation set — never to adopt the holdout's own crossover as the answer.

Survival is characterised here, not frozen. The A1 landmark/stress
decomposition showed two superimposed mechanisms, and that separation is
unfinished.

## Buckets

```
50, 46, 41, 37, 33, 29, 25, 21, 17, 13, 9, 5, 3
```

Thirteen, none shared with the fitting ladder (`52, 48, 44 … 4, 2, 1`), and
**eleven of thirteen odd**, so they occupy binary exponents the fitting set
never touched. Disjointness is structural rather than merely different: a
holdout that picked other round numbers on the same even ladder would not test
what this is for.

Verified before freezing: at every one of these buckets each Gamma landmark and
1/3, 1/2, 2/3 ULP stress, and each Chi-square landmark and midpoint tie, is
exactly representable, lands on its declared offset, and carries exactly `k`
significand bits. The construction recipe had previously only ever been
exercised on the even ladder.

## Shapes — density and cumulative

Eight, derived from irrational constants so their mantissas carry no structure
the fitting set shares. Six of eight have the low mantissa bit set.

| provenance | frozen binary64 |
| --- | --- |
| fine structure α | 0.0072973525693 |
| 1/(1000e) | 0.00036787944117144 |
| √2/100 | 0.014142135623731 |
| ln2/10 | 0.069314718055995 |
| π/10 | 0.31415926535898 |
| φ − 1 | 0.61803398874989 |
| √2/2 | 0.70710678118655 |
| e/3 | 0.90609394260742 |

## Shapes — survival

Survival needs its own set: five of the eight above are useless for it. The
lower tail falls below binary64 resolution relative to 1 from ln2/10 onward, so
`Q` rounds to exactly 1 across the whole ladder and the observable carries no
information.

**Eligibility criterion, fixed here and applied to the exact reference at every
holdout bucket:**

```
min(P, Q) > 2^-20   (9.54E-07)
```

This is a statement about conditioning and representability alone. It says
nothing about which branch performs better, and it could not — it is evaluated
on the high-precision reference, before any implementation is run.

The distinction matters and is the reason the criterion is written down:
choosing points so the observable carries information is legitimate; choosing
points by which algorithm wins is not.

| provenance | frozen binary64 | P across k=50…3 | Q |
| --- | --- | --- | --- |
| 1/(1000e) | 0.00036787944117144 | 0.761 – 0.770 | 0.230 – 0.239 |
| √2/2000 | 0.00070710678118655 | 0.592 – 0.605 | 0.395 – 0.408 |
| π/2000 | 0.00157079632679490 | 0.312 – 0.328 | 0.672 – 0.688 |
| fine structure α | 0.0072973525693 | 0.00444 – 0.00563 | 0.9944 – 0.9956 |
| √2/100 | 0.0141421356237310 | 2.75E-05 – 4.36E-05 | 0.999956 – 0.999972 |

All five pass the criterion with margin. They span a 38× range in shape, and
the first three bracket the tiny-shape region where A1 exposed the separate
downstream mechanism, without repeating any fitting value.

Measured rejections, recorded so the exclusions are not silently convenient:

| shape | P across the ladder | verdict |
| --- | --- | --- |
| ln2/10 | 4.44E-23 – 4.25E-22 | saturates |
| π/10 | 4.65E-102 – 1.30E-97 | saturates |
| φ − 1 | 4.03E-200 – 2.24E-191 | saturates |
| √2/2 | 7.16E-229 – 7.23E-219 | saturates |
| e/3 | 4.13E-293 – 2.73E-280 | saturates |

## Chi-square

Kernel comparability is preserved exactly as in A2: the kernel receives `df/2`,
so `DegreesFreedom = 2 × Shape`.

| Gamma shape | Chi-square df |
| --- | --- |
| 0.00036787944117144 | 0.00073575888234288 |
| 0.00070710678118655 | 0.0014142135623731 |
| 0.00157079632679490 | 0.0031415926535898 |
| 0.0072973525693 | 0.0145947051386 |
| 0.014142135623731 | 0.028284271247462 |

Verified disjoint from the A2 survival df set.

## The frozen binary64 is authoritative

The mathematical expressions above are **provenance for how each value was
selected**, not definitions. The frozen decimal is what the exporter writes and
`echo_shape` / `echo_df` is what the analyzer computes from.

Two reasons. Computing `Sqr(2#)` or `Exp(1#)` at run time would put another
host-library evaluation inside evidence identity. And a truncated literal is a
different Double: `0.00036787944117` and `0.00036787944117144` do not denote
the same value, and an earlier draft of the exporter used the shorter one while
claiming it was 1/(1000e).

Every value above is verified to round-trip exactly through `Val`.

## Decimal twins

Moved off the fitting set's k=8 and k=32 to **k=9 and k=33**, both verified to
round-trip:

```
k=9   1.26480805335359E-321
k=33  2.12199579096527E-314
```

## Expected volume

| arm | slices | constructions | rows |
| --- | --- | --- | --- |
| Gamma | 8 + 8 + 5 = 21 | 13×4 + 3 + 2 = 57 | 1,197 |
| Chi-square | 8 + 8 + 5 = 21 | 13×2 + 1 + 2 = 29 | 609 |

Gamma builds four constructions per bucket, Chi-square two — halving admits
only an exact result or a midpoint tie. An earlier revision of this table gave
29 and 15, having applied the Chi-square count to Gamma and halved it again.
The point set was never affected; only this predicted volume was wrong, and the
Gamma export reports 1,197.

## What is not frozen here

The dispatch cutoffs. Phase A remains measurement only: no contract, no
threshold, no production source change. The holdout tests four provisional
claims and characterises a fifth surface; it does not authorise a patch.
