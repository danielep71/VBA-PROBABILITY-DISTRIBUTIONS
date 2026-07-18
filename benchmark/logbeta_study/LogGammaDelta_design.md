# Deferred design: stable `PROB_LogGammaDelta` for unbalanced LogBeta

Status: **banked, not started.** Build only after pending work is pushed and
`RunAll` is confirmed green. This is a validated feature project, not a quick fix.

## Problem (measured, not theoretical)

`PROB_LogBeta` fails the 5E-15 Beta claim for moderately/strongly unbalanced,
non-half-integer arguments. Benchmark `benchmark/logbeta_study/` shows:

- general three-log-gamma identity cancels: error ~ `macheps * (Large/Small)`,
  crossing 5E-15 near ratio 1E-2, reaching a few percent near 1E-14;
- the current one-term asymptotic only fires at ratio <= 1E-15 and is itself
  inaccurate at moderate ratios (1.3% at 1E-1, 5.7E-5 at 1E-3);
- so **no single threshold** on the current two methods closes the band
  1E-2 .. 1E-13. Confirmed by the "best achievable" overlay.

Directly exposed: Beta density; incomplete-beta for arbitrary Alpha, Beta; F with
disparate df. **Student-t is largely protected** — its `Beta(df/2, 1/2)` trips the
half-integer shortcut (B=1/2 exactly) for df>=2; df=1 is Beta(0.5,0.5), balanced.
(Verified.)

## The right abstraction

Isolate the ill-conditioned operation as a stable log-gamma increment:

```
PROB_LogGammaDelta(LargeArg, Increment) = LogGamma(LargeArg + Increment) - LogGamma(LargeArg)
```

Then `LogBeta(Large, Small) = LogGamma(Small) - PROB_LogGammaDelta(Large, Small)`.

Verified in principle: computing the delta WITHOUT forming/subtracting the two
large log-gammas drops the residual-band error to ~0. This — not more asymptotic
terms — is the fix for the 1E-2..1E-4 band.

## Three-regime dispatch (crossovers set BY BENCHMARK, not guessed)

1. **Balanced** -> direct identity `LogGamma(A)+LogGamma(B)-LogGamma(A+B)`.
   Safe region determined by benchmark (ratio ~>= 0.1; verify, 1E-2 may be too far).
2. **Moderately unbalanced** -> stable Lanczos log-gamma difference (below).
3. **Strongly unbalanced** -> adaptive multi-term Bernoulli/digamma expansion;
   4 terms give ~1E-16 for ratio <= ~1E-4. Adaptive truncation: add terms while
   they shrink, stop below target scale or at smallest term, fixed max-term cap.

## Stable Lanczos difference (regime 2) — derivation to implement & verify

With Lanczos `LogGamma(z) = C + (z-1/2)Log(z+g-1/2) - (z+g-1/2) + Log(A(z))`,
let `T = z + g - 1/2`:

```
LogGamma(z+s) - LogGamma(z)
  = s*Log(T) + (z+s-1/2)*Log1p(s/T) - s + Log(A(z+s)/A(z))
```

Series ratio without subtracting two series: if `A(z)=P0 + Σ Pk/(z-1+k)` then
`A(z+s)-A(z) = -s * Σ Pk / [(z-1+k)(z+s-1+k)]`, so
`Log(A(z+s)/A(z)) = Log1p((A(z+s)-A(z))/A(z))`.

Properties: no subtraction of two `z*log z`-scale values; heavy `Log1p`; **same
Lanczos constants as the existing `PROB_LogGamma`** (reuse, don't re-derive).

## Benchmark / regression additions required

- **Direct kernel grid**: ratio 1E0..1E-18 x non-integer Small {0.25,0.7,1.3,2.5,
  5.75,10.25} x large scales {1E2,1E4,1E8,1E12,1E20,1E50}. Per point retain:
  reference / general / one-term / multi-term / stable-Lanczos / dispatched LogBeta.
- **Exact public identities**: `Beta(1,b)=1/b`; `PDF(x;1,b)=b(1-x)^(b-1)`,
  `CDF=1-(1-x)^b`, `SF=(1-x)^b` via `Log1p(-x)` reconstruction; b in
  {1E4,1E8,1E10,1E12}.
- **F grid**: (1,1E4),(2.5,1E8),(10,1E10),(1E6,3); CDF, survival, inverse
  round-trips, complement identities.

## Accuracy claim (interim — path 1, already drafted)

Until this kernel ships, the Beta/F/(non-half-integer) accuracy claim is scoped:
holds for balanced-to-moderately-unbalanced args; degrades for extreme imbalance.
Documented in `benchmark/README.md` and the `PROB_LogBeta` policy block
(claim_scope/ deliverables). This banked design is the eventual full correction.

## Build preconditions

1. Pending work pushed and imported.
2. `RunAll` confirmed green on that baseline.
3. Then build regime 2 first (it closes the residual band), verify each seam
   against mpmath at 50 digits, place crossovers from the benchmark envelope.
