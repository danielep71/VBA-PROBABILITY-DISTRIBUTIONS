# Positive-ratio arms on the integrated Gamma kernel (#13)

Recorded 2026-11-13 (Europe/Rome). This file is an addition. The
preregistration, `HOLDOUT_DESIGN.md`, `POST23_RESULT.md`, `MOVEMENT_MANIFEST.md`
and the committed study CSVs are unchanged.

## Why the arms were run again

`POST23_RESULT.md` replayed the four frozen arms on `f157e67`. The public
Gamma integration replaced the incomplete-gamma kernel after that replay, so
the "current" column of the replay no longer described production.

The first Excel run of the production wiring showed it: run 43480835038 on
`cd5f835` applied the registered cutoffs and failed six existing bit-level
assertions of the cumulative at `X = 2^-1074` (PF 000, PF 014, PF 028, for
Gamma and Chi-square). The recovered values differed from the committed
correctly rounded ones by 2, 129 and 145 ulp.

## The measurement

Excel run 43482488530 on `release/1.0.0` at `5556c5b`, before any wiring,
exported the four arms. The files are in `post_integration/`:

| File | SHA-256 |
| --- | --- |
| `gamma_probe.csv` | `34d88e685d3df549fe390d639916c9bee869ebba6b1be249e65aae1cd38042ca` |
| `chisquare_probe.csv` | `0ee7feb86dbd5aa756b84c2ef1d9efcc20b4df23f13ec5512335b07102a606ba` |
| `gamma_holdout.csv` | `df02c56a2345f5a5dce1607623705b22336b4155ea0b5eefb885bebf7584167d` |
| `chisquare_holdout.csv` | `ef32f65266a802ebb0fd70fd42791c846d6a8640616c070131019235e96b9733` |

The hashes are those of the files in the run's artifact, which have CRLF line
ends; Git stores them with LF.

Compared with the committed arms:

- **Density.** No current value and no candidate value changed, in any arm.
  Both density claims pass.
- **Cumulative and survival.** The candidate values are unchanged. The
  current values changed in most rows, because the kernel changed.
- **Claims on the holdouts** (`analyze_positive_ratio_subnormal.py --claims`):
  Gamma density PASS, Gamma cumulative PASS at 40 bits, Chi-square density
  PASS, **Chi-square cumulative REJECTED at 48 bits** (bucket 46: worst
  current 1.29E-14, worst candidate 4.87E-14).

## What the cumulative now depends on

Worst relative error of the Gamma cumulative on the holdout, by retained
bits and by how the argument was built. `landmark` rows have scale 1, so the
stored argument is exact; `transform_stress` rows round the ratio.

| Bits | landmark: current | landmark: candidate | transform: current | transform: candidate |
| ---: | ---: | ---: | ---: | ---: |
| 50 | 9.0E-17 | 9.1E-15 | 8.1E-16 | 1.1E-13 |
| 46 | 5.5E-17 | 7.9E-14 | 1.3E-14 | 9.6E-14 |
| 41 | 8.4E-17 | 3.0E-14 | 4.1E-13 | 7.4E-14 |
| 37 | 1.1E-16 | 2.2E-14 | 6.6E-12 | 1.3E-13 |
| 33 | 7.1E-17 | 5.3E-14 | 1.1E-10 | 1.2E-13 |
| 25 | 8.4E-17 | 8.0E-14 | 2.7E-08 | 1.3E-13 |
| 13 | 7.7E-17 | 6.5E-14 | 1.1E-04 | 1.5E-13 |
| 3 | 7.3E-17 | 4.3E-14 | 0.10 | 1.0E-13 |

The current kernel is exact for the argument it receives. When the transform
is exact, the log-domain candidate is worse by a factor of several hundred at
every bit count. When the transform rounds, the current error is the rounding
of the ratio and grows to 10% at 3 bits, while the candidate stays near
1E-13. Retained bits alone do not separate the two cases. The Chi-square
cumulative shows the same pattern.

## Decision for the production wiring

Taken by the maintainer on 2026-11-13, in two steps.

1. Wire what this measurement confirms:
   - a positive ratio stored as zero is recovered on all six surfaces;
   - the two densities use the 48-bit precision dispatch.
   The cumulatives get no precision dispatch. `PROB_PRS_CUTOFF_40BIT` keeps
   its definition and has no caller.
2. Register a new rule for the cumulative as an amendment, before it is
   wired: recover a stored positive argument only when the transform rounded.
   It needs a fresh holdout, because both existing holdouts have now been
   seen on the current kernel.

Until step 2, the cumulative keeps a known limit: when `X / ScaleParam`, or
`X / 2`, is a subnormal whose rounding lost bits, the result carries that
rounding. The table gives its size.
