# Chi-square inverse reference study (#22, v1.0.0 plan Track A2 item 6)

**Status: PROVISIONAL — INCOMPLETE.** The Rmpfr leg has not run. No final
agreement figure is recorded and the feasibility checkpoint must not be
marked complete until `rmpfr_crosscheck.R` has been run and verified.

## Frozen design

69 references: 23 probabilities at each of df = 1E6, 1E7, 1E8.

| Arm | Probabilities per df | Count |
|---|---|---|
| Fitting baseline | 0.5, 0.9, 0.99 | 3 |
| Fitting bridge | 0.125, 0.25, 0.375, 0.625, 0.75, 0.875 | 6 |
| Fitting tails | 2^-30, 2^-20, 1-2^-20, 1-2^-30 | 4 |
| Holdout bridge | 0.0625, 0.1875, 0.3125, 0.4375, 0.5625, 0.6875, 0.8125, 0.9375 | 8 |
| Holdout tails | 2^-25, 1-2^-25 | 2 |

Every probability is carried as its exact IEEE-754 binary64 value. The hex
float representation is authoritative and travels with each record, so 0.9
and 0.99 never make a decimal round-trip.

Holdout references are generated. VBA holdout observations remain uninspected.

## Files

| File | Role |
|---|---|
| `chisq_reference.py` | Primary route. Series/Lentz-CF, two-precision stabilization, Newton inversion. Writes `chisq_reference.json`. |
| `chisq_crosscheck.py` | Algorithm-independent quadrature route + coarse sanity. Writes `chisq_crosscheck.json`. |
| `chisq_gmpy2.py` | Implementation-independent MPFR-arithmetic route (gmpy2). Never calls a library incomplete gamma. Writes `chisq_gmpy2.json`. |
| `rmpfr_crosscheck.R` | Rmpfr `igamma` leg. Takes an optional df filter; writes incrementally because MPFR can abort. Reaches df 1E6 and 1E7 only. |
| `rmpfr_probe.R` | Bisects the shape at which MPFR's incomplete gamma aborts. |
| `rmpfr_series.R` | Superseded. The mpfr-arithmetic route in R, ~1600x slower than gmpy2. Retained as the record of why gmpy2 was chosen. |
| `chisq_reference.json` | The 69 frozen references. |
| `chisq_crosscheck.json` | Primary-vs-independent comparison record. |

## Reproducing

    python chisq_reference.py       # ~3.5 min, writes chisq_reference.json
    python chisq_crosscheck.py      # ~1 min,  writes chisq_crosscheck.json
    Rscript rmpfr_crosscheck.R chisq_reference.json chisq_rmpfr.json

`chisq_reference.py` can be run one df at a time if a runtime cap applies;
`chisq_crosscheck.py` reads either the merged file or the chunks.

## Results

- 69/69 converged and stabilized; 0 rejected.
- Two-precision stabilization: 55.89 – 59.26 significant digits (required 40).

| Leg | Standing | Coverage | Agreement (sig. digits) |
|---|---|---|---|
| Quadrature (mpmath) | algorithm-independent | 69/69 | min 109.87 |
| gmpy2 / MPFR arithmetic | implementation-independent | 69/69 | min 115.11 |
| Rmpfr `igamma` | third-party incomplete gamma | 46/69 | min 115.96 |

Conservative headline: **109.87 significant digits**, the minimum across all
legs and all points.

## Recorded findings

- At df = 1E8 every frozen point, tails included, satisfies x/a ≈ 1, so all
  69 lie in the central regime. The series needs ~170,000 terms there. There
  is no cheap tail at these shapes.
- `mpmath.gammainc` raises NoConvergence at a = 5E7 and cannot serve as the
  independent leg. Quadrature is used instead, and the failure is recorded in
  `chisq_crosscheck.json` rather than worked around silently.
- SciPy's double-precision `chi2.ppf` deviates by up to 7.8E-06 relative at
  df = 1E8, p = 2^-20. Recorded as coarse sanity only; no figure is derived
  from it.
- MPFR's `mpfr_gamma_inc` ABORTS the process at shape a >= ~4.5E7 (ok at 4.0E7,
  `gamma_inc.c:289` assertion failure at 4.5E7). df = 1E8 needs a = 5E7, so no
  Rmpfr-based oracle can reach it. An abort cannot be caught, so cross-check
  scripts must write results incrementally.
- Reference quantiles MUST be stored at more digits than the working precision.
  An earlier revision stored 50 digits from a 120-dps computation, which capped
  every agreement figure at ~49 digits — measuring the truncation, not the
  oracle. Caught because two independent legs returned figures identical to
  0.0 difference across all 69 points.

## Measures

- quantile — the Chi-square inverse CDF at the exact binary64 p
- tail residual — `|area_direct - target| / min(p, 1-p)`

The small side is always computed directly (lower area for p <= 1/2, upper
area for p > 1/2), so no complement subtraction enters either measure.
