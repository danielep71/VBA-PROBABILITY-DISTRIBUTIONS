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
| `chisq_crosscheck.py` | Independent quadrature route + coarse sanity. Writes `chisq_crosscheck.json`. |
| `rmpfr_crosscheck.R` | **Required, not yet run.** Independent MPFR leg. Writes `chisq_rmpfr.json`. |
| `chisq_reference.json` | The 69 frozen references. |
| `chisq_crosscheck.json` | Primary-vs-independent comparison record. |

## Reproducing

    python chisq_reference.py       # ~3.5 min, writes chisq_reference.json
    python chisq_crosscheck.py      # ~1 min,  writes chisq_crosscheck.json
    Rscript rmpfr_crosscheck.R chisq_reference.json chisq_rmpfr.json

`chisq_reference.py` can be run one df at a time if a runtime cap applies;
`chisq_crosscheck.py` reads either the merged file or the chunks.

## Results so far (mpmath legs only)

- 69/69 converged and stabilized; 0 rejected.
- Two-precision stabilization: 55.89 – 59.26 significant digits (required 40).
- Primary vs independent quadrature: minimum 49.32 significant digits.

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

## Measures

- quantile — the Chi-square inverse CDF at the exact binary64 p
- tail residual — `|area_direct - target| / min(p, 1-p)`

The small side is always computed directly (lower area for p <= 1/2, upper
area for p > 1/2), so no complement subtraction enters either measure.
