# Scaled Gamma integration experiment

2026-11-08, #24/#26. `scaled_gamma.py` connects scaled arithmetic to complete
small-shape P/Q calculations. The frozen v1/v2 models and public VBA remain
unchanged. This is a revealed development experiment, not a fresh holdout.

The explicit domain is positive finite binary64 a <= 1 and x <= 32. Unlike v2,
it admits subnormal a and x. Logarithm first separates the binary exponent;
logGamma1p uses the frozen polynomial coefficients with scaled operations.
The lower series retains scaled a, logP, log1p and expm1 until final rounding.
The upper path uses scaled forward convergence screening and backward fraction
reconstruction, including the full prefactor, shape and fraction product.

Stopping screens are heuristic, not accumulated-error certificates. Reusing
coefficients does not establish their error after the arithmetic change. No
published bound is transferred. The selected test guard of two binary64 ULPs
is a development screen, not a replacement for the public acceptance thresholds.
The 36 selected (a,x) cases compare both tails with direct 160-digit mpmath
incomplete-gamma evaluation; small Q is never obtained as 1-P in the oracle.
This covers minimum-subnormal inputs, the former shape cutoff, the routing
boundary and x=32. It does not cover the full Gamma domain or all boundary
neighbours. Invalid/out-of-study-domain inputs are rejected.

Run `python benchmark/gamma_cf_study/coverage/test_scaled_gamma.py`.
The fixture is registered in the full tooling runner (71 commands).
Next: broader boundary/error analysis, independent oracle stability checks,
private VBA integration and exact-source execution before any public routing
change. Existing Excel evidence remains tied to its original commits; the
41 public failures remain unresolved by this Python-only experiment.
