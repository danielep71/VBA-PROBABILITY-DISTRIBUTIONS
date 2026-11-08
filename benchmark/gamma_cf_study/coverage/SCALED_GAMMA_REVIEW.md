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


## Expanded revealed checks

The integration fixture now also replays all 69 B6 reference points inside
its declared domain; 20 B6 points are outside and are not silently accepted.
Both tails pass the same two-ULP development screen. These are already revealed
reference points, not new independent acceptance evidence.

At eight selected shapes (minimum subnormal, 1e-100, 1e-6, 1e-4, .01,
.25, .5 and 1), the three adjacent binary64 arguments around x=a+1 are
checked for nondecreasing P and nonincreasing Q without a tolerance allowance.
Direct lower/upper mpmath references at 160 and 320 decimal digits give the
same rounded binary64 values for all 24 points. The model agrees within two
ULPs. All four integration test methods pass. This finite check does not prove
monotonicity throughout the public domain, certify accumulated error, or solve
the inverse/caller integration requirements. No public VBA was changed.
