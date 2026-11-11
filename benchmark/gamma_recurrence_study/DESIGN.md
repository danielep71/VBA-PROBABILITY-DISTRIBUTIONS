# Exact recurrence subexpression reuse (#307)

## Frozen before measurement — 2026-11-11

Baseline production tree: release/1.0.0 at 9f1f8105b9bff50070b25074b8a3174193b4ef82.
The validated DD profile identifies the wide lower series and upper continued
fraction as the shape-10000 bottlenecks. This study changes only temporary
workbook code, leaving production source and all accuracy manifests unchanged.
The untouched 117-point routing holdout is neither opened nor evaluated.

## Candidates

- `baseline`: unchanged production, including the validated LogGamma cache.
- `lower`: compute the DD tuple `AA + N` once per lower-series iteration and
  reuse it in the term denominator and next-ratio denominator. Preserve the
  subsequent DD `+ 1` operation, all rounding stores and convergence checks.
- `upper`: retain the initial DD tuple `(XX + 1) - AA` before the evolving B
  variable changes; reuse that tuple in the backward recurrence and final H.
  Do not replace these DD operations with ordinary Double subtraction.
- `combined`: both transformations, independently reversible.

No validation call, tolerance, cap, arithmetic primitive or public routing rule
is weakened. Do not promote a candidate on timing alone.

## Revealed measurement cases and order

Use the same eight forward (A,X) pairs as #298:
(0.0001,0.5), (0.5,0.25), (2.5,1), (2.5,6), (100,90),
(100,110), (10000,9900), (10000,10100), each through P and Q.
Also measure Gamma inverse at (shape=2,p=.5) and (shape=100,p=.9), scale=1.
Three rounds, four timed calls per forward sample and one per inverse sample,
preceded by one untimed warmup. Use four candidate orders: baseline/lower/upper/
combined in rounds 1 and 3, reversed in round 2. Reload source before each case
and mode to reset private static state. No cold-cache claim is made.
Total: 216 samples, 792 timed calls and 216 warmups = 1008 evaluations.
Compare every call and warmup bitwise with the corresponding baseline result.

Report all samples, per-case medians and baseline/candidate ratios. Primary
performance target: at least 5% median reduction for both P and Q at the affected
shape-10000 point, with no bit mismatch. Results under the target remain useful
but do not establish material acceleration. Three samples and uncontrolled host
load provide preliminary timing evidence, not a guaranteed bound. Retain noisy
or adverse controls; a confirmation run and full Excel regression/private-study
and fresh accuracy exports are required before production promotion.

## Evidence

Retain candidate commit SHA, source and generated-source hashes, Excel host and
build, timing macro and runner, complete samples and artifact identity/digest.
A fail-closed generator pins the baseline and exact replacement sites and checks
that reversing each transformation recovers the production source byte for byte
under LF normalization. A mismatch or runtime error fails the run.
