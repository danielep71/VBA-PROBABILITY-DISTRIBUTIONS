# Gamma performance: accepted scope and retained studies (#307)

Status date: **2026-11-11 (Europe/Rome)**.

The maintainer explicitly accepted the results reached and closed
[#307](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/issues/307).
Further optimization and the ultrafast routing objective are deferred. This
status supersedes older statements in study snapshots that #307 must remain
open or that it remains a release blocker. Frozen protocols and measured
outcomes are preserved, including failed targets and incomplete attempts.

## What users receive

The exact one-entry LogGamma cache was already integrated via #313/#315.
Repeated shape-100 P/Q medians improved about 8.5–12x and the cold shape-100
inverse about 2.2–2.5x in its studies; cold forward costs remained mixed.
See the [cache report](../../benchmark/gamma_cache_study/RESULT.md).

This integration retains the subsequent experiment code and complete evidence.
It does not change any production VBA module, public dispatcher, numerical
contract, reference grid or inverse routine. B4/B5 are study procedures appended
only to temporary workbooks by their dedicated runners. B5's measurements must
not be presented as deployed worksheet performance.

## Experimental outcomes

| Work | Result and disposition |
| --- | --- |
| Exact recurrence reuse (#317) | Independent confirmation missed the frozen >=5% reduction requirement on both P and Q. Not adopted. [Report](../../benchmark/gamma_recurrence_study/RESULT.md). |
| Residual-preserving B4 (#318) | Revealed host fidelity passed; retained as a development reference. Two timing points regressed against production and the ultrafast target was not met. [Design and history](../../benchmark/gamma_route_study/PAIR_CANDIDATE.md). |
| Scalar-pair B5 (#319) | 576 hi/lo arithmetic checks, 11 normalization checks and 1,371 exact inputs passed (1,265 matching P/Q answers, 106 expected declines). All eight warm pair points improved versus B4 in the same run, by 21.5–82.9%. [Complete evidence](gamma_scalar_evidence/run_43070672334/README.md). |

At shape 10,000, B5's paired warm cost is 28.137 ms at X=9900 and 8.212 ms at
X=10100, versus B4's 164.150 ms and 37.170 ms (5.83x and 4.53x speedups).
Across the eight points, B5 costs 3.343–28.137 ms. These are experimental paired
P/Q timings, not cold, public scalar-call or inverse timings.

B5 run 43070672334 passed at source e498658f3e94ef259187ce1d4da8a34992357c2d.
The ordinary library regression separately passed on retry in run 43070683108,
at PR merge candidate bfd138c0e25bb44f18d883bcb7416bc129a39fd8. Its first attempt
failed with RPC errors; no failed measurement is reclassified as successful.
The retained B5 ZIP digest was verified and all retained rows pass the validator.
The snapshot's prose describing #307 as open predates the closure decision.

## Cost target and deferred work

The 100-us target originated as the binary64 fit margin below the 200-us
per-call H6 routing holdout limit. It is not a universal double-double release
criterion. Those frozen rules remain unchanged; B5 still does not meet the
100-us paired comparison target, and this acceptance does not claim otherwise.

The performance blocker is resolved by explicit maintainer acceptance of the
narrower scope, not by claiming an unperformed B5 production rollout. The other
release blockers and all accuracy, monotonicity and certification requirements
remain in force. No v1.0.0 release is published by this integration.

If B5 production promotion is resumed, independent timing confirmation,
validation through public forward/inverse paths and boundary/fallback validation
remain prerequisites. A new routing rule requires preregistration and its own
validation. The reserved 117-point routing holdout remains unevaluated; the
previously rejected routing rule remains rejected.
