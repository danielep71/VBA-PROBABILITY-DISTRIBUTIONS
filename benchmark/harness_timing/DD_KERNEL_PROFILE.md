# Validated DD kernel: measured profile (#307)

## Evidence and scope

Excel run [43017348872](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/actions/runs/43017348872),
candidate `e7cf3c8f7be6ac408f2e61f13474b1a4f96cdfed`, Excel 16.0 build 20430.
The Windows x64 Excel runner measured the unchanged production modules.
The immediately preceding PR regression, run 43015666373, passed 5987/5987
assertions on merge candidate `604869d62e75208ff6c695d355a14d325d13efb8`,
and independently recorded Office 64-bit. That is separate regression evidence,
not a certification rebound to the profiler commit.

[Raw samples and scope map](run_43017348872_profile.json) retain the exact
candidate, generated-source hashes, artifact identity and SHA-256. The downloaded
ZIP matched GitHub's artifact digest; both generated source files also matched
local regeneration byte for byte. [Measurement procedure](../kernel_profile/README.md).

**PASS: 216 samples, 18 cases, four passes, three samples per case/pass.**
All 1008 evaluations (including warmups) agreed bit for bit with their original
uninstrumented result. No scope-stack mismatch or failed evaluation occurred.
The eight forward points were previously revealed in #298. The two inverse
cases use scale 1. No #307 holdout was read or evaluated; no routing, arithmetic,
production VBA, tolerance, or historical evidence manifest changed.

## Uninstrumented timings

Three warmed samples, four calls per forward sample, one per inverse sample.
These are observed wall-clock costs, not guaranteed bounds. **There is substantial
within-session timing drift:** the first small-shape baseline samples were about
1.4–2.1 times the restored medians. The code was restored unchanged, so this is
NOT a speedup. CPU scheduling, frequency, thermal state or other machine load
were not controlled and their individual effects cannot be inferred. Both
passes are retained rather than choosing a favorable figure.

| Shape | x / inverse probability | Function | Initial median ms | Restored median ms |
|---:|---:|:---|---:|---:|
| 0.0001 | 0.5 | p | 9.058 | 4.939 |
| 0.0001 | 0.5 | q | 10.405 | 4.964 |
| 0.5 | 0.25 | p | 10.685 | 5.887 |
| 0.5 | 0.25 | q | 7.866 | 5.689 |
| 2.5 | 1 | p | 13.812 | 8.766 |
| 2.5 | 1 | q | 13.437 | 8.495 |
| 2.5 | 6 | p | 24.466 | 14.851 |
| 2.5 | 6 | q | 20.172 | 15.638 |
| 100 | 90 | p | 167.231 | 114.856 |
| 100 | 90 | q | 124.822 | 113.246 |
| 100 | 110 | p | 120.542 | 114.311 |
| 100 | 110 | q | 119.457 | 114.928 |
| 10000 | 9900 | p | 131.931 | 128.233 |
| 10000 | 9900 | q | 128.609 | 130.357 |
| 10000 | 10100 | p | 45.877 | 44.550 |
| 10000 | 10100 | q | 44.795 | 44.504 |
| 2 | 0.5 | inverse | 57.736 | 58.913 |
| 100 | 0.9 | inverse | 340.535 | 340.234 |

## Where the time goes

Percentages below are within the *coarse instrumented pass*, not differences
between the initial and final baselines. Inclusive parent and child times must
not be added together.

- **Shape 100, both tails and both x points:** `EG_LogGamma` consumes about
  **89.7–91.2%** of the coarse wall time. Its recurrence executes 100 `EG_LogPair`
  calls to construct LogGamma at each P/Q evaluation (the lower path adds one
  further log of the series total). This is the primary bottleneck here.
- **Inverse at shape 100, p=.9:** three DD pair evaluations, three LogGamma
  evaluations and **300** recurrence logarithms. LogGamma accounts for **89.4%**
  of the coarse wall time. Uninstrumented median is about **340 ms** in both
  initial and restored passes.
- **Shape 10000, x=9900:** `WG_Pair` exclusive time accounts for about **95.5–95.7%**
  of the coarse wall time. The lower-series recurrence and its geometric
  remainder checks lie here; prefactor/deviance/exponential scopes are small.
- **Shape 10000, x=10100:** `WG_Upper` inclusive time accounts for about **94.7–94.8%**
  of coarse wall time (about 92.7–92.8% exclusive of its exponential). The upper
  continued-fraction work dominates.
- **Small shapes:** no single universal hotspot. At shape .0001 the lower-pair
  exclusive work, logarithm and LogGamma polynomial are the leading components;
  at shape 2.5, x=6, the upper iterations dominate.

The arithmetic pass exposes the scale of internal work. Per forward evaluation:

| Shape / x | SG arithmetic operations | DD validity checks | PV2 sum calls |
|:---|---:|---:|---:|
| 100 / 90 | 8,418 | 25,159 | 68,296 |
| 100 / 110 | 8,665 | 25,893 | 69,226 |
| 10000 / 9900 | 10,123 | 30,379 | 81,165 |
| 10000 / 10100 | 3,938 | 11,817 | 27,047 |

P and Q have identical counts at each of these points. The fine profiler itself
is expensive: an empty entry/exit pair measured **2.852 microseconds**, and fine
instrumentation makes the measured calls several times slower. Counts are useful;
fine exclusive timings include probe overhead and do **not** prove that removing
validation or one particular primitive would yield a proportional speedup.

## Recommended next experiment

Start with **reuse of the exact computed DD LogGamma value for repeated identical
shape inputs**, especially within an inverse solve. Preserve the first
computation's operation order and stored H/L/E values; compare cached and
uncached paths bit for bit, including alternating inputs and failure paths.
No such optimization has been implemented or measured here, and no speedup is
promised. This target addresses shape 100 but not the shape-10000 bottleneck.

For shape 10000, inspect repeated loop-invariant work and Variant/array dispatch
within the recurrence while preserving every numerical operation and rounding
store. Do not loosen convergence or remove validation merely because its call
count is large. Each actual optimization needs fresh regression and cost runs.

The failed first profile run 43015659345 stopped at Python setup (the pinned
version lacked a Windows build), before measurement. The successful workflow
therefore generates instrumentation on Linux and measures only VBA on Windows.
