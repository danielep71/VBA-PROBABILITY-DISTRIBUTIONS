# Exact DD LogGamma reuse (#307)

The measured profile identified DD LogGamma as ~90% of the shape-100 cost.
The candidate keeps one successful result in `PROB_DD_EG_LogGamma`, keyed by
both IEEE-754 component bit patterns and the integer exponent. It returns a
fresh array made from scalar cached H/L/E components. Misses execute the exact
previous arithmetic in `PROB_DD_EG_LogGammaUncached`; failures do not publish a
cache entry. Memory use is fixed. VBA executes this module synchronously; no
cross-thread cache or external state is introduced.

`cache_vba.py` installs and strictly reverses the wrapper. Historical restoration
strips it before comparing the old frozen kernel; it does not recertify old
measurements. `prepare.py` additionally requires the stripped source hash to
match the uncached kernel measured in run 43017348872.

The Excel study checks complete DD tuples on repeated, alternating and adjacent
shapes, keys sharing a high component but differing in low bits/exponent, signed
zero, mutation of returned arrays, and a failing miss followed by a valid hit.
There are 40 private assertions for both uncached and cached paths.

Performance uses only the 18 already revealed profiling cases: eight points
for P/Q and two inverse cases. Three rounds alternate uncached/cached execution
order (the second round reverses it), and each records cold and warm calls.
Cold calls evict the one-entry cache with shape .75 outside the measured time;
warm calls prime the requested shape outside the measurement. Four calls per
forward sample and one per inverse sample give 216 samples. Every measured
value and warmup must retain the same binary64 bits. No routing boundary,
convergence tolerance or #307 holdout is changed or evaluated.

The workflow then runs the full production regression and all existing private
study batteries, without exporting or evaluating the untouched routing holdout.
The source change is a candidate until those results are collected. Historical
manifest bindings remain historical; no release/merge readiness is inferred.

Reproduce source guards:

```
python -m unittest discover -s benchmark/gamma_cache_study -p test_cache_vba.py
python benchmark/gamma_cache_study/prepare.py
```

Actual Excel measurement: `Gamma DD LogGamma Cache` on the study branch, or
`powershell -File benchmark/gamma_cache_study/Run-CacheStudy.ps1` after preparation.
