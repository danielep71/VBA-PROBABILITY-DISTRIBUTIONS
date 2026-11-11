# Exact DD LogGamma cache: Excel result (#307)

## Decision and implementation

The one-entry cache is implemented in production SPECIALFUNCS. It reuses the
exact H/L/E result of the internal `PROB_DD_EG_LogGamma` for an identical input
tuple; misses execute the unchanged original arithmetic. Hits return a fresh
array, and failed misses do not overwrite the last successful entry. Public
routing, convergence criteria, error handling in the public adapters and the
untouched #307 routing holdout remain unchanged.

**This materially accelerates repeated shape-100 calls and inverse solves. It
does not resolve the shape-10000 bottleneck or complete all of #307.** A cold
forward call still computes LogGamma; no universal speedup is claimed.

## Evidence

- Candidate `d6d952a1c67f64c1481f25d8d7ee4feea88538fd`.
- [Confirmatory cache study, run 43021836369](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/actions/runs/43021836369): PASS,
  Excel 16.0 build 20430 on the Windows x64 Excel runner.
- [Authorized regression/export workflow, run 43021836536](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/actions/runs/43021836536):
  **5987/5987 regression assertions, 8901/8901 private study assertions**, Office
  64-bit; fresh exports recomputed **1783 in-scope main rows and all 659
  historical-holdout rows**. The other 356 main-grid rows were out of scope
  and retained unchanged by the established exporter. All **2139 + 659 = 2798
  CSV observations are textually identical** to the preceding committed grid. The normal evidence importer refreshed the record and both
  manifests and passed the strict accuracy gate and retained-record checks.
- [Initial cache samples](results/run_43019849422.json), candidate
  `1066e45bfe34cbed441ac1103d9dfe1be75fbb2a`: the benchmark passed; its workflow
  later failed before regression execution because a repository-path argument
  was missing. It is supporting benchmark evidence, not a green workflow or
  regression certification. The cache/source and generated measurement files
  are identical to the confirmatory run; later fixes affect orchestration only.
- [Confirmatory raw samples](results/run_43021836369.json) retain the source
  hashes, generated-file hashes, artifact identity and digest. ZIP SHA-256 was
  checked against GitHub's artifact digest. The [export manifest](results/export_manifest_43021836536.json)
  and [raw regression log encoded as JSON](results/regression_log_43021836536.json)
  retain the export trace and all study counters. Encoding the log preserves
  its original BOM/CRLF bytes; `raw_utf8.encode('utf-8')` reconstructs the bytes
  whose SHA-256 the certification record binds.

In each of two runs: **216 samples / 1008 target evaluations** agreed bit for bit,
plus 40 private tuple/error/mutation checks per path. Across runs: **432 samples,
2016 target evaluations and 160 private checks**. The 18 cases are the previously
revealed profiling cases, never the untouched routing holdout.

## Performance interpretation

Three rounds per case alternate uncached/cached execution order; the second
round reverses it. Cold trials evict the cache using shape .75 outside timing;
warm trials prime the target shape outside timing. The inverse's cold trial
therefore starts with a miss but can reuse the result during its own iterations.

- **Shape 100, repeated P/Q:** across both runs, uncached medians are about
  **115–121 ms**, cached medians **9.9–13.5 ms**, or **8.5–12 times faster** across
  the four point/tail combinations.
- **Inverse, shape 100, p=.9:** cold medians **545–563 ms → 225–252 ms**
  (**2.2–2.5x**); warm **346–561 ms → 35–38 ms** (**9–16x** observed).
- **Shape 10000:** the changed LogGamma routine is not called on these paths.
  There is no algorithmic benefit here. The second run's lower-Q warm control
  even shows a spurious 1.94x apparent improvement, demonstrating substantial
  host timing variability. It is not attributed to the cache.
- **Cold forward calls:** no established benefit. In the confirmatory shape-100
  run, cached medians are 6–9% slower; the first run is mixed. With only three
  samples and noisy controls, a small cold-path regression is not excluded and
  no precise overhead bound is claimed. Small-shape timings are similarly
  mixed; this is a targeted reuse optimization, not a universal acceleration.

The table gives confirmatory medians, not guaranteed latency bounds or a
statistical confidence interval. All samples, including unfavorable ones, are
retained.

| Shape | x/probability | Function | Cache | Original ms | Cached ms | Ratio |
|---:|---:|:---|:---|---:|---:|---:|
| 0.0001 | 0.5 | p | cold | 7.086 | 7.667 | 0.92x |
| 0.0001 | 0.5 | p | warm | 5.093 | 3.830 | 1.33x |
| 0.0001 | 0.5 | q | cold | 6.970 | 8.652 | 0.81x |
| 0.0001 | 0.5 | q | warm | 5.191 | 3.877 | 1.34x |
| 0.5 | 0.25 | p | cold | 11.491 | 8.222 | 1.40x |
| 0.5 | 0.25 | p | warm | 5.971 | 5.476 | 1.09x |
| 0.5 | 0.25 | q | cold | 9.170 | 9.438 | 0.97x |
| 0.5 | 0.25 | q | warm | 5.497 | 6.276 | 0.88x |
| 2.5 | 1 | p | cold | 14.385 | 12.255 | 1.17x |
| 2.5 | 1 | p | warm | 8.304 | 8.154 | 1.02x |
| 2.5 | 1 | q | cold | 13.265 | 12.240 | 1.08x |
| 2.5 | 1 | q | warm | 9.291 | 6.196 | 1.50x |
| 2.5 | 6 | p | cold | 21.964 | 25.469 | 0.86x |
| 2.5 | 6 | p | warm | 16.650 | 12.655 | 1.32x |
| 2.5 | 6 | q | cold | 21.862 | 24.430 | 0.89x |
| 2.5 | 6 | q | warm | 24.197 | 13.185 | 1.84x |
| 100 | 90 | p | cold | 173.030 | 185.236 | 0.93x |
| 100 | 90 | p | warm | 114.949 | 13.543 | 8.49x |
| 100 | 90 | q | cold | 180.842 | 197.456 | 0.92x |
| 100 | 90 | q | warm | 117.912 | 9.859 | 11.96x |
| 100 | 110 | p | cold | 180.296 | 192.736 | 0.94x |
| 100 | 110 | p | warm | 116.709 | 11.755 | 9.93x |
| 100 | 110 | q | cold | 189.809 | 202.486 | 0.94x |
| 100 | 110 | q | warm | 117.762 | 11.866 | 9.92x |
| 10000 | 9900 | p | cold | 214.645 | 205.604 | 1.04x |
| 10000 | 9900 | p | warm | 129.774 | 131.907 | 0.98x |
| 10000 | 9900 | q | cold | 212.391 | 202.063 | 1.05x |
| 10000 | 9900 | q | warm | 266.614 | 137.564 | 1.94x |
| 10000 | 10100 | p | cold | 76.902 | 74.581 | 1.03x |
| 10000 | 10100 | p | warm | 46.256 | 44.534 | 1.04x |
| 10000 | 10100 | q | cold | 72.130 | 75.110 | 0.96x |
| 10000 | 10100 | q | warm | 43.898 | 44.517 | 0.99x |
| 2 | 0.5 | inverse | cold | 87.257 | 63.514 | 1.37x |
| 2 | 0.5 | inverse | warm | 56.673 | 31.972 | 1.77x |
| 100 | 0.9 | inverse | cold | 562.692 | 225.202 | 2.50x |
| 100 | 0.9 | inverse | warm | 561.418 | 34.918 | 16.08x |

## Reproduce and next scope

```
python benchmark/gamma_cache_study/summarize.py benchmark/gamma_cache_study/results/run_43021836369.json
python -m unittest discover -s benchmark/gamma_cache_study -p test_cache_vba.py
```

The next distinct target is repeated loop-invariant work in the large-shape DD
recurrences, preserving exact operations and Double stores. It is not included
in this change. The original rejected routing studies and their frozen decisions
remain preserved; no relaxation or retrospective holdout selection was made.
