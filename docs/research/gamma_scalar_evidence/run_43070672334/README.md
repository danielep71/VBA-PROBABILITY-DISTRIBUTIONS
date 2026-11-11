# B5 scalar-pair Excel evidence — 2026-11-11

Source: `e498658f3e94ef259187ce1d4da8a34992357c2d`.
[Run 43070672334](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/actions/runs/43070672334) passed all three jobs. Excel 16.0 build 20430, Windows/Office 64-bit runner.

All 576 operator hi/lo comparisons, 11 normalization checks and 1,371 exact-input checks passed (1,265 bit-identical P/Q pairs and 106 expected declines). The complete 48-row contemporaneous timing matrix passed, with no cleanup error.

| A | X | B4 pair ms | B5 scalar pair ms | Speedup | Cost reduction |
|---:|---:|---:|---:|---:|---:|
| 0.0001 | 0.5 | 6.946 | 4.737 | 1.47x | 31.8% |
| 0.5 | 0.25 | 8.470 | 6.645 | 1.27x | 21.5% |
| 2.5 | 1 | 9.708 | 7.357 | 1.32x | 24.2% |
| 2.5 | 6 | 10.642 | 3.343 | 3.18x | 68.6% |
| 100 | 90 | 21.078 | 7.826 | 2.69x | 62.9% |
| 100 | 110 | 10.251 | 3.471 | 2.95x | 66.1% |
| 10000 | 9900 | 164.150 | 28.137 | 5.83x | 82.9% |
| 10000 | 10100 | 37.170 | 8.212 | 4.53x | 77.9% |

These are medians of three warm paired P/Q samples, each of four individually timed calls. B4/B5 run order alternates by round. This is a within-run comparison, not a comparison with the earlier B4 machine/session.

## Decision

The frozen development signal passes: both A=10,000 points improve by more than 5%, and none of the eight points regresses. Removing recurrence Variant arrays is useful. However, measured B5 cost remains 3.343–28.137 ms per pair, so the conservative 100-us paired budget fails at every point. This is not a public-path, cold-cache or inverse measurement, nor a production routing approval. Independent timing confirmation and public-path/boundary validation remain required before promotion. #307 stays open and its untouched holdout remains untouched.

## Retention and replay

The artifact ZIP SHA-256 and original host-report digest are in [provenance.json](provenance.json). Only the host JSON formatting was normalized to UTF-8/LF; its values were preserved. The complete design, instrumentation, host results and summary are retained alongside this file.

Use the validator from the measured source commit:

```sh
python benchmark/gamma_route_study/validate_scalar_host.py docs/research/gamma_scalar_evidence/run_43070672334
```

The separate ordinary Excel regression run 43070683108 initially failed with RPC 0x800706BE and cleanup RPC 0x800706BA; a retry was requested. Its outcome is separate from this successfully completed candidate study.
