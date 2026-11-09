# #24 predictions before the Excel wave (2026-11-09)

Written and committed before the Excel wave that tests them. Nothing below may
be revised after that wave's output is observed.

## The candidate

`release/1.0.0` at `279f15c`, plus three changes, none of them arithmetic:

1. **The holdout tooling.** The untouched holdout of `HOLDOUT_DESIGN.md`, its
   exporter `M_STATS_PROBDIST_GLQ.bas` and its analyzer.
2. **The fallback.** The headers of `PROB_TryGammaRegularizedP` and `Q` now
   document that the binary64 fallback at the pair's iteration cap forms the
   complement. `Test_PublicGamma_Reach` gains one assertion per point: the
   complemented tail's relative error may not exceed 1.001 times the other
   tail's plus 1E-15.
3. **Main-grid coverage of the repaired path.** `promote_x05_ladder.py` adds
   22 rows: `Gamma_Survival(0.5, A, 1)` and `ChiSquare_Survival(1, 2A)` at
   the eleven `x05` shapes from 1E-6 to 0.1, claim `rel<=5E-15`.

## Model values for the fitting ladder

`wide_gamma.public`, the model that reproduces the production kernel bit for
bit (`test_public_fresh.py`), gives these public values at the 34 points of
`gamma_q_complement.csv`, already revealed. The holdout is not evaluated by
the model, by mpmath against the candidate, or in any other way before the
wave.

| Point | Shape | X | CDF (binary64) | Survival (binary64) | CDF error | Survival error |
| --- | ---: | --- | --- | --- | ---: | ---: |
| `x05_s0` | 1e-06 | 0.5 | `0x1.ffffed379433cp-1` | `0x1.2c86bcc43809cp-21` | 2.4e-17 | -1.4e-18 |
| `x05_s1` | 3e-06 | 0.5 | `0x1.ffffc7a6b8bc4p-1` | `0x1.c2ca3a1de27b2p-20` | -2.6e-17 | 4.9e-17 |
| `x05_s2` | 1e-05 | 0.5 | `0x1.ffff442b8ff55p-1` | `0x1.77a8e015625e1p-18` | 8.2e-18 | 1.6e-17 |
| `x05_s3` | 3e-05 | 0.5 | `0x1.fffdcc812ccd5p-1` | `0x1.19bf699954c33p-16` | -4.5e-17 | 2.3e-17 |
| `x05_s4` | 0.0001 | 0.5 | `0x1.fff8a99cf1ef2p-1` | `0x1.d598c384385d5p-15` | 1.0e-17 | -2.2e-17 |
| `x05_s5` | 0.0003 | 0.5 | `0x1.ffe9fc3fb1a71p-1` | `0x1.603c04e58f0a3p-13` | 4.4e-18 | -1.1e-17 |
| `x05_s6` | 0.001 | 0.5 | `0x1.ffb697480dc2fp-1` | `0x1.25a2dfc8f4278p-11` | -4.2e-17 | -7.8e-17 |
| `x05_s7` | 0.003 | 0.5 | `0x1.ff238b090ab22p-1` | `0x1.b8e9edea9bb55p-10` | -3.7e-17 | -3.3e-17 |
| `x05_s8` | 0.01 | 0.5 | `0x1.fd1e7d6454252p-1` | `0x1.70c14dd5ed73ep-8` | 5.4e-17 | 5.3e-17 |
| `x05_s9` | 0.03 | 0.5 | `0x1.f74564c4fa299p-1` | `0x1.17536760bace6p-6` | 2.2e-17 | -2.5e-17 |
| `x05_s10` | 0.1 | 0.5 | `0x1.e1ff805af1efep-1` | `0x1.e007fa50e1025p-5` | 3.8e-17 | -2.1e-17 |
| `x05_s11` | 0.5 | 0.5 | `0x1.5d897a241a6fap-1` | `0x1.44ed0bb7cb20bp-2` | -6.7e-17 | -3.1e-17 |
| `x2_s0` | 1e-06 | 2 | `0x1.fffffe5bf278cp-1` | `0x1.a40d873a1dc08p-25` | -4.1e-17 | 8.4e-18 |
| `x2_s1` | 3e-06 | 2 | `0x1.fffffb13d6685p-1` | `0x1.3b0a65ec3a85ap-23` | 6.3e-18 | 2.2e-17 |
| `x2_s2` | 1e-05 | 2 | `0x1.ffffef9769999p-1` | `0x1.06896666d8be2p-21` | -1.7e-17 | -4.3e-17 |
| `x2_s3` | 3e-05 | 2 | `0x1.ffffcec5d8037p-1` | `0x1.89d13fe4524d5p-20` | -4.0e-17 | 3.7e-17 |
| `x2_s4` | 0.0001 | 2 | `0x1.ffff5be4382d7p-1` | `0x1.48378fa514299p-18` | -4.1e-17 | 6.8e-17 |
| `x2_s5` | 0.0003 | 2 | `0x1.fffe138548ec9p-1` | `0x1.ec7ab7136affcp-17` | -3.5e-17 | 5.3e-17 |
| `x2_s6` | 0.001 | 2 | `0x1.fff9949b7ab45p-1` | `0x1.9ad92152eae8ap-15` | -3.0e-17 | -3.3e-17 |
| `x2_s7` | 0.003 | 2 | `0x1.ffecae6d37f25p-1` | `0x1.35192c80daf91p-13` | -3.0e-18 | -3.2e-17 |
| `x2_s8` | 0.01 | 2 | `0x1.ffbee6b41c72dp-1` | `0x1.04652f8e34bfap-11` | -6.3e-19 | -3.5e-17 |
| `x2_s9` | 0.03 | 2 | `0x1.ff36a1aca86ccp-1` | `0x1.92bca6af2675bp-10` | -3.6e-17 | 3.5e-17 |
| `x2_s10` | 0.1 | 2 | `0x1.fd18520f207c6p-1` | `0x1.73d6f86fc1cfbp-8` | -4.0e-18 | -7.2e-17 |
| `x2_s11` | 0.5 | 2 | `0x1.e8b4307d3627ap-1` | `0x1.74bcf82c9d860p-5` | -2.9e-18 | 6.1e-17 |
| `xsub_s0` | 1e-06 | 0x0.0000000800000p-1022 | `0x1.ffa09fedf8f7fp-1` | `0x1.7d80481c20374p-11` | -1.5e-17 | 6.7e-17 |
| `xsub_s1` | 3e-06 | 0x0.0000000800000p-1022 | `0x1.fee215133afc1p-1` | `0x1.1deaecc503f7bp-9` | 5.3e-17 | 5.5e-17 |
| `xsub_s2` | 1e-05 | 0x0.0000000800000p-1022 | `0x1.fc495d3c6ca43p-1` | `0x1.db5161c9ade45p-8` | -5.1e-17 | -1.7e-17 |
| `xsub_s3` | 3e-05 | 0x0.0000000800000p-1022 | `0x1.f4f0ba0e8e45dp-1` | `0x1.61e8be2e37470p-6` | 5.6e-17 | 3.5e-17 |
| `xsub_s4` | 0.0001 | 0x0.0000000800000p-1022 | `0x1.dc0dfac522c96p-1` | `0x1.1f9029d6e9b4dp-4` | -4.8e-17 | 4.1e-17 |
| `xsub_s5` | 0.0003 | 0x0.0000000800000p-1022 | `0x1.9b8eaf5440f6bp-1` | `0x1.91c542aefc253p-3` | -5.1e-17 | 6.8e-17 |
| `xsub_s6` | 0.001 | 0x0.0000000800000p-1022 | `0x1.ee806cef103f3p-2` | `0x1.08bfc98877e06p-1` | -2.0e-17 | -8.9e-17 |
| `xsub_s7` | 0.003 | 0x0.0000000800000p-1022 | `0x1.cd46c01c38fe6p-4` | `0x1.c65727fc78e03p-1` | 3.3e-17 | -3.5e-17 |
| `xsub_s8` | 0.01 | 0x0.0000000800000p-1022 | `0x1.6996f01df489bp-11` | `0x1.ffa59a43f882ep-1` | 2.0e-17 | 1.7e-17 |
| `xsub_s9` | 0.03 | 0x0.0000000800000p-1022 | `0x1.6884c12563fb1p-32` | `0x1.fffffffd2ef68p-1` | -5.3e-17 | 1.6e-17 |

Worst model error: P 6.7E-17, Q 8.9E-17. In the committed pre-#24 export the
survival errors reach 1.6E-9 at `x05_s0`.

## Predictions

| Id | Prediction |
| --- | --- |
| N1 | The regression suite passes all **14,888** assertions: the 14,886 of run 42679149964 plus the two fallback complement assertions. |
| N2 | Re-exported, the fitting ladder returns the model values above for every public CDF and survival, so L1 holds (CDF within 3E-15, survival within 5E-15). The raw binary64 series P equals the `e278318` export bit for bit (L2). |
| N3 | The untouched holdout passes H1 to H4 of `HOLDOUT_DESIGN.md`. |
| N4 | The 22 promoted rows equal the model values of their `x05` points, within 7.8E-17. |
| N5 | No other main-grid observation changes, and no holdout observation changes. |
| N6 | 166/166 main and 83/83 holdout contracts pass. The 27-row coverage debt is unchanged. |

Any failure of N1 to N6 stops the closure of #24, and the cause is explained
before anything else changes. A failure of N3 rejects the candidate for #24, as
the design states.
