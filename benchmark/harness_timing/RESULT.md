# #298 timing of the VBA suite and cost of the Gamma kernel (2026-11-10)

Excel VBA Regression run 42766552574, candidate `cdddfbe`, Excel 16.0 build
20430, 64-bit, on the self-hosted runner: 14,888 / 14,888 assertions. The
evidence commit is `1901311`; the log is kept here as
[`run_42766552574_log.txt`](run_42766552574_log.txt). Stored with LF line ends, it
hashes with CRLF line ends, as the runner wrote it, to the SHA-256 the
certification record of that run binds,
`b0a77d84f8c7818d0e1f7939afbc5328477266159c8ec98c7fb7a30afb7f34dd`.
[`summarize_timing.py`](summarize_timing.py) reproduces every number below
from it.

## Where the suite spends its time

The 103 timed test procedures took 614 s.

| Suite driver | Seconds | Share |
| --- | ---: | ---: |
| `RunContinuousSuite` | 582.3 | 94.8% |
| `RunTFamilySuite` | 31.1 | 5.1% |
| `RunDiscreteSuite` | 0.5 | 0.1% |
| `RunCoreSuite` | 0.2 | 0.0% |
| `RunNormalFamilySuite` | 0.0 | 0.0% |

One procedure, `Test_CN_GammaSurvival`, takes 579.9 s, 94.4% of the total.
It calls the eleven #26 study batteries: `Test_GammaCF_Coverage`,
`Test_GammaCF_AdditionalCallers`, `Test_GammaCF_CapAndXSeam`,
`Test_GammaCF_PrecisionCounterexample`, and the fidelity blocks of Precision
v2, the scaled product, the scaled double-double, the scaled, extended and
wide Gamma ports, and the public integration. The next is
`Test_TF_ChiSquareLargeDF` at 27.8 s (4.5%). The other 101 procedures take
6.5 s together, and 94 of them less than 0.1 s each.

## Cost per call of the Gamma kernel

Microseconds per call, each timed over at least 0.5 s. "binary64" is the
routine that answered the same region before #242: the lower series below
`X = A + 1`, the continued fraction above it.

| A | X | Double-double P | Double-double Q | binary64 | Slower by |
| ---: | ---: | ---: | ---: | ---: | ---: |
| 0.0001 | 0.5 | 5,066 | 6,424 | 6.3 | 1,024x |
| 0.5 | 0.25 | 7,065 | 7,767 | 1.8 | 4,354x |
| 2.5 | 1 | 11,139 | 10,101 | 2.0 | 5,643x |
| 2.5 | 6 | 19,043 | 18,005 | 3.0 | 6,438x |
| 100 | 90 | 150,513 | 140,625 | 6.1 | 24,517x |
| 100 | 110 | 156,250 | 154,053 | 5.2 | 29,790x |
| 10,000 | 9,900 | 164,917 | 170,288 | 38.7 | 4,395x |
| 10,000 | 10,100 | 56,152 | 63,477 | 13.1 | 4,830x |

A double-double Gamma P or Q call takes from 5 ms to 170 ms; the binary64
routine took from 2 to 39 microseconds. A worksheet that fills 1,000 cells
with `Gamma_Cumulative` at shape 100 recalculates for about 2.5 minutes.

## What follows

- **#300.** The study batteries are the whole runtime of the suite. Moving
  them to a study export, and keeping regression cases in the suite, would
  cut it from about 10 minutes to well under one. *Corrected by the #300
  result below: this run timed `Test_CN_GammaSurvival` as one block, and the
  attribution to the study batteries was wrong.*
- **The kernel's cost.** This is a usability defect of #242's kernel, not of
  the harness. It is recorded as its own issue.

## #300 result: the study batteries split off (2026-11-10)

Excel VBA Regression run 42775390763, candidate `5c1d66d`, evidence commit
`2130def`, same runner and Excel build. The six private-port batteries now run
apart, in `RunStudyBatteries`. The certified log is kept as
[`run_42775390763_log.txt`](run_42775390763_log.txt). Stored with LF line
ends, it hashes with CRLF line ends to the SHA-256 its record binds,
`16ce9124b779fa507f54ac5823489196357e319974980c6e4bffab290b08b3a7`.

| Part | Assertions | Elapsed in the log |
| --- | ---: | ---: |
| Regression (`Test_STATS_PROBDIST_RunAll_CI`) | 5,987 / 5,987 | 7 min 35 s |
| Study batteries (`Test_STATS_PROBDIST_RunStudies_CI`) | 8,901 / 8,901 | 1 min 10 s |

| Procedure | Seconds |
| --- | ---: |
| `Test_CN_GammaSurvival` (now without the six batteries) | 419.0 |
| `Test_WideGamma_Fidelity` (study) | 51.9 |
| `Test_TF_ChiSquareLargeDF` | 28.7 |
| `Test_ExtendedGamma_Fidelity` (study) | 10.2 |
| every other procedure | under 4.3 each |

The prediction above was wrong. The six batteries took 70 s, not the bulk of
the 580 s. Most of the time is in what `Test_CN_GammaSurvival` still runs:
- the four `Test_GammaCF_*` batteries;
- `Test_PublicGamma_Integration`;
- its own assertions.

All of these call the production kernel at 5 ms to 170 ms per call (#307).
The study batteries' private ports are far cheaper.

The regression a pull request runs falls from about 10 minutes to about 7.5.
The rest of the suite's runtime is the kernel's cost, so #307 is what can
shorten it further.
