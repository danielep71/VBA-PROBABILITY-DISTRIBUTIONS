# #348: where the time of a public Gamma call goes

A measurement, not a change. After #314 a public Gamma, Chi-square or
Poisson probability takes 0.06 to 0.23 ms; the binary64 series and continued
fraction of before #242, timed alone at the same points, take 0.004 to
0.077 ms. This study times the layers in between, and counts what the Gamma
inverse solver evaluates. No production source changes.

## Design, written before the run

`build_grid.py` writes `gamma_call_profile.csv`; Excel fills its last three
columns with `M_STATS_PROBDIST_GCP.bas` (study export `gcp_profile`);
`analyze.py` prints the tables below from the file.

**The layers of one call**, each timed alone at the eight points of the cost
benchmark of #298:

| Part | What is called |
| --- | --- |
| `public_cdf`, `public_sf` | `K_STATS_Gamma_Cumulative`, `K_STATS_Gamma_Survival`, scale 1 |
| `routed_p`, `routed_q` | `PROB_TryGammaRoutedP`, `PROB_TryGammaRoutedQ`: the route test and the routine it selects |
| `b64_kernel` | The binary64 routine the route selects: the series below `X = A + 1`, the continued fraction from there, each with its prefactor |
| `legacy_loop` | The series or continued fraction of before #242, with its own prefactor. The series stops at machine precision; the continued fraction stops at 3E-14 for shapes above one |
| `kernel_p`, `kernel_q` | The double-double entry points, for scale |
| `log_pdf` | `PROB_TryGammaLogPdf`, which the inverse solver calls once per iteration |
| `stirling` | `PROB_StirlingError`, for shapes above one |

A layer's own cost is its time minus the time of the layer below.

**Private routines.** The prefactor `PROB_B64_Factor`, the small-shape
log-gamma and the argument validation are `Private`. The module reaches them
through `Application.Run`, times that call on a trivial private routine
(`run_pow2`) and the analysis subtracts it. If Excel refuses such a call the
row says `ok=0` and the table shows a dash; the difference between layers
remains.

**Every timing** goes through one dispatcher and repeats its call until
0.5 s have elapsed. The part `empty` is the dispatcher alone and is
subtracted.

**The solver.** `ReplicaInverse` is `PROB_TryGammaInvP` written again from
the public routines it calls, storing every iterate. `test_profile.py`
compares its statements with the production routine's, and the export
records whether it returned the same Double; a row where it did not is
refused by `analyze.py`. For each iterate the module records which routine
served the tail (binary64 series, binary64 fraction, or the double-double
kernel, by the tests of the routed entry points) and times that evaluation
and the density evaluation, each for 0.1 s.

## What this cannot show

- The deviance inside the prefactor: it returns through `ByRef` arguments,
  which `Application.Run` does not carry. It is in the prefactor's time.
- Chi-square and Poisson: they share the routed entry points and differ in
  their wrappers, which are not timed here.
- Other hosts. One run, on the self-hosted runner.
- A subtraction of two timings carries the noise of both. Differences below
  about a microsecond are not read.

## Result, 2026-11-16

Excel run 43938463084 on `f85b4bc`, evidence `d5a9c63`; Excel 16.0 build
20710, 64-bit. Every row was exported, the three private routines were
reached, and the solver replica returned the Double of `PROB_TryGammaInvP`
on all ten inverse rows. The tables are the output of `analyze.py`.

### Layers of a public call, microseconds

| Shape | X | Public CDF | Public SF | Routed P | Routed Q | Binary64 kernel | Loop before #242 | Double-double P | Double-double Q |
| ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: |
| 0.0001 | 0.5 | 6.34 | 4943.75 | 3.80 | 5004.78 | 2.91 | 5.03 | 5065.82 | 4943.75 |
| 0.5 | 0.25 | 6.16 | 6.28 | 3.63 | 4.19 | 2.61 | 1.36 | 6469.63 | 6469.63 |
| 2.5 | 1 | 10.99 | 10.87 | 8.01 | 8.01 | 7.05 | 1.48 | 7079.98 | 6835.84 |
| 2.5 | 6 | 10.87 | 11.35 | 8.25 | 8.01 | 7.17 | 2.20 | 14648.34 | 14648.34 |
| 100 | 90 | 15.40 | 15.88 | 12.66 | 12.78 | 11.82 | 4.55 | 11352.44 | 12817.28 |
| 100 | 110 | 12.66 | 12.30 | 10.63 | 9.56 | 8.84 | 4.01 | 13549.71 | 13793.85 |
| 10000 | 9900 | 64.75 | 66.66 | 62.84 | 61.89 | 62.84 | 29.23 | 150390.53 | 160156.15 |
| 10000 | 10100 | 20.41 | 19.93 | 17.07 | 17.54 | 16.35 | 10.15 | 50781.15 | 53710.84 |

### What each layer adds, microseconds

| Shape | X | Wrapper, CDF | Wrapper, SF | Route test | Kernel over the old loop | Prefactor alone | Validation alone | Stirling or small log-gamma |
| ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: |
| 0.0001 | 0.5 | 2.53 | -61.03 | 0.89 | -2.12 | 0.83 | 3.34 | -0.24 |
| 0.5 | 0.25 | 2.53 | 2.09 | 1.01 | 1.25 | 1.19 | 3.34 | 0.24 |
| 2.5 | 1 | 2.98 | 2.86 | 0.95 | 5.57 | 7.63 | 3.34 | 0.14 |
| 2.5 | 6 | 2.62 | 3.34 | 0.83 | 4.98 | 5.48 | 3.10 | 0.14 |
| 100 | 90 | 2.74 | 3.10 | 0.83 | 7.27 | 5.72 | 3.81 | 0.14 |
| 100 | 110 | 2.03 | 2.74 | 0.71 | 4.83 | 5.96 | 3.58 | 0.14 |
| 10000 | 9900 | 1.91 | 4.77 | 0.00 | 33.62 | 5.96 | 2.38 | 0.12 |
| 10000 | 10100 | 3.34 | 2.38 | 1.19 | 6.20 | 4.77 | 3.10 | 0.12 |

### Parts timed once, microseconds

| Part | Time | Loop cost removed |
| --- | ---: | ---: |
| `empty` | 0.099 | - |
| `set_status` | 0.207 | 0.108 |
| `try_divide` | 0.462 | 0.363 |
| `classify` | 0.268 | 0.169 |
| `is_finite` | 0.222 | 0.123 |
| `run_pow2` | 12.875 | 12.776 |

### The inverse solver

| Shape | p | Public inverse, us | Solver alone, us | Iterations | Series | Fraction | Double-double | Tail evaluations, us | Density evaluations, us | Rest, us |
| ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: |
| 0.01 | 0.1 | 9.0 | 46.6 | 7 | 7 | 0 | 0 | 20.4 | 26.2 | 0.0 |
| 0.01 | 0.9 | 29540.9 | 29785.1 | 20 | 11 | 0 | 9 | 31284.7 | 77.0 | -1576.6 |
| 0.5 | 0.1 | 18.5 | 14.4 | 4 | 4 | 0 | 0 | 13.4 | 1.9 | -0.8 |
| 0.5 | 0.9 | 43.3 | 39.5 | 9 | 9 | 0 | 0 | 37.2 | 4.7 | -2.5 |
| 2.5 | 0.1 | 33.8 | 28.8 | 3 | 3 | 0 | 0 | 23.8 | 1.6 | 3.4 |
| 2.5 | 0.9 | 43.8 | 39.0 | 4 | 0 | 4 | 0 | 33.9 | 2.2 | 2.9 |
| 100 | 0.1 | 328.0 | 328.0 | 23 | 23 | 0 | 0 | 292.8 | 23.1 | 12.1 |
| 100 | 0.9 | 77.1 | 67.6 | 6 | 0 | 6 | 0 | 55.8 | 6.0 | 5.9 |
| 10000 | 0.1 | 5920.3 | 6530.7 | 52 | 51 | 0 | 1 | 6105.4 | 40.9 | 384.3 |
| 10000 | 0.9 | 5798.2 | 5798.2 | 53 | 0 | 51 | 2 | 5511.8 | 40.5 | 246.0 |

### What the numbers say

**The premise of the issue does not hold on this run.** A public Gamma
cumulative takes 6 to 65 microseconds here, not 60 to 230. `cost_lines.txt`
holds the cost benchmark of the same run, which agrees: 6.6 to 78 for the
Gamma cumulative, 6.3 to 93 for Chi-square, 14 to 69 for Poisson. Two other
runs of the same day, 43916635364 and 43920764675, give the same figures
within a few percent. The table of
`../gamma_route_study/HOLDOUT_ROUTE_RESULT.md`, run 43573743649, is about
ten times higher for the public functions and about 2.6 times higher for
the old loop and for the kernel. I do not know why: the benchmark code is
the same, that record itself notes a second run that moved figures by a
factor of two, and other source changes lie between the two runs. I did not
repeat the old run, so the old figures are superseded, not explained.

**Where a call's time goes.** Reading down one column of the second table:

- the public wrapper, about 2 to 3 microseconds: validation is most of it,
  about 3 through `Application.Run`, a figure that carries that
  subtraction's noise; the status calls, the division and the classification
  are 0.1 to 0.4 each;
- the route test, about 1;
- the binary64 kernel over the loop of before #242, 1 to 7, and 34 at
  shape 10,000 below the seam; at shape 0.0001 the kernel is the faster
  of the two. For shapes above one the prefactor alone is
  5 to 8, so it accounts for that difference everywhere but there. At
  (10,000, 9,900) both routines are series stopped at machine precision,
  and the binary64 one takes 28 microseconds more than its prefactor
  explains. I have not measured why: it carries a compensated sum and a
  second test on the remainder, and I did not count its terms;
- the old loop, 1.4 to 29.

Nothing between the wrapper and the loop is large. At small shapes the
wrapper and the kernel are comparable; at shape 10,000 the series is the
call.

Review of the pull request corrected this section and the table of parts
above. I had written that the old loop stopped at 3E-14 everywhere and had
put the difference at (10,000, 9,900) down to the stop. Only the old
continued fraction stops there; the old series does not.

**The solver.** The tail evaluations are 83 to 95 percent of an inverse and
the density evaluations most of the rest. What is unexpected is the count:

- 52 and 53 iterations at shape 10,000, and 23 at shape 100 and p = 0.1,
  where shapes 0.5 and 2.5 take 3 to 9;
- at shape 0.01 and p = 0.9, 20 iterations of which nine fall to the
  double-double kernel, and those nine are the whole 30 ms.

At shape 0.01 and p = 0.1 the public inverse takes 9 microseconds and does
not reach the solver: the quantile is formed in log space (#14). The solver
columns of that row time a routine the public call does not use.

Differences in the last column are within the noise of the two sums at
shape 10,000 and at shape 0.01, p = 0.9, where a single evaluation is timed
a few times only.

### Decision

- **The public call stays as it is.** No part of it is worth an Excel wave:
  outside the series the largest single cost is a prefactor of 5 to 8
  microseconds. The one unexplained figure, the series at shape 10,000
  below the seam, is 28 microseconds in a call of 65.
- **The solver's iteration count is worth a study**, and is issue 415. If
  it came down to the six of shape 100 at p = 0.9, an inverse at shape
  10,000 would take about a tenth of what it takes. The cause is not
  diagnosed here.
- **The published cost figures are stale.** The root README is corrected
  with this study. The v1.0.0 manual carries the old figures and is not
  edited; the v1.0.1 manual takes the new ones (#264, recorded in #400).
