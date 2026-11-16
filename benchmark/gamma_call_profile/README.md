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
| `legacy_loop` | The series or continued fraction of before #242, stopped at 3E-14, with its own prefactor |
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
