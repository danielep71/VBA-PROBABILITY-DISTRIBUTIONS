# #24 lower-region Gamma Q holdout

#24: below the branch seam `X = A + 1`, `PROB_TryGammaRegularizedQ` used to
form survival as `Q = 1 - P`. That turned the absolute error of P into the
absolute error of Q, and multiplied the relative error by `P / Q`, up to about
1E7 at the smallest shapes. Since #242 the lower-region Q comes from log P
before rounding. This directory holds the untouched holdout that tests that
repair in Excel.

| File | Role |
| --- | --- |
| [`HOLDOUT_DESIGN.md`](HOLDOUT_DESIGN.md) | The frozen design: shapes, arguments, routes, references and rules H1 to H4 |
| `design.py` | The design's inputs, read by every script here |
| `check_disjoint.py` | One-time proof, run before the evaluation, that no shape or doubled shape occurs in any tracked file. It is not in CI: later documents may cite the shapes |
| `generate_gamma_lower_q.py` | Builds `gamma_lower_q_grid.csv` from the design, with mpmath references; `--check` proves the committed grid is the design |
| `M_STATS_PROBDIST_GLQ.bas` | `Export_GammaLowerQ`: fills the echoes and the observed results in Excel. The runner exports it as the study grid `glq_lower_q` |
| `analyze_gamma_lower_q.py` | Scores the holdout (H1 to H4) and the re-exported fitting ladder of `positive_ratio_subnormal/gamma_q_complement.csv` (L1, L2) |
| `test_gamma_lower_q.py` | Proves the grid is the design and that the analyzer catches each failure; it runs in the evidence-tool suite |

Run the analyzer after an export:

```sh
python3 benchmark/gamma_lower_q/analyze_gamma_lower_q.py
```
