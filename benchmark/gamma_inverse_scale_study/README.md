# #14: the Gamma inverse at the final scale

**Status: the first candidate was rejected on its seam rule; a second is registered and its holdout is not yet evaluated.**

`K_STATS_Gamma_InverseCumulative` and `K_STATS_ChiSquare_InverseCumulative`
lose a quantile that is representable at the final scale when its unit-scale
value is below the Double range. The rule registered here answers small unit
quantiles in log space and forms the result once, at the final scale.

| File | Role |
| --- | --- |
| [`PREREGISTRATION.md`](PREREGISTRATION.md) | The rule, its crossover, the acceptance rules G1 to G7, the contracts and the predictions, frozen before the holdout is evaluated |
| [`RESULT_1.md`](RESULT_1.md) | The first verdict: rejected on G7, with the predictions against the outcome |
| [`PREREGISTRATION_2.md`](PREREGISTRATION_2.md) | The second candidate: another source of `LogGamma(1 + A)` up to shape 2, a wider seam, a new holdout |
| `inv_scale_model.py` | The log branch as a model over the pair routines of `gamma_route_study/route_model.py` |
| `holdout_build.py`, `gamma_inverse_scale_grid.csv`, `gamma_inverse_scale_grid2.csv` | The 455 rows of the first holdout and the 525 of the second, with their references; `--check` proves each grid is its design |
| `M_STATS_PROBDIST_GIS.bas` | `Export_GammaInverseScaleHoldout`: the public inverses, the round trip and the classifying value at every row, as bit patterns. The runner exports it as the study grid `gis_holdout` |
| `holdout_analyze.py` | The verdict against G1 to G7 |
| `test_holdout.py` | Fixtures for the model, the construction and each rule, on another seed and other seam shapes |
