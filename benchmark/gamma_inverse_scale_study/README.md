# #14: the Gamma inverse at the final scale

**Status: registered; the holdout is not yet evaluated.**

`K_STATS_Gamma_InverseCumulative` and `K_STATS_ChiSquare_InverseCumulative`
lose a quantile that is representable at the final scale when its unit-scale
value is below the Double range. The rule registered here answers small unit
quantiles in log space and forms the result once, at the final scale.

| File | Role |
| --- | --- |
| [`PREREGISTRATION.md`](PREREGISTRATION.md) | The rule, its crossover, the acceptance rules G1 to G7, the contracts and the predictions, frozen before the holdout is evaluated |
| `inv_scale_model.py` | The log branch as a model over the pair routines of `gamma_route_study/route_model.py` |
| `holdout_build.py`, `gamma_inverse_scale_grid.csv` | The 455 holdout rows and their references; `--check` proves the grid is the design |
| `M_STATS_PROBDIST_GIS.bas` | `Export_GammaInverseScaleHoldout`: the public inverses, the round trip and the classifying value at every row, as bit patterns. The runner exports it as the study grid `gis_holdout` |
| `holdout_analyze.py` | The verdict against G1 to G7 |
| `test_holdout.py` | Fixtures for the model, the construction and each rule, on another seed and other seam shapes |
