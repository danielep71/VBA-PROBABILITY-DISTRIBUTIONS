# #307 Gamma routing study

#307: since #242, every public Gamma P and Q call runs the double-double
kernel, at 5 ms to 170 ms per call. The binary64 routines it replaced take 2
to 39 microseconds. This study looks for a routing rule: binary64 wherever it
meets the contracts, double-double only where it does not.

Step 1 is a map of the binary64 route in Excel. It is a fitting set; the rule
derived from it will be preregistered and validated on untouched points.

| File | Role |
| --- | --- |
| [`MAP_DESIGN.md`](MAP_DESIGN.md) | The frozen map design: shapes, arguments, references and scoring |
| `design.py` | The map's inputs |
| `generate_binary64_map.py` | Builds `binary64_map_grid.csv` with high-precision references; `--check` proves the committed grid is the design |
| `M_STATS_PROBDIST_GBR.bas` | `Export_GammaBinary64Map`: calls the binary64 series and continued fraction at every point. The runner exports it as the study grid `gbr_binary64_map` |
| `analyze_binary64_map.py` | Scores the pre-#242 route against the governing contract at every point |
| `test_binary64_map.py` | Proves the grid is the design and the analyzer forms the pre-#242 complement; it runs in the evidence-tool suite |
| [`MAP_RESULT.md`](MAP_RESULT.md), `map_results.json` | The map's result (run 42783017314) |
| `analyze_exact_deviance.py` | Diagnosis: the binary64 route with an exact deviance |
| [`HOLDOUT_DESIGN.md`](HOLDOUT_DESIGN.md), `holdout_design.py` | The routing holdout, frozen before the map was read: inputs, references, rules H1 to H6 |
| `check_disjoint.py` | Proves no holdout shape, or its doubling, occurs in any tracked file |
| `accurate_deviance.py` | Prototype: Loader's deviance as a pair from error-free transformations |
| `route_model.py` | The repaired binary64 route, statement by statement, with VBA's 64-bit intermediates |
| `b64_route_vba.py` | Generates the route into `M_STATS_PROBDIST_SPECIALFUNCS.bas` (`PROB_TryGammaB64SeriesP`, `PROB_TryGammaB64ContinuedFractionQ`); `--check` proves the module carries it |
| `test_route_model.py` | Pins the production procedures the model ports, the generated block and its constants, the deviance's accuracy and the routed tails on the map |
| [`FIT_DESIGN.md`](FIT_DESIGN.md), `fit_design.py` | The fit set: fidelity, accuracy and cost of the repaired route, and the frozen rule for the shape ceiling A2 |
| `generate_b64_fit.py` | Builds `b64_fit_grid.csv` with references; `--check` proves the committed grid is the design |
| `M_STATS_PROBDIST_GBF.bas` | `Export_GammaB64Fit`: calls and times both repaired routines at every point; study grid `gbf_b64_fit` |
| `analyze_b64_fit.py` | Fidelity bit for bit, routed accuracy, cost fit and A2 |
| `test_b64_fit.py` | Proves the grid is the design and the analyzer's fidelity, cost and ceiling logic |
| [`FIT_RESULT.md`](FIT_RESULT.md), `fit_results.json` | The fit's result (run 42798193976): 392/392 bit for bit, every routed tail within contract, A2 = 10,000 |

## Phase B checkpoint

The [frozen A2=10000 rule](ROUTING_RULE.md) is rejected by a
[development seam preflight](PHASE_B_PREFLIGHT.md): four CDF monotonicity
regressions in the source-derived model. The holdout remains untouched and
public routing remains unchanged.
