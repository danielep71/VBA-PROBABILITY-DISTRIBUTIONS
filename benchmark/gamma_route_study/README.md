# #307 Gamma routing study

**Status: closed.** #307 was closed on 2026-11-11 under the maintainer's
revised scope ([current disposition](../../docs/research/GAMMA_PERFORMANCE_STATUS.md)).
No routing rule ships: the frozen A2 = 10,000 rule is rejected (below), and
the 117-point routing holdout remains unevaluated. The repaired binary64
routines that #312 generated into production were never called by a public
function and left production after the closure; `b64_route_vba.py`
still reproduces them byte for byte. Statements in the dated documents below
that #307 "remains open" predate the closure.

The question was the per-call cost of the double-double kernel that every
public Gamma P and Q call has run since #242: 5 ms to 170 ms per call when the
study began, against 2 to 39 microseconds for the binary64 routines it
replaced. The study looked for a routing rule: binary64 wherever it meets the
contracts, double-double only where it does not. It started from an Excel map
of the binary64 route, used as a fitting set.

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
| `b64_route_vba.py` | Generates the route's VBA (`PROB_TryGammaB64SeriesP`, `PROB_TryGammaB64ContinuedFractionQ`) as the fit candidate `88ff6ba` carried it; `--check` proves that, and that production no longer carries it |
| `test_route_model.py` | Pins the production procedures the model ports, the generated block and its constants, the deviance's accuracy and the routed tails on the map |
| [`FIT_DESIGN.md`](FIT_DESIGN.md), `fit_design.py` | The fit set: fidelity, accuracy and cost of the repaired route, and the frozen rule for the shape ceiling A2 |
| `generate_b64_fit.py` | Builds `b64_fit_grid.csv` with references; `--check` proves the committed grid is the design |
| `M_STATS_PROBDIST_GBF.bas` | `Export_GammaB64Fit`: calls and times both repaired routines at every point. It compiles only against the fit candidate's source, so the study grid `gbf_b64_fit` is no longer registered for export |
| `analyze_b64_fit.py` | Fidelity bit for bit, routed accuracy, cost fit and A2 |
| `test_b64_fit.py` | Proves the grid is the design and the analyzer's fidelity, cost and ceiling logic |
| [`FIT_RESULT.md`](FIT_RESULT.md), `fit_results.json` | The fit's result (run 42798193976): 392/392 bit for bit, every routed tail within contract, A2 = 10,000 |
| [`ROUTING_RULE.md`](ROUTING_RULE.md) | The phase-B rule frozen at A2 = 10,000, before any holdout evaluation |
| [`PHASE_B_PREFLIGHT.md`](PHASE_B_PREFLIGHT.md) | The seam preflight that rejects it: four CDF monotonicity regressions |
| [`TRANSITION_DEVELOPMENT.md`](TRANSITION_DEVELOPMENT.md) | Error diagnosis, the rejected one-sigma band and the smaller-tail candidate |
| [`ARITHMETIC_DIAGNOSIS.md`](ARITHMETIC_DIAGNOSIS.md) | Isolates the arithmetic error on 17 retained failing pairs |
| [`PAIR_CANDIDATE.md`](PAIR_CANDIDATE.md) | B4: the residual-preserving pair recurrence |
| [`SCALAR_CANDIDATE.md`](SCALAR_CANDIDATE.md) | B5: scalar pairs in the B4 recurrences; a retained prototype, not wired into public functions |

## Phase B checkpoint

The [frozen A2=10000 rule](ROUTING_RULE.md) is rejected by a
[development seam preflight](PHASE_B_PREFLIGHT.md): four CDF monotonicity
regressions in the source-derived model. The holdout remains untouched and
public routing remains unchanged.

[Transition development](TRANSITION_DEVELOPMENT.md) records the independent
error diagnosis, a rejected one-sigma DD band, and the smaller-tail candidate
that passes 48 triples but fails the wider neighboring-input replay. The
holdout remains untouched; further fast-kernel arithmetic work is required.
