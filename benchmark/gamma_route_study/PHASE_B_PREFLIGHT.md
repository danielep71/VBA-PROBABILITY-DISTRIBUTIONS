# #307 phase B: routing rejected before holdout

Rule `63f88c895d4040cd4a7c25c2cf19b0daf34c90c6` fixes A2=10000.
A development replay at the 26 eligible shapes of the already revealed fit
set tests predecessor/fl(A+1)/successor. These are development probes at known
shapes; they are not the untouched holdout and not all are original fit rows.

The proposed routed P decreases across the series/continued-fraction seam:

| Shape | Drop in binary64 ulps |
| --- | ---: |
| 1e-8 | 1 |
| 1e-4 | 1 |
| 0.1 | 2 |
| 10 | 1 |

The unchanged double-double model is monotone over all 26 triples. Thus the
proposed switch introduces regressions in the source-derived model. Phase A
proved pointwise fit accuracy and host fidelity of the component routines;
that does not prove monotonicity of their composition. These are model results,
not a claim that this new adapter has executed in Excel.

`preflight_phase_b.py` reproduces `phase_b_preflight.json`, including exact
hexadecimal inputs/outputs, baseline outputs and input hashes. It exits 1 on
rejection. The existing fit test replays this rejection and checks its exact
counterexample set and the baseline monotonicity; the new route is not accepted
by treating these failures as expected production behavior.

No production adapter, regression expected values, threshold or exported grid
has been changed. The original 117-point holdout remains unevaluated. #307
stays open. Next, develop a separately identified seam-safe candidate on these
revealed development points, preserving the A2 ceiling, then freeze the revised
rule and complete probes before consuming the holdout. Do not silently edit
the rejected rule or waive monotonicity. Raising A2 is separate research.
