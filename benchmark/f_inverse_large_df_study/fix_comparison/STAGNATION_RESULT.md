# P3 guarded stagnation: staged screen passes, 2026-11-12

Protocol, model and runner frozen before evaluation at
`4e6031c659324826de32c4ba29e4dbe72225d9ff`. Machine evidence is
`stagnation_result.json`; all bound files were hash-verified after the run.
P2's rejection and every earlier frozen study remain unchanged.

**53/53 rows pass: SCREEN_PASS_NOT_ADOPTION.** All 45 revealed identities
passed before any of the eight new identities were evaluated. References
agree at 50/80 dps to 40 relative digits, with explicit oracle CF convergence.
No blockers or threshold exceptions were introduced.

| Population | Passed | Maximum P3 relative F error | Median Newton F-0 / P3 |
|---|---:|---:|---:|
| Retained affected/controls | 13/13 | 3.44473e-15 | 12 / 9 |
| Previously revealed P2 screen | 8/8 | 2.02875e-15 | 24 / 8.5 |
| Original extended fitting | 24/24 | 6.56980e-15 | 11.5 / 9 |
| New prospective screen | 8/8 | 6.14036e-16 | 8 / 8 |

The original fitting gate is relative to baseline: a point already above
2^-48 may pass if it does not regress beyond 2x. Thus this result is **not**
a claim that all 53 quantiles are below 2^-48. All eight new points are below
that bound and within 2x baseline error. Both median iteration gates pass.

## Resolution of P2's two fitting failures

| p | df1=df2 | F-0 F error | P2 F error | P3 F error | Newton F-0 / P2 / P3 |
|---:|---:|---:|---:|---:|---:|
| .05 | 1e8 | 2.18880e-17 | 2.18880e-17 | 2.18880e-17 | 11 / 32 / 10 |
| .999 | 1000 | 4.19439e-17 | 4.07119e-16 | 4.19439e-17 | 14 / 61 / 13 |

In both cases a guarded stagnation acceptance avoids the unnecessary distant
bisection. The guard requires an unchanged rounded Newton proposal, finite
positive density, strictly interior adjacent floats, finite monotone neighbor
CDF residuals enclosing zero, and a current residual no larger in magnitude
than either neighbor's. All other P2 numerical choices remain unchanged.
The guard is checked before the bracket replaces a rounded endpoint proposal.
The original model and P2 files are not modified.

## Validation and remaining gates

19 focused tests pass, including unsafe neighbor ordering, unbracketed targets,
nonfinite inputs, endpoints, moving proposals and state restoration after an
exception. The full repository unit suite passes 397 tests with four skips;
Ruff passes for the study directory. The existing registered evidence-test
entry includes the new tests. No VBA candidate has been executed in Excel.

This resolves the two observed P2 regressions in the study model. It does not
establish global safety, speed in VBA, or adoption readiness. Before a VBA
patch: prospective geometry/CF seam tests; Beta and Student-t shared-path
invariance with faithful small-shape support; genuinely independent holdout;
and exact-source Excel 365 64-bit validation. The current forward model still
refuses some shapes below 0.5 and other unmodelled branches. Those capability
gaps must be resolved explicitly, not counted as numerical passes.
#35 remains open; PR #326 remains draft. No merge or production edit occurred.
