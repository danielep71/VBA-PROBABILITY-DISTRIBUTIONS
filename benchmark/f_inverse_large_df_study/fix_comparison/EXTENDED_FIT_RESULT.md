# P2 extended fit: rejected, 2026-11-12

Protocol and runner frozen at `a405124816e9a6f472e96c4c905b821c9fa579ef`
before the 24 identities were evaluated. Results: `extended_fit_result.json`.
No candidate change during the run. All source bindings verified after it.

**22 rows pass, 2 fail; no oracle/model blockers.** The machine status
`REJECT_OR_BLOCKED` is specifically a candidate rejection here. Every oracle
agrees at 50/80 dps to 40 relative digits and its CF converges explicitly.
The earlier 21-row screen remains a limited success, not adoption evidence.

| p | df1=df2 | F-0 relative F error | P2 relative F error | Newton F-0 / P2 | Failing gate |
|---:|---:|---:|---:|---:|---|
| 0.05 | 1e8 | 2.18880e-17 | 2.18880e-17 | 11 / 32 | Balanced-huge Newton count exceeds 2x |
| 0.999 | 1000 | 4.19439e-17 | 4.07119e-16 | 14 / 61 | F error increases about 9.71x |

The second error remains below 2^-48, but fails the independent, unchanged
2x nonregression requirement. It must not be reported as a pass by quoting
only the absolute threshold. Median Newton iterations improve from 11.5 to
9.5; that aggregate does not waive either pointwise failure.

## Trace diagnosis, not a validated fix

At p=0.05 and df1=df2=1e8, P2 reaches the final returned beta root
0.4999177573190031 at iteration 9. At iteration 10 its residual is
4.63656890659081e-14. The rounded Newton proposal equals the current upper
bracket endpoint; the bracket guard replaces it with the midpoint
0.49989092032227594 before the stagnation check can accept it. P2 then takes
32 iterations overall, with 24 bisection fallbacks versus 3 for F-0.
The recorded loop CF iterations sum to 14133 versus 1576 (excluding the
orientation probe). This localizes the extra work to the interaction of
forward rounding, bracket rejection and stagnation handling; it does not
prove that reordering the checks is a safe fix.

Next design work must address both stagnation and balanced-control rounding,
with endpoint/false-convergence safeguards and new prospective validation.
No tolerance relaxation, production VBA edit, independent holdout access,
seam/invariance run or candidate-specific Excel run follows this rejection.
#35 stays open and PR #326 stays draft. Historical protocols/results remain
unchanged. The CI Excel regression on bcece52 passed existing production
code; it does not validate this Python-only candidate in Excel.

Validation: 15 focused tests pass, including an analytic uniform-Beta CF
check and an exhaustion test that rejects partial oracle fractions; Ruff
passes for the study directory. The complete repository unit-suite result
is recorded in the PR when available.
