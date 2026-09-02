# Unbalanced-F inverse isolation (#35)

**Status: preregistered criteria NOT met — recorded, not amended.** The
mechanism in `PREREGISTRATION.md` §3 is confirmed by a post-hoc diagnostic,
but the preregistered single-component sweep (§5) cannot isolate it and the
faithfulness criterion (§6) fires. Both are defects in the preregistration,
not in the hypothesis, and the design says such a result is recorded rather
than tuned. Fix selection has not begun.

## What was found

`PROB_TryBetaInvRegularized` selects its solving variable by comparing `p`
with `1−p`, then forms the unsolved member of the `(x, y)` pair as `1# − U`.
Which member is tiny is fixed by the shapes, not by `p`. Whenever the branch
solves for the large member, the tiny member inherits ~eps absolute error —
`eps/tiny` relative — and `K_STATS_F_InverseCumulative`'s `BetaX/BetaY` is
dominated by it.

Eight committed rows match the predicted `c·2^-52/min(x,1−x)` within a
factor 0.4–0.83; five clean controls are clean because the branch happens to
solve the tiny member. Source term identified before the model was built,
disclosed in §0.

## What the model showed

| Result | Outcome |
| --- | --- |
| Clean controls reproduce as clean (all five) | ✓ |
| Affected rows reproduce as affected, at the `eps/tiny` scale | ✓ |
| Affected magnitudes within ×4 of observed (§6) | **✗** — model 4–100× smaller |
| Any single component satisfies §5 | **✗** — none, including C6 |
| C5 + C6 substituted **together** (post-hoc) | worst **1.38E-15** — error gone |

The magnitude shortfall is not a flaw in the mechanism: the tiny member's
error is a count of ulps in `U`, and the model's correctly-rounded forward
lets Newton land within ~1 ulp where the VBA's CF lands a few. A per-row ×4
criterion assumed a deterministic magnitude for what is a rounding outcome.

The single-substitution failure is a decomposition defect. The mechanism is
"`U` is carried into pair formation at binary64" — C5 alone hands a rounded
`U` to a binary64 subtraction; C6 alone subtracts from an already-rounded
`U`. Neither can pass §5; together they do.

## Correction made along the way

`f_inverse_model.py` was briefly broken by a misread traceback: a test script's
`from mpmath import mp` error was "fixed" in the module, whose `import mpmath
as mp` needs `mp.mp.dps`. The change made every precision setting a no-op and
the model ran at 15 digits. Caught when three "50-dps" answers disagreed at
1E-10; every number here post-dates the correction.

## What this study does not do

It selects no fix. It writes no VBA. It changes no contract, threshold, grid,
observation or manifest. The 559-row holdout was not read. `ed03f15` landed a
stray copy of `PREREGISTRATION.md` in `t_density_large_df/`; removed in
`faf6350`.

## Files

| File | Role |
| --- | --- |
| `PREREGISTRATION.md` | frozen design, with §0 disclosure |
| `f_inverse_model.py` | seven-component swappable model |
| `run_isolation.py` | applies §5–6 mechanically; writes `results.json` |
| `results.json` | faithfulness, both sweeps, per-component verdicts, post-hoc |
