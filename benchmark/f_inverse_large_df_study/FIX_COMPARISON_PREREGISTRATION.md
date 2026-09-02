# Preregistration — prospective fix comparison for #35

**Status: FROZEN before any candidate is evaluated.** Governs
`benchmark/f_inverse_large_df_study/fix_comparison/`. Excel-free. No `.bas`
change; no contract, threshold, grid, observation or provenance manifest;
the 559-row holdout is not inspected.

This is a **new** protocol. `e68c7ad` is cited as design evidence only: its
preregistered isolation did not meet its own criteria (§5 `NONE`, §6
failed) and is not treated as confirmation. What `e68c7ad` established is
narrower and is all this document relies on — that carrying the solved
member at high precision into pair formation removes the error on every
committed counterexample, and that the seven-component decomposition
could not localise that because the mechanism straddles two of its
components. This protocol therefore models the **actual VBA operation
order** and compares candidate fixes prospectively.

## 1. The defect, as recorded on #35

`PROB_TryBetaInvRegularized` selects its solving variable by comparing the
two probabilities (`Probability <= ComplementProbability`), solves for that
member `U`, and reconstructs the other as `1# − U`. Which member is tiny is
determined by the shapes. When the branch solves the large member, the
tiny one inherits ~eps absolute error, `eps/tiny` relative, and
`K_STATS_F_InverseCumulative`'s `BetaX/BetaY` is dominated by it.

Blast radius (from #35): `K_STATS_F_InverseCumulative`, the public Beta
inverse, and Student-t's small-df inverse path, which share the routine.

## 2. Faithful model — the actual operation order

`vba_order_model.py` reproduces `PROB_TryBetaInvRegularized` step for step
in binary64 (Python `float`), with the reference at 50 dps. Nothing is
decomposed into swappable components; the model exists to be *faithful*,
and candidates are modelled as **edits to it**, exactly as they would be
edits to the VBA.

Sequence, from the source:

1. **Branch.** `SolveDirect = (p <= 1 − p)`. If direct: `Sa, Sb, Target = A, B, p`;
   else `Sa, Sb, Target = B, A, 1 − p`.
2. **`LogBetaAB = PROB_LogBeta(Sa, Sb)`** at binary64.
3. **Seed.** If `Sa > 1 And Sb > 1`: the Cornish-Fisher form —
   `Z = Φ⁻¹(Target)`, `R = (Z² − 3)/6`, `S1 = 1/(2Sa − 1)`, `S2 = 1/(2Sb − 1)`,
   `HH = 2/(S1 + S2)`, `W = Z·√(HH + R)/HH − (S2 − S1)(R + 5/6 − 2/(3HH))`,
   `U = Sa/(Sa + Sb·e^{2W})`. Else `U = exp((log Target + log Sa + LogBetaAB)/Sa)`.
   Fallback `U = 0.5·Target + 0.25` if out of `(0, 1)` or non-finite.
4. **Bracketed Newton**, `Low = 0`, `High = 1`, up to `PROB_INV_MAX_ITER`:
   - forward `Ibeta = PROB_TryBetaRegularized(U, 1# − U, Sa, Sb)` — **the
     complement argument is `1# − U` in binary64**;
   - `Residual = Ibeta − Target`; tighten `Low`/`High` by its sign;
   - density `exp((Sa−1)·Log(U) + (Sb−1)·Log(1# − U) − LogBetaAB)` — **again
     `Log(1# − U)`**;
   - `UNew = U − Residual/Density`, bisecting to `(Low + High)/2` if the step
     leaves the bracket or the density is zero;
   - stop when `|UNew − U| <= PROB_MACH_EPS·|UNew|` or `UNew = U`.
5. **Pair unwind.** Direct: `X = U, Y = 1# − U`; else `X = 1# − U, Y = U`.
6. **F transform.** `exp(Log d2 − Log d1 + Log X − Log Y)`.

Inside step 4, `PROB_TryBetaRegularized` is modelled with its Loader
decomposition for `A + B >= PROB_IBETA_LOADER_MIN_SHAPE = 1000` (deviance
form via `PROB_TryDeviancePart`), the literal form below it, the standard
route switch, and the continued fraction with `PROB_NUM_EPS = 3E-14` and
`PROB_BETA_MAX_ITER`; exhaustion fails, never returns a partial sum.

**Faithfulness criterion** (learned from `e68c7ad`, whose ×4 magnitude bound
assumed a deterministic error): on the eight retained counterexamples the
model's relative error must lie in `[eps/(8·m), 8·eps/m]` with
`m = min(x, 1−x)` — the mechanism's scaling, not a fixed magnitude — and
on the five retained controls it must be below `2^-48`. The model must
also reproduce the VBA's Newton iteration count to within ±2 on every row
where the count is known. Failure means the model is not faithful; record
and stop.

## 3. Candidates

Each is a minimal edit to step 1 of the model. None touches steps 2–6.

**F-A — solve for the small member, decided by one forward evaluation.**
`SolveDirect = (p <= I_{1/2}(A, B))`. The root `x` is then ≤ 1/2 when
solving direct and the root `y` is < 1/2 when solving complement, so the
reconstructed member is always ≥ 1/2 and `1# − U` loses nothing. Cost: one
extra `PROB_TryBetaRegularized` call at `U = 1/2` before the loop.

**F-B — solve for the small member, decided by the seed.** Compute the
existing seed for the direct branch; if it exceeds 1/2, switch to the
complement branch (recomputing the seed there). No extra forward
evaluation. Near `p ≈ I_{1/2}` the seed may choose either side, but there
both members are ≈ 1/2 and neither choice loses precision.

**F-0 — baseline.** The current rule, unchanged. Every candidate is
compared to it, and the balanced comparator in §4 must not regress from
it.

A candidate that solves the large member and attempts to *recover* the
tiny one afterwards in binary64 is not included: `e68c7ad` shows the
information is gone once `U` is rounded, so no post-hoc reconstruction in
binary64 can succeed.

## 4. Design

**Retained rows** — the eight counterexamples and five controls from
`PREREGISTRATION.md` §2, exact binary64 identities as recorded there.

**Prospective fitting grid**, disjoint from the retained rows:

- `p ∈ {0.05, 0.25, 0.75, 0.999}`
- `(d1, d2) ∈ {(1E5, 2), (1E7, 4), (3, 1E7), (10, 1E9), (1E3, 1E3), (1E8, 1E8)}`

`(1E8, 1E8)` is #34's balanced shared-CF comparator; it is in the design so
that any candidate is shown **not to regress it**, and its own error is
#34's concern, not #35's.

**Holdout grid**, interleaved and disjoint:

- `p ∈ {0.1, 0.4, 0.6, 0.98}`
- `(d1, d2) ∈ {(3E5, 7), (7, 3E5), (50, 1E8), (1E8, 50)}`

**Shared-path invariance set** — Student-t small-df inverse via the same
routine: `df ∈ {1, 2, 3, 5, 10, 30}` at the #22 frozen fitting offsets and
bridge points. A candidate must leave every point within one binary64 ulp
of baseline or improve it; it may not degrade any.

Holdout and the invariance set are evaluated **once**, after §5 has been
applied to the retained rows and the fitting grid.

## 5. Measures and decision rule

Per point: relative error of `F` against the 50-dps reference; relative
error of each pair member; tail residual `|I(q̂) − p| / min(p, 1−p)` via the
50-dps forward; Newton iteration count.

A candidate is **adoptable** iff, against F-0:

1. every retained affected row falls below `2^-48`;
2. no retained control rises above `2^-48` or worsens by more than ×4;
3. no fitting-grid point worsens by more than ×2 or rises above `2^-48`
   where F-0 was below it;
4. the balanced `(1E8, 1E8)` points are within ×2 of F-0 in every measure;
5. median Newton iterations across the fitting grid do not exceed F-0's
   median by more than 2;
6. after 1–5 hold, the holdout grid and the invariance set satisfy 1–4
   with no exception.

If both F-A and F-B are adoptable, the one with fewer forward evaluations
per call is preferred unless the other is strictly more accurate at any
retained row — in which case both are reported and the choice is a
maintainer decision, not this study's. If neither is adoptable, that is
the result.

## 6. Failure criteria

- model not faithful under §2;
- no candidate adoptable;
- a candidate passes retained rows but fails holdout or invariance (record
  as a fitting-only success — it is not adoptable);
- oracle instability: the 50-dps reference must agree with a 80-dps
  recomputation to 40 digits at every point.

## 7. What this study does not do

It writes no VBA. The adoptable candidate, if any, becomes the #35 patch
that lands only inside a real-Excel wave with regressions, exports and
truthful provenance — and only after its effect on #34's shared-CF
comparator has been checked in that same wave. It selects nothing beyond
the comparison above.
