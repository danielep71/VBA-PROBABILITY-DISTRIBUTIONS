# Preregistration — Student-t large-df tail correction (#34)

**Status: FROZEN before any coefficient or value was generated.** This file
is the design. The derivation commit that follows produces results against
it and does not amend it; if the design proves wrong, the failure is
recorded and a new preregistration is written.

Governs: `benchmark/student_t_large_df_study/`. Excel-free. No `.bas`
change, no contract, no threshold, no grid, no observation, no provenance
manifest. The 559-row `benchmark/holdout/holdout_grid.csv` is **not
inspected** at any point in this study.

## 1. Defect under study

`PROB_TryStudentTTail` evaluates the Student-t upper tail through the
regularized incomplete beta `I_x(df/2, 1/2)` with `x = df/(df+t²)`. At
`df ≳ 1E6` the Lentz continued fraction in binary64 is ill-conditioned:
every coefficient is `1 − O(m/a)` with `a = df/2`, rounding accumulates
faster than the fraction converges, and the achievable accuracy is ~1E-9
regardless of stopping rule. Committed counterexample: `StudentT_Survival`
at `t = 2, df = 1E8` is 1.17E-9 relative; the inverse at `p = 0.99,
df = 1E8` inherits it at ≈968 000 ulps. The complement route escapes only
where `y = t²/(df+t²)` is small enough to converge before rounding builds,
which is why it is clean to `t ≈ 4` and fails again at `t = 6`. Moving the
route boundary relocates the failure; it does not remove it.

The density is unaffected — it already has a `large_df` route clean to
`df = 1E20` — and that route is the precedent for the fix.

## 2. Candidate fix

A large-df asymptotic expansion of the upper tail in powers of `1/df`:

    S(t; df) = Q(t) + φ(t) · Σ_{k=1..K} g_k(t) / df^k

with `Q` the standard-normal upper tail, `φ` the standard-normal density,
and `g_k` polynomials in `t` **derived symbolically**. No coefficient is
taken from memory or from a reference table. A check against the classical
published leading term `g_1 = (t³ + t)/4` is permitted as a sanity
assertion on the derivation; it is not a source of coefficients.

## 3. Oracles

| Role | Method | Precision | Standing |
| --- | --- | --- | --- |
| Primary | `_ibeta.t_cdf` — regularized incomplete beta via mpmath, sign-aware | 60 and 100 dps | evidential |
| Independent | tanh-sinh quadrature of the Student-t density on the tail | 100 dps | evidential, algorithm-independent |
| Coarse | R base `pt` (double) if available | double | not evidential — gross-error detection only |

Rmpfr has no incomplete-beta primitive. That is a documented oracle gap of
the same kind as the `mpfr_gamma_inc` ceiling recorded in #22, and it is
recorded rather than worked around. The independent leg is quadrature.

**Oracle stability requirement.** The primary at 60 dps and at 100 dps must
agree at every design point to at least 50 significant digits. Failure is a
study failure, not a data point.

## 4. Precision pair

Validation of the expansion is performed at **60 dps and 100 dps**. Both
must report the same truncation error at every point to within 2
significant digits of the error itself, or the point is rejected.

## 5. Design — forward `df × t` (CDF and survival)

Disjoint from the inverse design in §6. Values are exact binary64.

**Fitting df:** `1E3, 1E4, 1E5, 1E6, 1E7, 1E8`
**Fitting t:** `0.5, 1, 1.5, 2, 3, 4, 5, 6, 8`

**Holdout df (interleaved, disjoint):** `3E3, 3E4, 3E5, 3E6, 3E7`
**Holdout t (interleaved, disjoint):** `0.75, 1.25, 1.75, 2.5, 3.5, 4.5, 7`

54 fitting points, 35 holdout points. The holdout is generated fresh from
these values; it is not drawn from any committed grid.

## 6. Design — inverse (reuse of #22's frozen Student-t sets)

For each df in `{1E6, 1E7, 1E8}`:

**Fitting p:** offsets `0.5 ± 2^-53, 2^-43, 2^-33, 2^-23`; dyadic bridge
`0.625, 0.75, 0.875` and complements; envelope baseline `0.9, 0.99`.
**Holdout p:** offsets `0.5 ± 2^-48, 2^-38, 2^-28, 2^-24`; bridge
`0.5625, 0.6875, 0.8125, 0.9375` and complements.

Inverse validation is a round-trip: the oracle inverts `p` to `q` at 100
dps; the expansion's survival at `q` must recover `1 − p` to within the
same truncation bound as the forward design. The tail residual is
`|Ŝ(q) − (1−p)| / min(p, 1−p)`.

The exact-zero median (`p = 0.5`) is excluded from the round-trip because
the kernel returns `0` exactly by construction and the expansion is never
consulted there.

## 7. Seams

Every seam is probed **immediately below, exactly at, and immediately
above** (offset `±1E-9` relative in t):

1. **CF route switch** — `t_switch(df) = sqrt(df · 1.5 / (df/2 + 1))`,
   i.e. `1.7320` at `df = 1E3` rising to `sqrt(3)` in the limit. Probed at
   every fitting df. This seam belongs to the *existing* code path and is
   probed to characterize it, not to preserve it.
2. **Proposed dispatch seam in df** — the adopted `DF_MIN` (see §8),
   probed at `t` values `1, 2, 4`.
3. **Proposed dispatch seam in |t|** — the adopted `T_MAX`, probed at
   `df` values `1E6, 1E8`.

At each seam the two candidate routes (binary64 CF and the expansion) are
both evaluated; the design records which is more accurate on each side so
the dispatch predicate is a measured boundary, not a guessed one.

## 8. Crossover decision rule

Let `E_K(df, t)` be the relative error of the order-`K` expansion against
the 100-dps primary. Let `C(df, t)` be the relative error of the
binary64 Lentz CF (faithful reproduction, `PROB_NUM_EPS = 3E-14`) at the
same point.

**Accuracy bound:** `B = 2^-50` (≈ 8.9E-16, four binary64 ulps).

**Order selection.** The adopted order `K*` is the smallest `K ≤ 8` such
that `E_K(df, t) ≤ B` at every fitting point with `df ≥ DF_MIN` and
`|t| ≤ T_MAX`.

**Region selection.** Candidates are `DF_MIN ∈ {1E5, 1E6}` and
`T_MAX ∈ {4, 6, 8}`. The adopted region is the one with the smallest
`DF_MIN` and, at that `DF_MIN`, the largest `T_MAX` for which some `K* ≤ 8`
exists.

**Improvement requirement.** Inside the adopted region, at every fitting
point where `C(df, t) > B`, the expansion must satisfy
`E_{K*}(df, t) < C(df, t)`. The expansion may not be adopted anywhere it is
worse than what it replaces.

**Holdout confirmation.** The adopted `(K*, DF_MIN, T_MAX)` must then hold
at every holdout point in §5 and every inverse holdout point in §6, with
no exception. Holdout is inspected once, after `K*`, `DF_MIN` and `T_MAX`
are fixed from fitting alone.

**Dispatch form.** The output is a bounded predicate
`df ≥ DF_MIN AND |t| ≤ T_MAX`, with the CF route retained outside it. No
finer partition is adopted from this study; if the simple predicate does
not meet the bound, the study fails rather than searching for a shape that
fits.

## 9. Failure conditions

Any one of these fails the study. A failed study is recorded as such and
does not produce a dispatch recommendation:

- the derived `g_1` does not equal `(t³ + t)/4`;
- the derived coefficients are not exact rationals;
- `E_K` does not decrease monotonically in `K` at any fitting point with
  `df ≥ 1E6`, `|t| ≤ 4`;
- no `K ≤ 8` meets `B` for any candidate region;
- oracle instability (§3) or precision-pair disagreement (§4) at any point;
- the improvement requirement (§8) is violated anywhere inside the adopted
  region;
- any holdout point (§5 forward or §6 inverse) violates the bound after
  adoption;
- any seam probe shows the two routes disagreeing by more than `B` on the
  side where both are claimed accurate.

## 10. Recorded outputs

The derivation commit must record, machine-readably:

- every `g_k` as an exact rational polynomial, with the SymPy expression
  that produced it;
- a fixture that recomputes each `g_k` and compares to the stored form;
- `E_K` for `K = 1..8` at every fitting point, both precisions;
- `C` at every fitting point;
- the worst point identity for each `K`;
- oracle stability figures;
- the adopted `(K*, DF_MIN, T_MAX)` and the decision-rule trace that
  produced them;
- holdout and inverse round-trip results, evaluated once;
- seam probe results;
- runtime.

## 11. What this study does not do

It writes no VBA. It changes no contract, threshold, grid, observation,
or manifest. It does not inspect the 559-row holdout. It produces a
validated series and a measured dispatch predicate that the Phase 1 export
session can implement and verify against real Excel; it is not itself
evidence that the kernel is fixed.
