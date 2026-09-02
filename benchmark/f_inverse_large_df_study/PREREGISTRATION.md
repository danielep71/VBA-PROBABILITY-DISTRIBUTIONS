# Preregistration — unbalanced-F inverse accuracy isolation (#35)

**Status: FROZEN.** Governs `benchmark/f_inverse_large_df_study/`. Excel-free.
No `.bas` change, no contract, threshold, grid, observation or provenance
manifest. The 559-row `benchmark/holdout/holdout_grid.csv` is not inspected.

## 0. Disclosure of what preceded this document

This study is **confirmatory**, not exploratory, and that must be stated
plainly. Before the model in §4 was built, two things had already happened:

1. The committed grid rows in §2 were read and their errors tabulated.
2. From the side-flipping pattern in that table — errors at `p = 0.5` when
   `d1 ≫ d2`, at `p ≥ 0.9` when `d1 ≪ d2`, and clean otherwise — a
   hypothesis was formed and checked against the source of
   `PROB_TryBetaInvRegularized`, which confirmed the mechanism.

The hypothesis, its prediction table and the source term are recorded in
§3 as they stood before the model was built. The model's job is to confirm
or refute them by component substitution, mechanically, under the protocol
in §5. If the substitution sequence fails to reproduce the pattern, the
hypothesis is recorded as refuted regardless of how convincing §3 reads.

## 1. Two mechanisms, deliberately separated

`F_InverseCumulative` errors at large df fall into two groups with different
causes. This study concerns only the second.

| Group | Rows | Mechanism | Owner |
| --- | --- | --- | --- |
| Balanced, both df huge (`d1 = d2 = 1E8`) | forward CDF/survival at 7.8E-13 | shared binary64 CF conditioning, 1892 iterations | **#34** — mandatory comparator for any CF change |
| Unbalanced, one df huge (`1E6/3`, `1E8/10`, `2.5/1E6`, `5/1E5`, `100/1E10`) | inverse at 7.8E-11 to 9.3E-9 | **not** CF conditioning: the CF with an exact prefactor is accurate to 3E-16 at every one of these | **#35** — this study |

## 2. Counterexamples — exact binary64 identities

From `benchmark/probability_accuracy_grid.csv`, `F_InverseCumulative`. `hex`
is authoritative; `p` values `0.9/0.95/0.99` are not exactly representable.

| p | d1 | d2 | hex (p, d1, d2) | rel err | regime |
| --- | --- | --- | --- | --- | --- |
| 0.5 | 1E6 | 3 | `0x1.0000000000000p-1, 0x1.e848000000000p+19, 0x1.8000000000000p+1` | 7.77E-11 | validated **and** envelope_domain (duplicate) |
| 0.5 | 1E8 | 10 | `0x1.0000000000000p-1, 0x1.7d78400000000p+26, 0x1.4000000000000p+3` | 9.49E-10 | envelope_domain |
| 0.9 | 2.5 | 1E6 | `0x1.ccccccccccccdp-1, 0x1.4000000000000p+1, 0x1.e848000000000p+19` | 1.83E-11 | **validated** |
| 0.95 | 2.5 | 1E6 | `0x1.e666666666666p-1, 0x1.4000000000000p+1, 0x1.e848000000000p+19` | 1.58E-11 | **validated** |
| 0.99 | 2.5 | 1E6 | `0x1.fae147ae147aep-1, 0x1.4000000000000p+1, 0x1.e848000000000p+19` | 8.99E-12 | **validated** |
| 0.9 | 5 | 1E5 | `0x1.ccccccccccccdp-1, 0x1.4000000000000p+2, 0x1.86a0000000000p+16` | 1.40E-12 | validated |
| 0.9 | 100 | 1E10 | `0x1.ccccccccccccdp-1, 0x1.9000000000000p+6, 0x1.2a05f20000000p+33` | 9.31E-9 | envelope_domain |
| 0.99 | 100 | 1E10 | `0x1.fae147ae147aep-1, 0x1.9000000000000p+6, 0x1.2a05f20000000p+33` | 7.00E-9 | envelope_domain |

**Clean controls** — same shapes, other `p`, which the hypothesis must
predict as clean:

| p | d1 | d2 | rel err |
| --- | --- | --- | --- |
| 0.9 | 1E6 | 3 | 4.64E-16 |
| 0.99 | 1E6 | 3 | 1.77E-16 |
| 0.5 | 2.5 | 1E6 | 1.52E-14 |
| 0.9 | 1E8 | 10 | 2.02E-16 |
| 0.5 | 100 | 1E10 | 3.84E-16 |

The `validated` contract carries threshold `2E-10`, so the three
`(2.5, 1E6)` rows pass it legitimately. The label is truthful; whether the
threshold documents an accuracy claim or accommodates this defect is what
#35 decides.

## 3. Hypothesis and source term (recorded before the model)

`K_STATS_F_InverseCumulative` calls
`PROB_TryBetaInvRegularized(p, 1−p, d1/2, d2/2)` for the pair `(BetaX, BetaY)`
and returns `F = (d2/d1) · BetaX / BetaY` in log space.

`PROB_TryBetaInvRegularized` selects its solving variable by comparing the
two **probabilities**:

    If Probability <= ComplementProbability Then
        SolveDirect = True      ' solve for x, then  ResultY = 1# - U
    Else
        SolveDirect = False     ' solve for y, then  ResultX = 1# - U

The member of the pair that is *not* solved for is formed as `1# − U`. Which
member is tiny is determined by the **shapes**, not by `p`: with `a ≫ b` the
quantile `x → 1` and `y` is tiny; with `a ≪ b` the quantile `x → 0` and `x`
is tiny. Whenever the branch solves for the large member, the tiny member
inherits ~eps absolute error, hence `eps / tiny` relative error, and that
member dominates `BetaX / BetaY`.

**Prediction:** relative error ≈ `c · 2^-52 / min(x, 1−x)` with `c ∈ (0.4, 1)`
on rows where the branch solves the large member, and at the Double floor on
rows where it solves the tiny member. Checked against §2 before the model:

| p | d1 | d2 | min(x, 1−x) | predicted | observed | ratio |
| --- | --- | --- | --- | --- | --- | ---: |
| 0.5 | 1E6 | 3 | 2.37E-6 | 9.4E-11 | 7.8E-11 | 0.83 |
| 0.5 | 1E8 | 10 | 9.34E-8 | 2.4E-9 | 9.5E-10 | 0.40 |
| 0.9 | 2.5 | 1E6 | 5.45E-6 | 4.1E-11 | 1.8E-11 | 0.45 |
| 0.95 | 2.5 | 1E6 | 6.93E-6 | 3.2E-11 | 1.6E-11 | 0.49 |
| 0.99 | 2.5 | 1E6 | 1.03E-5 | 2.2E-11 | 9.0E-12 | 0.42 |
| 0.9 | 5 | 1E5 | 9.24E-5 | 2.4E-12 | 1.4E-12 | 0.58 |
| 0.9 | 100 | 1E10 | 1.19E-8 | 1.9E-8 | 9.3E-9 | 0.50 |
| 0.99 | 100 | 1E10 | 1.36E-8 | 1.6E-8 | 7.0E-9 | 0.43 |
| 0.9 | 1E6 | 3 | 5.84E-7 | — (solves tiny) | 4.6E-16 | clean ✓ |
| 0.5 | 2.5 | 1E6 | 1.87E-6 | — (solves tiny) | 1.5E-14 | clean ✓ |

Eight affected rows within a factor 0.4–0.83 of prediction; both controls
clean as predicted. This is the state of knowledge when the model is built.

## 4. Faithful binary64 model

`f_inverse_model.py` reimplements the F inverse path with each component
swappable between **binary64** (Python `float`, matching VBA semantics) and
**50-dps mpmath**:

| Component | binary64 form | high-precision form |
| --- | --- | --- |
| C1 branch | `p <= 1−p` selects x or y | same rule (the rule is the hypothesis; it is not "fixed" here) |
| C2 seed | Cornish-Fisher / mean-based seed in `float` | mpmath seed |
| C3 forward `I` | binary64 Lentz CF with `PROB_NUM_EPS`, Loader prefactor | `_ibeta.ibeta` at 50 dps |
| C4 density | binary64 log-density | mpmath |
| C5 Newton | `float` iterate to `PROB_EPS` | mpmath iterate |
| C6 pair return | `other = 1.0 − U` in `float` | `other = 1 − U` at 50 dps |
| C7 F transform | `exp(log d2 − log d1 + log X − log Y)` in `float` | mpmath |

Reference: `_ibeta.beta_invcdf` at 50 dps, transformed exactly.

## 5. Substitution protocol

Run on every row in §2 including the clean controls.

**Forward sweep.** Start all-binary64. Substitute components one at a time
in the order **C6, C7, C3, C5, C2, C4, C1**, restoring each before the next.
Record the relative error after each single substitution.

**Reverse sweep.** Start all-high-precision. Degrade components one at a
time in the same order. Record the relative error after each single
degradation.

**Responsible term.** A component `Cn` is the responsible term iff:

- substituting `Cn` alone (forward sweep) reduces the error on every affected
  row to below `2^-48` (≈ 3.6E-15), **and**
- degrading `Cn` alone (reverse sweep) raises the error on every affected row
  to within a factor 4 of the observed value, **and**
- neither sweep changes any clean control by more than a factor 4.

C6 is listed first because it is the hypothesis. The order is fixed here so
that "first responsible term found" is not order-dependent in a way that
could be tuned after the fact.

## 6. Failure criteria

- No single component satisfies all three conditions in §5 → hypothesis
  refuted; record and stop.
- More than one component satisfies them → record both; do not pick.
- The all-binary64 model fails to reproduce each observed error to within a
  factor 4 → the model is not faithful; record and stop. A model that does
  not reproduce the defect cannot isolate it.
- Any clean control reproduces as *not* clean in the all-binary64 model →
  same.

## 7. What this study does not do

It selects no fix. Once the responsible term is confirmed, the result is
reported and a fix is chosen separately. It writes no VBA. It changes no
contract, threshold, grid, observation or manifest, and does not inspect the
559-row holdout.
