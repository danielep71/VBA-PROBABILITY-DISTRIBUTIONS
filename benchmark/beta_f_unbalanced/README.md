# Unbalanced-Beta switch study (PROB_LogBeta)

This is the study report for the `PROB_LogBeta` accuracy investigation. It records
the measured behavior of the current implementation across the balanced-to-
extremely-unbalanced argument range, the conclusion that a threshold change alone
cannot fix it, and the remediation that would.

## Purpose

`PROB_LogBeta(A, B)` = `Log(Beta(A, B))` selects among three routes:

- half-integer shortcuts (`A = 1/2` or `B = 1/2`) via `PROB_LogGammaHalfDiff`;
- a one-term asymptotic `LogGamma(Small) - Small*Log(Large)` when
  `Small/Large <= PROB_EPS` (1E-15);
- otherwise the defining identity `LogGamma(A) + LogGamma(B) - LogGamma(A + B)`.

The defining identity cancels catastrophically as the arguments become unbalanced
(`LogGamma(Large)` and `LogGamma(Large + Small)` are large and nearly equal). The
study measures where this matters and whether the switch is well placed.

## Tested values

- **Small values** (the smaller argument), chosen to include non-integer cases:
  `0.8, 1, 1.5, 3, 10`. (Half-integer `0.5` is excluded — it takes the
  `PROB_LogGammaHalfDiff` shortcut.)
- **Ratios** `Small / Large`, swept in decades from `1E-1` down to `1E-18`
  (18 points), so `Large` ranges from `10 x Small` up to `1E18 x Small`.
- Full grid: 5 x 18 = 90 points, each with a 50+ digit mpmath reference.

## Files

- `generate_logbeta_switch.py` — writes `logbeta_switch_grid.csv` (90 rows: 5
  `Small` values x 18 ratios) with 50+ digit mpmath references.
- `logbeta_switch_grid.csv` — the grid; `arg1 = Large`, `arg2 = Small`.
- `M_STATS_PROBDIST_LOGBETA_STUDY.bas` — standalone export macro (`Export_LogBeta_Study`) that
  fills `observed_vba` by calling `PROB_LogBeta` directly.
- `analyze_logbeta_switch.py` — prints measured relative error vs ratio per
  `Small`, flags rows over the 5E-15 Beta claim, marks where the branch fires.

## Measured results (actual VBA `PROB_LogBeta`)

Best achievable from the current two methods (choosing the better of general
identity and one-term asymptotic at each ratio), worst case over `Small`:

| Small / Large | best achievable | meets 5E-15 |
|---:|---:|:--:|
| 1E-1 | 8.9E-16 | yes |
| 1E-2 | 1.4E-14 | no |
| 1E-3 | 2.9E-13 | no |
| 1E-4 | 3.6E-12 | no |
| 1E-6 | 1.8E-10 | no |
| 1E-10 | 1.9E-12 | no |
| 1E-13 | 1.3E-15 | yes |
| 1E-14 | 2.9E-16 | yes |
| <= 1E-15 | ~1E-16 (branch) | yes |

**No single switch position meets 5E-15 across ratios ~1E-2 to 1E-13.** Moving the
threshold only chooses which method's mediocre error applies in the middle band.

## Integer vs non-integer `Small`

The one-term asymptotic is *exact* for integer `Small` (the Gamma recurrence makes
`LogGamma(Large+n) - LogGamma(Large)` telescope to `Log(Large)` terms), but has
material truncation error for non-integer `Small` at moderate ratios (about 1.3%
at ratio 1E-1). An early hypothesis that the branch is "accurate everywhere, just
widen the switch" was an artifact of testing mostly integer `Small`; the benchmark
disproved it. Widening the one-term switch would have produced large errors for
non-integer `Small` — the study prevented an incorrect change.

## Public exposure

Directly affected: Beta density; Beta CDF/survival via incomplete-beta
normalization; Beta inverse; F with disparate degrees of freedom. Student t is
largely protected — its `Beta(df/2, 1/2)` normalization trips the half-integer
shortcut for `df >= 2`, and `df = 1` is `Beta(0.5, 0.5)` (balanced).

## How to reproduce

1. Import `M_STATS_PROBDIST_LOGBETA_STUDY.bas` into the workbook and `Debug > Compile`.
2. Run `Export_LogBeta_Study`; when prompted, select `logbeta_switch_grid.csv`.
   This fills the `observed_vba` column by calling `PROB_LogBeta` directly.
3. Commit the filled CSV.
4. Analyze: `python3 analyze_logbeta_switch.py` prints the measured relative
   error versus ratio per `Small`, flagging rows over the 5E-15 Beta claim and
   marking where the asymptotic branch engages.

To regenerate the references from scratch: `python3 generate_logbeta_switch.py`
(requires mpmath). This overwrites `logbeta_switch_grid.csv` with empty
observations, so re-run the export afterwards.

## Current production status

The limitation is documented (this study, the `PROB_LogBeta` ACCURACY LIMITATION
note, the benchmark README caveat) and guarded by an expected-limitation test in
`M_STATS_PROBDIST_TEST` that trips if the gap is ever closed. The switch itself is
left unchanged, because widening it with the current one-term asymptotic is worse.

## Proposed remediation (deferred, validated feature)

A stable log-gamma increment kernel `PROB_LogGammaDelta(Large, Increment)` =
`LogGamma(Large + Increment) - LogGamma(Large)` computed without subtracting two
large independently rounded values, dispatched in three regimes:

- balanced -> direct identity;
- moderately unbalanced -> cancellation-free Lanczos log-gamma difference;
- strongly unbalanced -> validated multi-term digamma/Bernoulli expansion.

Acceptance: the ratio x non-integer-`Small` x absolute-`Large` grid, switch-seam
continuity, public Beta/F paths, and `LogBeta(A,B) = LogBeta(B,A)` symmetry.
Accuracy claims are tightened only after independent validation.
