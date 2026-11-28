# #504: the Negative Binomial and Geometric mass functions as pairs of Doubles — preregistration

Registered 2026-11-26 (Europe/Rome), before the candidate exists in VBA. The
100 holdout points are committed in
`../excel_function_map/negative_binomial_pair_holdout_grid.csv` with
references only. No file under `src/` changes in the commit that carries this
file.

## The defect, as measured

The holdout of #496 carried 27 Negative Binomial points that its rule did not
change, to learn whether that mass function is inferior to Excel too. It is:
8 of the 27 are less accurate than `NEGBINOM.DIST`, 3 of them above 1E-13,
the worst 6.1E-13 at `K_STATS_NegativeBinomial_PMF(21622, 150000, 0.9)`
(Excel 5.5E-13 there; at `(12200, 150000, 0.9)` the library is at 5.9E-13 and
Excel at 2.8E-14).

The Geometric mass function has not been compared with Excel away from the
few points of the floor grid. This registration measures it.

## The mechanism

**Negative Binomial.** The library forms the mass as Loader does:

    dnbinom(x; r, p) = dbinom(r; r + x, p) * r / (r + x)

`PROB_DS_TryNegativeBinomialLogMass` takes the Binomial logarithm from
`PROB_DS_TryBinomialLogMass` in one Double and adds `Log(r) - Log(r + x)`.
The Binomial logarithm carries what #498 found: two deviances in one Double
each, and means rounded before the deviance sees them.

**Geometric.** `PROB_DS_TryGeometricLogMass` returns
`Log(p) + k * Log1p(-p)`. The product is rounded at its own size, which is
the size of the logarithm of the mass: at a mass of 1E-200 one rounding of
the product is 5E-14 of the result.

## The rule

**R1. The Negative Binomial mass.** A new procedure takes the Binomial
logarithm as a pair from `PROB_TryBinomialLogPair` (#496, unchanged), adds
`Log(r / (r + x))` as a pair, and returns `Exp(high) * (1 + low)`.
`PROB_DS_TryNegativeBinomialPMF` takes it where `x > 0`.

**R2. The Geometric mass, conditionally.** A new procedure forms `1 - p`
exactly as a pair, `Log(1 - p)` as a pair, the product by `k` as a pair,
adds `Log(p)`, and returns `Exp(high) * (1 + low)`.
`PROB_DS_TryGeometricPMF` takes it where `k > 0`.

R2 enters the candidate **if and only if at least one `ge` point of the
holdout is inferior to Excel in the export of before** (Excel's relative
error smaller and the library's above 1E-14). If none is, R2 is not written,
the Geometric mass keeps its code, and rule N5 requires that no `ge`
observation changes. `nb_study.py` reads the condition from the frozen
verdicts of before (`geometric_rule`); no judgement enters it.

**R3. Where the rule does not apply.** At `x = 0` and `k = 0`, and where a
pair procedure returns FALSE (an argument outside 1E-100 to 1E+100), the
route is that of today.

## What calls the changed procedures

Found by scanning every procedure of every module for the name:

| Procedure | Called by |
| --- | --- |
| `PROB_DS_TryNegativeBinomialPMF` | `K_STATS_NegativeBinomial_PMF` only |
| `PROB_DS_TryGeometricPMF` | `K_STATS_Geometric_PMF` only |

`PROB_DS_TryNegativeBinomialLogMass` and `PROB_DS_TryGeometricLogMass` keep
their code, so the two LogPMF functions keep their results. The cumulative,
survival and inverse functions of both distributions call neither changed
procedure. `PROB_TryBinomialLogPair` is called and not changed, so the
Binomial and Hypergeometric mass functions keep their results.

## Development evidence

`nb_model.py` evaluates the rule in plain binary64. Without the rule it
stands on the ports of the source in
`../beta_balanced_survival/beta_route_model.py`. It is not bit-faithful: at
the worst point of the issue it returns 7.3E-13 where Excel measured
6.1E-13, and at `(2096, 40, 0.3)` 1.6E-13 against 5.1E-14. Its figures bound
the rule without predicting observations.

The 27 Negative Binomial points of the holdout of #496 are development
evidence: they were seen in Excel before this rule was written. Under the
rule the model puts all eight inferior ones within 1.7E-15.

On random arguments with a mass above 1E-290, in the model: Negative
Binomial, 589 points, today worst 6.3E-13 with 33 above 1E-13, under the rule
worst 1.5E-15; Geometric, 963 points, today worst 1.3E-13 with 3 above
1E-13, under the rule worst 6.9E-16.

The model was run on the holdout points before this registration
(`nb_study.py --model`): Negative Binomial, today worst 6.2E-13 with 12 of
52 above 1E-13, under the rule worst 1.4E-15; Geometric, today worst 1.2E-13
with 2 of 48 above 1E-13, under the rule worst 5.0E-16. The points are not
untouched for the model. They are untouched for Excel.

## The rows

`nb_study.py --build` writes 100 points, 100 rows, in the format of the Excel
floor grid. The points are where the logarithm of the mass has fallen from
its mode by 0.5, 20, 150, 400 and 600, on each side where the function
reaches it; for the Geometric mass, whose mode is at zero, at that count and
at the next. Points whose value is below 1E-290 are left out, and so is the
count zero.

| Points | Class | What they hold |
| ---: | --- | --- |
| 52 | `nb` | `NegativeBinomial_PMF` at six `(r, p)`: r from 3 to 900000, p from 0.02 to 0.97 |
| 48 | `ge` | `Geometric_PMF` at five p from 0.00001 to 0.93 |

None of the six `(r, p)` is one of the three of the holdout of #496, and no
point is in an earlier holdout. References are 50-digit values from the
definitions. Excel's side of a `ge` row is `NEGBINOM.DIST(k; 1; p; FALSE)`.

Before the candidate is written, the grid is exported once from the source of
today. `nb_study.py --freeze-before` keeps those verdicts, and those of the
floor grid of #493 and of the holdouts of #175, #498, #501 and #496 as
committed, in `nb_before.json`.

## Acceptance rules

`nb_study.py --check` implements them on the exported grids. The judged
points are the `nb` points, and the `ge` points if R2 applies.

| Id | Rule |
| --- | --- |
| N1 | Every judged point is within **1E-14** of the reference |
| N2 | No judged point is inferior to Excel; every row holds an observation and no library value is an error |
| N3 | Each of the eight points of the issue, rows of the holdout of #496, is within **1E-14** |
| N4 | In the floor grid of #493 and in the holdouts of #175, #498, #501 and #496, no point is newly inferior to Excel with a library error above **3E-14**, and no more points are inferior than before |
| N5 | If R2 does not apply, no `ge` observation changes |

Also required in the same wave: the full Excel regression, and the strict
accuracy gate on the main grid and the consolidated holdout with no contract
changed.

The wave exports no frozen study grid.

## Predictions

1. N1 to N5 hold.
2. The worst `nb` error is below 1E-14; before the rule it is above 3E-13.
3. Before the rule at least 8 of the 52 `nb` points are inferior to Excel;
   after it, none.
4. The eight points of the issue are within 3E-15.
5. R2 applies: at least one `ge` point is inferior to Excel before the rule.
   This is the least certain of the five: the model puts only 2 of the 48
   above 1E-13, and Excel's own error there is not known.
6. Outside the two mass functions no observation of any exported grid
   changes.

## After acceptance

The Negative Binomial points of the holdout of #496 are no longer inferior;
`floor_dispositions.json` holds no disposition for them and does not change
unless a floor-grid point changes verdict.

## Rejection policy

A failure of N1 to N5, of the regression or of the gate rejects the rule. The
bounds are not retuned and the result is recorded.

## What this does not address

- The counts `x = 0` and `k = 0`.
- `K_STATS_NegativeBinomial_LogPMF` and `K_STATS_Geometric_LogPMF`.
- The cumulative, survival and inverse functions of both distributions
  (#495 holds the cumulative and survival functions at large shapes).
- Masses below 1E-290.
- The cost per call of the pair route, which is measured before the release.
