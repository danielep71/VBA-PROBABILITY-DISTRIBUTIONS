# #496: the Hypergeometric mass function as pairs of Doubles — preregistration

Registered 2026-11-26 (Europe/Rome), before the candidate exists in VBA. The
96 holdout points are committed in
`../excel_function_map/hypergeometric_pair_holdout_grid.csv` with references
only. No file under `src/` changes in the commit that carries this file.

## The defect, as measured

The Excel floor comparison of #493 found
`K_STATS_Hypergeometric_PMF(100, 200, 2000, 50000)` at 1.2E-13 relative,
where `HYPGEOM.DIST` is within 1.9E-14. Since #175 changed the far branch of
the deviance the same point measures 3.5E-14 (Excel run 45848441378): still
less accurate than Excel, no longer above 1E-13.

One point does not describe the function. In the model, on 1181 random
arguments up to 25 standard deviations from the mean with a mass above
1E-290, the worst error of today is 1.1E-12 and 176 are above 1E-13. Whether
Excel is closer there is what the export of before measures.

## The mechanism

The library forms the mass as Loader does:

    dhyper(k; n, K, N) = dbinom(k; K, p) * dbinom(n - k; N - K, p) / dbinom(n; N, p),   p = n / N

`PROB_DS_TryHypergeometricLogMass` takes each Binomial logarithm from
`PROB_DS_TryBinomialLogMass` in one Double and returns `B1 + B2 - B3`. Each
term carries what #498 found in the Binomial mass function: two deviances in
one Double each, and means `K p` and `K (1 - p)` rounded before the deviance
sees them. Here `p = n / N` is itself rounded first, and enters all three
terms.

## The rule

**R1. The Binomial logarithm as a pair.** The part of
`PROB_TryBinomialMassPair` that forms the logarithm becomes a private
procedure of its own, which takes the probability as a pair and returns the
logarithm as a pair. `PROB_TryBinomialMassPair` calls it with its Double
probability; its arithmetic does not change, and the holdout of #498 must
read as before.

**R2. The Hypergeometric mass.** A new procedure forms `p = n / N` as a
pair, takes the three logarithms from R1, adds the first two and subtracts
the third in pairs, and returns `Exp(high) * (1 + low)`.
`PROB_DS_TryHypergeometricPMF` takes it where each of the three Binomial
terms is away from its edges: `0 < k < K`, `0 < n - k < N - K` and
`0 < n < N`.

**R3. Where the rule does not apply.** On the edges of a Binomial term, and
where a pair procedure returns FALSE (an argument outside 1E-100 to 1E+100),
the route is that of today.

Nothing else changes. `PROB_DS_TryHypergeometricLogMass` keeps its code, so
`K_STATS_Hypergeometric_LogPMF` and the cumulative and survival functions,
whatever they call, keep their results. `PROB_DS_TryBinomialLogMass` keeps
its code, so the Negative Binomial functions keep theirs.

## Development evidence

`hyper_model.py` evaluates the rule in plain binary64. Without the rule it
stands on the ports of the source in
`../beta_balanced_survival/beta_route_model.py`, and returns the 3.49E-14
Excel measured at the point of the issue. That is one point of agreement; it
is not bit-faithful everywhere, and its figures bound the rule without
predicting observations.

On the 1181 random arguments above: under the rule the worst is 3.5E-15 and
none is above 1E-14.

The model was run on the holdout points before this registration
(`hyper_study.py --model`): under the rule the worst is 3.2E-15, at
`Hypergeometric_PMF(15609, 70001, 250000, 999983)`; today 7 of the 69 are
above 1E-13, the worst 7.7E-13. The points are not untouched for the model.
They are untouched for Excel.

A first draft of the holdout used only parameter sets whose ratio `n / N` is
short in decimal; the model found today's worst there at 1.2E-13. Three sets
with a ratio that is not short were added before this registration, and the
worst rose to 7.7E-13. All nine sets are in the grid.

## The rows

`hyper_study.py --build` writes 96 points, 96 rows, in the format of the
Excel floor grid. For each parameter set the points are where the logarithm
of the mass has fallen from its mode by 0.5, 20, 150, 400 and 600, on each
side where the function reaches it. Points whose value is below 1E-290 are
left out.

| Points | Class | What they hold |
| ---: | --- | --- |
| 69 | `hg` | `Hypergeometric_PMF` at nine `(n, K, N)`, N from 400 to 1000000, every term away from its edges |
| 27 | `nb` | `NegativeBinomial_PMF` at three `(s, p)`; the rule does not change it |

The `nb` points are there to measure, not to judge: the Negative Binomial
mass goes through the same Binomial logarithm in one Double, and whether it
is inferior to Excel is not known. No point is in an earlier holdout.
References are 50-digit values from the definitions.

Before the candidate is written, the grid is exported once from the source of
today. `hyper_study.py --freeze-before` keeps those verdicts, and those of
the floor grid of #493 and of the holdouts of #175, #498 and #501 as
committed, in `hyper_before.json`.

## Acceptance rules

`hyper_study.py --check` implements them on the exported grids.

| Id | Rule |
| --- | --- |
| M1 | Every `hg` point is within **1E-14** of the reference |
| M2 | No `hg` point is inferior to Excel; every row holds an observation and no library value is an error |
| M3 | The point of the issue, a row of the floor grid, is within **1E-14** |
| M4 | In the floor grid of #493 and in the holdouts of #175, #498 and #501, no point is newly inferior to Excel with a library error above **3E-14**, and no more points are inferior than before |

Also required in the same wave: the full Excel regression, and the strict
accuracy gate on the main grid and the consolidated holdout with no contract
changed. The holdout of #498 must show no changed observation: R1 moves code
and not arithmetic.

The wave exports no frozen study grid.

Reported, not gating: the `nb` points before and after, and how many are
inferior to Excel.

## Predictions

1. M1 to M4 hold.
2. The worst `hg` error is below 1E-14; before the rule it is above 3E-13.
3. Before the rule at least 10 of the 69 `hg` points are inferior to Excel;
   after it, none.
4. The point of the issue is within 3E-15.
5. No observation of the holdout of #498 changes.

No prediction is made on the `nb` points.

## After acceptance

The disposition of the point of the issue is removed from
`floor_dispositions.json`, with those of any other floor-grid point that
stops being inferior. If the `nb` points show the Negative Binomial mass
inferior to Excel above 1E-13, that is a new issue, not an extension of this
rule.

## Rejection policy

A failure of M1 to M4, of the regression or of the gate rejects the rule. The
bounds are not retuned and the result is recorded.

## What this does not address

- The edges of the Binomial terms: `k = 0`, `k = K`, `n - k = 0`,
  `n - k = N - K`.
- `K_STATS_Hypergeometric_LogPMF`, and the cumulative and survival functions.
- The Negative Binomial and Geometric mass functions.
- Masses below 1E-290.
