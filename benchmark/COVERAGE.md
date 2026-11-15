# Main-grid row disposition

`check_grid_coverage.py` prevents a measured main-grid row from being silently
ignored because no accuracy contract matches its function and regime.

The canonical row identity is the function, four argument values as IEEE-754
binary64 bit patterns, regime, and evidence set. Decimal spellings therefore do
not create false drift, while a one-ulp argument change remains distinct.

## Modes

- `--mode strict` is the release authority. Every main-grid row must match an
  active or characterization contract, or an explicit final entry in
  `accuracy_row_exemptions.json`. Missing, stale, duplicate, conflicting, or
  malformed dispositions fail.
- `--mode transition` was temporary for issue #22 and is retired. It permitted
  only the unresolved subset of `coverage_debt_v1_0_0.json`, the frozen 36-row
  audit baseline from `bde92dd7037e4fde05e620745a1c54b0cbc3a261`. The #22
  closure deleted that fingerprint with the governing contracts in place, so
  the gate selects strict mode. The mode stays in the tool for its fixtures;
  no committed fingerprint exists for it to read.

## How the 36 baseline rows were resolved

| Rows | Disposition |
| ---: | --- |
| 9 | Student-t inverse, `envelope_domain`: claimed by the 5E-15 contract of #34 |
| 9 | Chi-square inverse, `envelope_domain`: claimed by `ChiSquare_InverseCumulative.envelope_domain.quantile_rel` |
| 18 | F inverse, `envelope_domain`: claimed by `F_InverseCumulative.envelope_domain.quantile_rel` |

The two new contracts carry the inverse guarantee of the accuracy policy,
1E-12, as their threshold: a policy floor, not a value fitted to the rows.
The measured worst is 1.09E-15 for Chi-square and 2.15E-15 for F.

With the same change, in the main grid: three `validated` F rows that
duplicated `envelope_domain` rows at (1E6, 3) were removed; five `validated`
F rows with a degree of freedom of 1E6 moved to `envelope_domain`, where
their domain says they belong; and the eight Student-t rows at p = 0.5 moved
to `exact_median`, scored by absolute error against a threshold of zero and
held by exact-zero assertions of the Excel regression. In the consolidated
holdout, eight F rows at (1.5, 2E5) and (5E5, 4) moved to `envelope_domain`.

`coverage_summary.md` is generated from the current grid and registry. The
current remaining-debt count is never copied from the frozen baseline metadata.
