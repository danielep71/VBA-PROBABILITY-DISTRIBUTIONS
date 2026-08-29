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
- `--mode transition` is temporary for issue #22. It permits only the unresolved
  subset of `coverage_debt_v1_0_0.json`, the frozen 36-row audit baseline from
  `bde92dd7037e4fde05e620745a1c54b0cbc3a261`. It labels those rows `KNOWN
  COVERAGE DEBT`, rejects any new or changed unclaimed identity, and rejects a
  fingerprinted row that is deleted instead of receiving a real disposition.

The fingerprint records its owner, issue, v1.0.0 expiry, production-module and
grid identity, initial count, and every canonical row. It is an anti-regression
fence, not an accuracy contract or exemption. The #22 closure commit must delete
it; the gate then selects strict mode automatically.

`coverage_summary.md` is generated from the current grid and registry. The
current remaining-debt count is never copied from the frozen baseline metadata.
