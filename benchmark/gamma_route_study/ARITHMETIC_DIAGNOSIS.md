# #307: isolate the arithmetic error before another routing rule

A2 remains 10000. B1, B2 and B3 stay rejected. This diagnostic uses only
17 previously retained failing adjacent pairs (four B1 and thirteen B3).
It does not evaluate or reveal any holdout point and is not a new candidate.

For each exact binary64 input, direct P, direct Q and the common prefactor
x^a exp(-x)/Gamma(a) are independently evaluated with mpmath at 80 and 120
decimal digits. All three agree to 1e-65 relative. Separate interventions
replace either the prefactor with its rounded reference, or the normalized
recurrence with its rounded reference, preserving the final binary64 product
and the respective B1/B3 dispatch logic. The latter can change routing when
a substituted value crosses 0.5: this is part of the observed composition,
not a controlled estimate of an isolated runtime component's error.

| Intervention | Remaining B1 failures (4) | Remaining B3 failures (13) |
|---|---:|---:|
| Original implementation | 4 | 13 |
| Rounded reference prefactor | 2 | 4 |
| Rounded reference recurrence | 2 | 8 |
| Direct reference P/Q pair | 0 | 0 |

Neither isolated replacement repairs every known failure. This rules out
claiming that a prefactor-only or recurrence-only replacement with these
specific rounding boundaries suffices. It does not prove that no other
single-component change can work. The reference-pair control is monotone by
construction and is not a proposed implementation.

Next development must consider the recurrence, prefactor and final composition
together, carrying residuals until final rounding. A candidate must first
replay all retained failures and wider revealed boundary neighborhoods, then
pass a new Excel fidelity/cost wave before another routing rule is frozen.
Do not expose the untouched holdout to select arithmetic changes.

Reproduce with `python benchmark/gamma_route_study/arithmetic_diagnosis.py`.
The fit fixture replays the JSON report, including exact inputs/outputs and
source hashes. The report is diagnosis only; no VBA, public dispatch, accuracy
contract, threshold, or expected regression result changes in this work.
