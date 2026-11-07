# #26: post-prefactor continued-fraction diagnosis

Baseline: release/1.0.0 at `ed4b138`, after #23. This is diagnosis, not closure.
No production source, observation or contract is changed by this study.

`run_isolation.py` reproduces **all 12 X=2 Excel observations exactly** using
source-derived Loader/Stirling helpers and the already validated VBA
64-bit-significand intermediate model. The exported ladder is from Excel run
42365758589. References and the finite continued fraction use 90 decimal digits. All twelve
references are checked at 60 versus 90 digits (relative difference <=1E-55)
and against independent direct quadrature (relative difference <=1E-75).

The multiplicative decomposition is exact to the recorded precision:

```
1 + total_error = (1 + prefactor_error) * (1 + finite_fraction_truncation)
                * (1 + recurrence_rounding) * (1 + final_multiply_rounding)
```

| Component on the X=2 ladder | Magnitude |
|---|---:|
| Finite-fraction truncation at the production stop | 3.70E-14 to 4.26E-14 |
| Prefactor relative error | at most 1.63E-15 |
| Recurrence rounding | at most 2.41E-15 |
| Final multiplication rounding | at most 6.52E-17 |

The dominant cause is stopping at `Abs(Del - 1) <= 3E-14`. Most rows stop at
iteration 38; shape 0.5 stops at 36. A study-only machine-epsilon stop takes
46–49 iterations and puts all twelve X=2 results below 5E-15, worst 4.44E-15.
This is **not sufficient evidence to ship that one-line change**.

## Expanded fitting checks

The exploratory fitting set adds the exact seam `X=A+1`, its next binary64
neighbour above, and `(A+1)` times 1.25, 2, 4 and 16. Shapes are the twelve
ladder shapes plus 0.6, 1, 2, 5, 10, 100, 1000, 10000 and 1000000.
Underflowing reference probabilities are reported outside the normal-output
comparison, not as zero error.

At shape 1E-5, X=1.00001, merely tightening the stop leaves **8.19E-15** error.
Of the 72 fitting probes with shape at most 0.5, seven still breach 5E-15
with the tighter stop. The backward candidate breaches none; its worst relative
error is 4.72E-15. These are fitting results, not a frozen new domain contract.

Re-evaluating the same finite fraction backward avoids the accumulated Lentz
product rounding in the explored small-shape fitting set. It is a candidate,
not production validation. At the time of the original diagnosis, no untouched
holdout had been evaluated. The subsequent frozen model checkpoint is below.

The broad exploratory tail probes also expose pre-existing prefactor-dominated
errors (for example A=10000, X=12501.25, about 9.38E-12). Those are not removed by
changing the continued fraction. Do not describe the small-shape repair as a
universal 5E-15 Gamma survival guarantee. The existing large-shape contracts
have narrower near-body domains and must not be extended to these tail points.

## Reproduction

```
python benchmark/gamma_cf_study/run_isolation.py
```

`isolation.json` binds the report to the baseline source, model and exported
ladder. Model mismatches, incomplete decomposition and iteration exhaustion
fail explicitly. The old algorithm remains available as the comparator.

## B1 independent model checkpoint

`CANDIDATE_PREREGISTRATION.md` and `candidate_protocol.json` were published in
commit `f95d3e04f9ce19749ab479e061966668c3d1901b` before evaluating the new
coordinates. B1 is the existing machine-epsilon stop plus backward reevaluation
at exactly that depth; no further tuning was performed.

| Set | Points | B1 failures against 5E-15 | Maximum relative error |
|---|---:|---:|---:|
| Previously inspected fitting | 84 | 0 | 4.713E-15 |
| New independent coordinates | 84 | 0 | 4.490E-15 |

All 168 references pass 60/90-digit stability and independent direct-quadrature
checks. The protocol and exact binary64 coordinates are immutable inputs to the
replay. Missing, unstable, non-finite, nonpositive and unexpectedly subnormal
references are rejected, as are mismatched frozen candidate bytes. The report
records all comparator results, iteration counts and source/input hashes.

Run `python benchmark/gamma_cf_study/check_candidate.py` to reproduce the result.
The hosted tooling shim runs this check. Report reproducibility is distinct
from numerical acceptance: a future honestly recorded rejection would not be
silently converted into a successful candidate.

**Status: model screen passed; Excel not run for B1.** The independent margin
is narrow. No new production route, public-function accuracy claim or issue
closure follows. The next stage requires its own production protocol covering
the remaining domains and callers, permanent VBA regressions and a fresh
exact-source Excel wave. The historical #23 ladder and global holdout remain
unchanged. These revealed coordinates cannot be reused as an untouched set
when tuning another candidate.
