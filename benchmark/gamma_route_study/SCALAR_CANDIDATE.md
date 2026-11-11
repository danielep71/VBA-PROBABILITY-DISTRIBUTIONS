# B5: scalar pairs in the B4 recurrences (#307)

## Frozen development question (2026-11-11, before the host run)

Does removing repeated Variant-array construction and extraction materially
reduce B4's warm paired P/Q cost without changing a single returned bit?

B4 host run 43064890043 at `6a41de8f4e9951e273b31296094425bd8b157b22`
passed 1,371 exact-input checks but cost 6–179 ms per pair. The scalar experiment
starts at that commit. It is development work using revealed data only.

## Candidate and invariants

`scalar_candidate_vba.txt` contains B5. Its recurrence state and all arithmetic
helper inputs/outputs are scalar Double hi/lo pairs. No Array call or Variant
value occurs inside any recurrence loop or scalar arithmetic helper. Subroutine
inputs are ByVal snapshots; output aliasing therefore preserves the old operands.
Explicit Double stores preserve PV2's operation boundaries, including both
quotient-refinement steps and every multiplication cross term.

The prefactor, logarithm/exponential envelope, final scaled conversion and
complement are the unchanged B4 paths. Only the final recurrence pair is boxed
once for the existing, validated B4 normalizer. Range checks, the 100,000-step
caps, 2^-100 stopping rules, domain and A2=10,000 remain unchanged. No direct
public dispatcher, source module, reference grid or holdout is modified.

## Predictions fixed before host evaluation

- Exactly the existing 1,371 revealed inputs: 1,265 answers and 106 declines.
- Every B5 answer equals the unchanged B4 Python model bit for bit (both P and Q).
- The 576 scalar arithmetic comparisons (12 x 12 operand pairs x four operations)
  must match both hi and lo words of the audited PV2 operators. Operands include
  both signs, cancellation, low residuals and exponents from -100 to 100.
- The existing 11 scaled-normalization checks must still pass.
- Expect less allocation/extraction overhead; no numerical improvement is claimed.

## Host protocol and decision

The dedicated `Gamma Scalar Pair Fidelity` workflow uses a temporary workbook
with both unmodified B4 and B5 appended to SPECIALFUNCS. Predictions are generated
by the unchanged B4 Python model. The runner reconstructs fidelity inputs from
IEEE words, checks returned input words, fails on incomplete evidence and records
source SHA, workbook-host version, cleanup status and instrumentation hashes.

After fidelity passes, measure the same eight (A,X) points as B4, three rounds,
four individually timed calls per mode per point. Each sample warms its own
mode once. Compare B4 and B5 **in the same run**, alternating order by round.
Both modes must match the predicted P/Q words on timing points. Total: 48 timing
rows. Report median microseconds per paired call, every point, including any
regression. This is not a comparison of different-run wall times.

A useful development signal requires at least 5% lower median cost at **both**
A=10,000 points. Separately flag any point with more than 5% extra cost. This is
an exploratory filter, not statistical or release acceptance. The conservative
100-us paired budget is reported separately and is not relaxed. A useful result
still needs independent timing confirmation and public-path validation before
promotion. Even exact fidelity on revealed inputs cannot close #307 or authorize
looking at the frozen holdout.

## Reproduce

```sh
python benchmark/gamma_route_study/test_pair_candidate.py
python benchmark/gamma_route_study/test_scalar_candidate.py
python benchmark/gamma_route_study/prepare_scalar_host.py
# Windows / Excel 365 64-bit, on the same clean source SHA:
powershell -NoProfile -ExecutionPolicy Bypass -File benchmark/gamma_route_study/Run-ScalarStudy.ps1
python benchmark/gamma_route_study/validate_scalar_host.py artifacts/gamma-scalar
```

The temporary workbook is not committed. The uploaded JSON/text evidence and its
exact source SHA are retained; no timing is inferred from local Python tests.
