# B4: residual-preserving pair recurrence development

B1/B2/B3 remain rejected. This new candidate removes exponent normalization
from each recurrence operation, reuses the audited two-Double operation graphs,
and retains scaled high-precision prefactor and final P/Q composition. The
prefactor uses the existing bounded one-entry LogGamma cache. It is not a pure
binary64 candidate and no microsecond speedup is assumed.

The prototype uses a positive lower series with a geometric remainder screen
4*remainder < total*2^-100. Upper Lentz iterations stop at |delta|<2^-100 and
are reevaluated backwards. Unscaled values outside 2^-450..2^450 decline;
the first Lentz C is 2^400. Caps are 100000. The provisional envelope is
1e-8 <= A <= 10000 and 1e-12 <= X <= 20000. These are development limits, not
a frozen dispatch rule. A future public integration must retain fallback and
independently test its envelope boundaries as well as algorithm seams.

## Revealed development results

- All 17 retained failing B1/B3 pairs become monotone.
- 43 windows of 33 neighboring Doubles (1419 evaluations) have no decrease in
  P or increase in Q.
- 286 of the 392 revealed fitting inputs are answered; 106 are declined.
- Worst scored relative errors: P 1.069e-16, Q 9.965e-17, below the unchanged
  governing contracts. References below 1e-300 follow the existing separate
  underflow policy; no global accuracy certificate is claimed.

`screen_pair_candidate.py` reproduces `pair_candidate_results.json`. No
holdout input/reference is read. These are Python model results only.

## Excel wave specified before observing the host

`prepare_pair_host.py` binds the LF-normalized production source hash, appends
study procedures only in a temporary Excel workbook and emits all unique
revealed fit inputs plus the 43 neighborhoods and retained pairs. Every input
must match predicted P/Q bits or the predicted decline status. The host
records its commit and Excel version/build; incomplete runs, runtime errors,
cleanup failures and any fidelity mismatch fail the wave.

After fidelity, compare the eight existing cost points, three rounds, baseline
then candidate in rounds 1/3 and reversed in round 2. One warmup precedes four
timed paired P/Q evaluations. Report every point, median warm pair cost and
speedup; do not divide by two or equate it to public scalar-call latency.
The earlier routing study's 100-us budget remains a target, not an inherited
certification for this different candidate. A cost above 100 us per pair
cannot substantiate meeting that conservative target. Cold performance and
inverse callers require subsequent measurements. No optimization is accepted
from this initial fidelity/cost wave alone.

A2 remains 10000 as the maximum permitted study ceiling; the cost wave may
show this candidate is too expensive even there. Only after arithmetic,
Excel fidelity, costs and new boundary probes are satisfactory can a revised
routing rule be frozen and the untouched 117-point holdout evaluated.

## First host attempt and normalization correction

Run 43055630408 at c27e2f0b417cf84d3925d1d6757da7a421a08f29 failed
before recording any completed input or timing. The saved report records RPC
failure and cleanup errors; it is retained in results/run_43055630408_failure.json.
Artifact 13873093355 ZIP SHA-256 was verified before extraction:
4efebe22fe02517f718bf054897a15e55074a1741fb429a05a4f5f0eba303155.

The debugger screenshot showed the scaled arithmetic input guard. Code review
found that B4_Scaled called Canonicalize on an unscaled pair: Canonicalize only
repairs an already normalized mantissa, and does not perform frexp range
normalization. This is a VBA port defect; Python S._normalize already performs
that step. The corrected port splits the high part with Frexp, shifts the low
part by the inverse power of two, then canonicalizes and validates.

The next host wave first checks 11 conversion cases (positive/negative pairs
at exponents -400,-1,0,1,400 plus zero), including separate high/low roundtrip
preservation. Missing/failed checks invalidate the report. Probe error handlers
return explicit ERROR diagnostics to fail the harness, and the current input
ID is retained if a runtime error interrupts a call. No numerical tolerance,
model output or domain boundary is changed; full fidelity/cost must be rerun.

## Second host attempt: exact-input transport correction

Run 43062050989 at ed44d02807d061fbe3bd062c7c13154384ec0dc2 passed
all 11 normalization checks and the first eight fidelity inputs. The ninth
input echo differed: requested X bits 3FF0000002AF31D2, observed
3FF0000002AF31D1 (the immediately preceding Double). Its output therefore
repeats the preceding row, and is not a same-input numerical counterexample.
The report retains nine rows, zero timings, `Input echo mismatch`, and no
cleanup error. Artifact 13875295233 ZIP SHA-256 verified before extraction:
3e326af5d0ebfecb7c1f0f0d90dd912051ae169fd2bfe86b558612d94a767465.

The fidelity loop previously passed decimal values parsed by Windows
PowerShell ConvertFrom-Json. It now reconstructs A and X from the existing
16-hex-digit IEEE words in each case ID, using UInt64/BitConverter, avoiding
decimal parsing of neighboring inputs. The unchanged Excel echo check still
rejects any transport discrepancy. Model predictions, tolerances, kernel,
normalization check and input matrix are unchanged. Repeat the complete wave;
partial runs are not combined and no timing claim is available.
