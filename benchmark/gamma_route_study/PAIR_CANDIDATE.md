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
