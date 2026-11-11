# #307 phase B: frozen public routing rule

Preregistered on 2026-11-11, before evaluating the routing holdout.
Base: `98ff56111c77c5631d560b1d9eb82e3a3710aa6a` (merged #312).
The user accepted A2 = 10,000; raising it requires a separate preregistered study.
The original HOLDOUT_DESIGN.md, its 117 points and rules H1-H6 are unchanged.
Pre-freeze disjointness: 17 shapes and their doublings absent from 600 tracked
files, as reported by check_disjoint.py on the base above.

## Ordered rule

Keep current invalid-input and X = 0 behavior. Only finite A > 0, X > 0,
A <= 10000 are eligible for the new fast path. Evaluate X < fl(A + 1)
with the same Double comparison established by the phase-A Excel run.

- Below that seam, call PROB_TryGammaB64SeriesP once. If it succeeds with a
  finite value in [0,1], P returns that value. Q returns 1 - P only if
  P <= 0.5; otherwise Q uses the existing double-double public route.
- At or above the seam, call PROB_TryGammaB64ContinuedFractionQ once. If it
  succeeds with a finite value in [0,0.5], Q returns it and P returns 1 - Q.
- Any decline, invalid output or excluded input goes to the existing public
  route, including its previously validated cap fallback. Clear a declined
  probe's diagnostic before this fallback. Do not change its error policy.
- Above A2 never attempt the new fast routines. Their phase-A implementation,
  exponent guards, factor floors, stopping rules and iteration limits remain
  unchanged. No contract or underflow convention changes.

The closed interval Q <= 0.5 is conservative: the phase-A upper route returns
Q below 0.5. No host result has been used to extend the permitted interval.

## Predictions fixed before evaluation

Eligible lower P and eligible upper P/Q will meet their governing contracts
at binary64 cost; lower Q with P > 0.5 and all shapes above A2 remain on the
existing route. H1-H5 will pass without relaxing any threshold. Every accepted
fast call, including adapter and caller overhead, will satisfy H6 (<200 us).
The phase-A worst eligible routine cost is 67 us; it is not a guarantee for
adapter overhead or future hosts. The eight #298 cost points and inverse
workload must be measured again; no speed claim is made for excluded shapes.

## Mechanical boundary probes

Materialize and retain an input/probe manifest before scoring or running the
host holdout; record exact hex inputs and the rule predicate for each route.
Do not choose points by looking at their errors. Deduplicate exact input pairs.

1. At each frozen holdout shape, predecessor/fl(A+1)/successor in X.
2. At each eligible holdout shape, the lower-Q decision transition: search
   positive binary64 X below fl(A+1) by integer-bit bisection for the first
   successful series result greater than 0.5. Retain its predecessor, itself
   and successor. This construction may inspect the fixed predicate only,
   not errors or timings; retain all construction diagnostics. If the
   transition cannot be bracketed, record that fact rather than invent one.
3. At predecessor(10000), 10000, successor(10000), use X = 10000 + k sqrt(10000)
   for every frozen HOLDOUT_DESIGN offset, and the three X seam probes for
   each of these shapes. Mark these as boundary probes, not independent shapes.
4. At the existing internal boundaries A=1, X=0, X=708 for A<=1, and the
   deviance exponent/factor-floor guards for A>1, retain adjacent representable
   triples. For A=1 use each SMALL_X from holdout_design.py at predecessor(1),
   1 and successor(1). For every eligible holdout shape include X=0 and its
   two neighbours (the negative neighbour is an invalid-input control). For
   every holdout shape below 1 use the X=708 triple. For every eligible
   holdout shape above 1 search all positive finite binary64 X: exponent-field
   transitions at 1 and 2001, and the factor-floor predicate Eh < -708,
   separately below and above X=A. Integer-bit bisection locates the first
   point of the new predicate state; retain it and its two neighbours.
   The factor predicate is inspected only inside the accepted exponent range.
   A missing bracket is reported explicitly. Include both accepted and
   declined sides; never filter failures. Invalid-input controls retain the
   established error contract and are not scored as valid H1 probabilities.

For X triples apply H3 and H5. For shape triples test the appropriate Gamma
shape monotonicity (P nonincreasing, Q nondecreasing at fixed X). Quantify any
boundary jump as the difference of the two output errors against independently
computed direct-tail references; its absolute value must not exceed the sum
of the governing absolute error allowances. This removes the genuine change
in the mathematical function from the discontinuity test, without waiving H1.

## Execution and rejection policy

Freeze this rule in Git before constructing the holdout probes. Implement the
public adapters and regenerate expected bit patterns only on already revealed
regression cases. Preserve private-kernel fidelity tests and all original
accuracy, seam, resource and error assertions. Commit the complete probe
manifest, exporter and analyzer before the first holdout host wave.

Compute P and Q independently at two precisions at exact binary64 inputs.
An unstable reference is unevaluated, never PASS. Preserve failed attempts.
Time each of the four public Gamma/chi-square calls separately with the #298
count-doubling method (at least 0.5 seconds per measurement); report all calls,
including slow fallback calls. H6 applies only to calls that return via the new
fast path, but fallback overhead remains visible. Execute the #298 cost and
inverse benchmark in the same run and on the same host.

Any H1-H6 failure rejects this frozen rule. Do not retune A2, predicates,
thresholds or probes against the revealed holdout. Retain the rejection and
use a new preregistered study for any subsequent candidate. #307 stays open
until acceptance, existing main/holdout/study gates, current Excel evidence,
README cost/limitations and review are complete.
