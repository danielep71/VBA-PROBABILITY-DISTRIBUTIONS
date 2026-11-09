# Public Gamma integration candidate: freeze before Excel evaluation

2026-11-08. #24/#26; #241/#242 stay open and draft.

Public regularized Gamma P/Q now call the exact arithmetic dependency closure
of WG_Pair from frozen Excel port b8efaf3. The 54 private procedures are copied
with private-name qualification; arithmetic statements, stores and constants
are unchanged. New private bit types remain in SPECIALFUNCS. Public signatures
and the six-module inventory are unchanged. The adapter clears output/status,
rejects invalid returned probabilities and preserves established natural-cap
diagnostic wording. Existing SeriesP/ContinuedFractionQ remain as historical
paths; no new public API or numerical tolerance is introduced.

The two private rejected-routing controls are explicitly switched to copied
B6 regularized wrappers. Their original arithmetic remains executable against
the unchanged legacy series/fraction. All other 10848 existing assertions are
preserved. In particular the 41 old public failures are still live assertions,
not deleted, suppressed or translated into expected failures.

Add 2334 assertions: 290 revealed coordinate pairs, each tested through public
P/Q success and bit fidelity plus Gamma CDF/SF (scale two) and Chi-square CDF/SF,
and seven invalid coordinate pairs through both kernels. Expected aggregate
policy: 13182. Exact expected bits come from the previously frozen private
port, not this newly integrated source. This is composition/fidelity evidence,
not a fresh independent public holdout.

Freeze this source before executing Excel. Retain every outcome, including any
new caller/inverse failure. A passing aggregate is necessary but insufficient:
then freeze a fresh public validation cohort, verify both tails independently,
check boundary/underflow/resource behavior, obtain current exact-source exports
and pass final integration checks and review. Keep prior evidence attached to
its original candidate; stale exports must continue to fail the live gate.

Historical isolation tests use a strict inverse of the source generator: it
first verifies the exact new block and wrappers, then restores the frozen old
bodies. An independent live-source fixture verifies this operation, rejected
mutations, caller checks, API inventory and unchanged remaining source. The
inverse is never used to certify current public outputs from old artifacts.
