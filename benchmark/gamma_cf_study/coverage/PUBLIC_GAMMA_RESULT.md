# Public Gamma P/Q integration and fresh validation

## Public repair

Implementation `1440b180036091c7c7635eb87e155a2762bb2a30` routes the public
regularized Gamma P/Q adapters through the frozen scaled double-double kernel.
The production arithmetic is extracted from private port `b8efaf3`, preserving
its statements, forced Double stores, constants and dispatch. Only private
names/types and integration comments change. Public signatures and status/error
contracts remain unchanged. The legacy series/continued-fraction helpers stay
available with their original arithmetic.

The first public Excel execution passed **13,182 / 13,182 assertions** with
**zero failures**, including all 41 previously failing public labels, Gamma
and chi-square CDF/SF/inverse callers, and natural-cap/recovery tests. Historical
B6 comparison controls still execute the exact old adapter expressions; the
public defect regressions execute the new public routes. No public failure is
skipped, relabeled as expected, or given a looser tolerance.

| Evidence | Identity |
| --- | --- |
| First public run | [42523118255](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/actions/runs/42523118255) |
| Tested PR merge | `9278ce3b2ffaf6786bba42a5876066f5539d9828` |
| Identical implementation/merge tree | `3239ba269d3e4d03905d2aac4960d47212010889` |
| Artifact | `13634552161` |
| ZIP SHA-256 | `b8e8f2eb2261f209fa649ea960decca104f53deea0691dcf8b6dd755694a9487` |

`excel_public_gamma_artifact.json` retains every original ZIP member and its
hash, plus the raw tested merge commit. Fixtures verify its Git identity,
tree equality with the reachable implementation commit, all seven imported
source hashes, counters, stages and original log hash. The recorded candidate
SHA is not rewritten. Import, execution-backed compile, regression and cleanup
passed on Excel 16.0 build 20430 / 64-bit. No numerical grid was exported.

## Separately frozen fresh public cohort

Protocol `5c23f636337a315a5e13170cb53aacafef0ae707` fixes the unchanged public
candidate, 206 forward points on 14 shapes and 24 inverse references on six
of those shapes before evaluation. The shapes are disjoint from the explicitly
audited 99-shape inventory across 46 retained JSON files; this is a finite
inventory claim, not universal independence from every conceivable past input.

Direct P and Q references use 160/320-digit mpmath evaluation and the previously
frozen direct-density quadrature fallback where needed. Normal-tail thresholds
are P `3e-15`, Q `5e-15`; subnormal/zero outputs require exactly rounded
reference bits. Inverse references use direct smaller-tail log-coordinate
solves with precision stability; the public caller tolerance is `2e-14`.

The initial `public_fresh_screen.json` remains **REJECT**. One Q comparison at
`a=95.0625, x=1024` exposed double rounding in conversion of the mpmath
reference to binary64. The exact reference is approximately
`4235992831782713.33053521835056825067` minimum-subnormal units, whose nearest
binary64 is `0x0.f0c9cf0026f39p-1022`; the naive conversion returned the next
subnormal instead. Exact integer quotient/remainder conversion and an
independent direct-density quadrature check confirm the candidate's original
output. See [PUBLIC_REFERENCE_COMPLETION.md](PUBLIC_REFERENCE_COMPLETION.md).

Completion `c8b145ebf70a94652b97e8a49a9edf68a0f50568` froze that reference
conversion separately. It changes no candidate arithmetic, coordinates or
tolerance and does not overwrite the rejected result. The completed forward
screen has **zero accuracy and monotonicity failures**; the worst normal-tail
relative errors are approximately `1.0832e-16` for P and `9.4054e-17` for Q.
All 24 inverse references are stable. This is one completed cohort, not two
independent successful cohorts.

The host test commit `b005341702f5fb70323453bded9b2109250ee541` adds 1,696
assertions: eight per forward point (public P/Q success and bitwise values,
Gamma CDF/SF at scale two, chi-square CDF/SF at doubled shape/argument) and two
per inverse reference (Gamma and chi-square inverse). The policy requires
**14,878 assertions**. Actual Excel execution passed **14,878 / 14,878**,
with **zero failures**. All 1,696 fresh public assertions passed, including
the 48 inverse caller checks. The frozen model-screen result is left unchanged;
this is the separately recorded host acceptance of that finite cohort.

| Fresh host evidence | Identity |
| --- | --- |
| Excel run | [42524515131](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/actions/runs/42524515131) |
| Tested PR merge | `a0202380a6b6b07e102585db455ebb769b6fa0b0` |
| Identical host-test/merge tree | `2566038753f7bf0e58cbbea03b77331ce354d883` |
| Artifact | `13635460900` |
| ZIP SHA-256 | `c77ad4c66cf7f03c1cd4110c359615660b807f526a10efc86fa78c62a19ebe9b` |

`excel_public_fresh_artifact.json` retains the original members and raw merge
commit with the same identity/source checks as the first public artifact.
Import, execution-backed compile, regression and cleanup all passed on
Excel 16.0 build 20430 / 64-bit. The recorded session lasted approximately
8m 55s; the workflow took 10m 43s. This is observed suite time, not a general
performance guarantee. Both grid entries are `exported: false`.

## Scope and remaining acceptance

All 76 local tooling commands pass on the fresh host-test tree; all 18 strict
static checks pass. The focused #241 port carries the same production kernel
and the four original public regression groups, with 1,947 expected assertions.

Both #241 and #242 remain open and draft. These results do not refresh or
rebind the main/holdout grids, supply a global accumulated-error bound, prove
performance over the full input domain, or certify another Office environment.
Fresh exact-source Excel exports, assessment of their numerical contracts,
current CI and review remain required before integration/release acceptance.
On the tested host-test commit, hosted static checks, CodeQL, test selection,
all evaluation-tooling proofs and positive-ratio study proofs pass. The live
Accuracy Gate fails on stale main evidence and on the older 965-assertion
canonical export record versus the new policy; its README assurance validation
also fails on that obsolete record. These are not reported as green CI.
