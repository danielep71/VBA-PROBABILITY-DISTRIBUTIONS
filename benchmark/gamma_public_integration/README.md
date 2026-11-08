# Public Gamma integration (#26; PRs #241 and #242)

**Latest evidence:** [the fresh main/holdout export](EXPORT_RESULT.md) at
`35e7381` passes 1,947/1,947 Excel assertions, all 166 active main contracts
and all 83 independent-holdout contracts. It supersedes the stale-grid
limitation of the regression-only checkpoints below for this focused branch.

The regularized incomplete-Gamma P/Q entry points now use the scaled
double-double kernel developed and measured in #242. Gamma and chi-square
CDF/SF and inverse callers use it through their existing public routes.
Public signatures, parameterization, error returns and diagnostic ownership
are unchanged. The 100,000-iteration cap still fails explicitly, and a later
valid call must recover. The legacy direct series and continued-fraction
helpers remain available with their original arithmetic.

The private arithmetic and public adapters are byte-identical to implementation
commit `1440b180036091c7c7635eb87e155a2762bb2a30`. Its paired Excel run
[42523118255](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/actions/runs/42523118255)
executed **13,182 assertions: 13,182 PASS, 0 FAIL**, including all 41 previously
failing public regressions. Import, execution-backed compile, regression and
cleanup passed. This is evidence for that paired source tree, not a replacement
for executing the focused #241 harness.

The focused harness adds the original four coverage entry points and their
three helpers to its existing 965 assertions: **1,947 expected assertions**.
It retains the high-precision boundary/caller references, inverse cases,
natural iteration-cap failures and recovery tests without loosening their
tolerances. `provenance.json` records the exact arithmetic, adapter and
regression-block hashes; `tests/test_gamma_public_source.py` checks them
without fetching an off-branch commit.

The focused source was executed in
[Excel run 42524911524](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/actions/runs/42524911524):
**1,947 / 1,947 PASS, zero failures**. All six production modules are byte-identical
to the paired public implementation. Import, execution-backed compile,
regression and cleanup passed on Excel 16.0 build 20430 / 64-bit. The recorded
session lasted approximately 2m 21s; this is suite time, not a general
performance claim. No grid export occurred.

| Focused evidence | Identity |
| --- | --- |
| Production/harness commit | `fe4c90f47e7b58bd089d720046d50fdac9b4581e` |
| Tested PR merge | `f214de606b180e3b3781c5c9806aa051c7e636f4` |
| Identical source/merge tree | `786f5a8c4f1a5636bc8195c93bf26f33fc1e942b` |
| Artifact | `13634354530` |
| ZIP SHA-256 | `719620d6c0ab75f74a75819824bbcaf6b2f0a06f29c25027b80c75fc9a33d790` |

`excel_artifact.json` retains the original artifact members and raw tested
merge commit. Fixtures verify its Git identity, equality with the reachable
source tree, all seven imported source hashes, policy counts, stages and
original log digest without rewriting the artifact's candidate SHA.
All 61 local tooling commands and all 18 strict static checks pass on the
focused port; the retained-evidence fixture also passes. Hosted static checks
and CodeQL pass on that tested source. At that regression-only checkpoint,
the live accuracy gate was blocked by stale grid bindings and the older
canonical export record; the subsequent export resolves those stale bindings.

The independent fresh cohort and full private arithmetic history are retained
in [PR #242](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/pull/242),
under `benchmark/gamma_cf_study/coverage/`. That branch preserves each frozen
protocol and original result, including the initial subnormal-reference
conversion rejection and its separately frozen single-round completion.
The fresh public host run
[42524515131](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/actions/runs/42524515131)
passed **14,878 / 14,878 assertions**, including all 1,696 checks of the fresh
206-point forward cohort and 24 inverse references. Its six production
modules are the same as this focused port; the larger test module is specific
to #242. See its
[result and limitations](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/blob/e91c535a8f7ef2ba3830d9c690b39966b74cc6b3/benchmark/gamma_cf_study/coverage/PUBLIC_GAMMA_RESULT.md).

Both PRs remain draft. The regression-only checkpoints above exported no
grids; the subsequent export and atomic evidence update are documented in
EXPORT_RESULT.md. Release coverage, other blockers and final review remain
required. No claim is made for untested Office environments or shapes.
