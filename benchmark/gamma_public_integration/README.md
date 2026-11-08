# Public Gamma integration (#26; PRs #241 and #242)

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

The independent fresh cohort and full private arithmetic history are retained
in [PR #242](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/pull/242),
under `benchmark/gamma_cf_study/coverage/`. That branch preserves each frozen
protocol and original result, including the initial subnormal-reference
conversion rejection and its separately frozen single-round completion.

Both PRs remain draft. This regression integration does not rebind the main
or holdout accuracy grids: fresh exact-source Excel exports and the release
accuracy gates remain required. A green regression alone is not release
certification. No claim is made for untested Office environments or shapes.
