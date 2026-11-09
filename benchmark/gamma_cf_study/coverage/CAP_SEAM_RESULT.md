# B6 acceptance blocked by two new seam regressions

The expanded 1945-assertion suite adds 12 natural-cap/recovery checks and 52
adjacent-X monotonicity checks. Neither the production kernel nor thresholds
changed when adding these tests.

| Kernel | Head / synthetic merge | Excel run | PASS | FAIL |
| --- | --- | --- | --- | --- |
| Pre-B6 | c27dc1d / da405c6 | 42468593034 | 1757 | 188 |
| B6 | 15bb7ce / 0517cd9 | 42468461982 | 1905 | 40 |

The control tree replaces only SPECIALFUNCS with the release baseline blob.
Excel 16.0 build 20430, 64-bit, ran all assertions. All 12 cap/recovery checks
pass on both kernels: actual 100000-iteration exhaustion at A=1E14, X=A+1,
False with a diagnostic in the CF/P/Q kernels, #NUM! and diagnostic through the
Gamma CDF/SF wrappers, and successful subsequent recovery with cleared status.
No artificial iteration budget or success-on-exhaustion fallback was used.

There are **150 resolved failure labels, 38 common failures and two new B6
failures**. The 25 common accuracy failures have identical printed diagnostics;
the 13 common seam failures are Boolean assertions and carry no numerical
magnitude in the log. The first 1881-assertion comparison remains valid for its
sample but is superseded for acceptance by this broader result.

## New counterexample

A = 0.5, scale = 1, x0 = predecessor(1.5) = 1.4999999999999998,
x1 = 1.5. Exact input identities:

- A: `0x1.0000000000000p-1`
- x0: `0x1.7ffffffffffffp+0`
- x1: `0x1.8000000000000p+0`

B6 fails both required inequalities:

- Gamma_Cumulative(x0, A, 1) <= Gamma_Cumulative(x1, A, 1)
- Gamma_Survival(x0, A, 1) >= Gamma_Survival(x1, A, 1)

The baseline passes both. The new failure labels are `B6 X seam 05 Cumulative 0`
and `B6 X seam 05 Survival 0`. The source routes x0 through the unchanged lower
series/complement and x1 through B6 upper CF. This identifies the route junction;
it does not yet quantify the jump or isolate which arithmetic term must change.
An improved upper branch can expose a mismatch with the lower branch, but that
is not grounds to waive the new monotonicity failure. #241 must not merge in
this state. #26 and #24 remain open; fixing the seam requires a measured design
and fresh independent validation, not a boundary shift or tolerance relaxation.

Reproduce the complete label comparison:

```sh
python benchmark/gamma_cf_study/coverage/compare_excel.py \
  --baseline baseline_cap_seam_run_42468593034.txt \
  --candidate b6_cap_seam_run_42468461982.txt \
  --output paired_cap_seam_result.json
```

The excerpts retain every failure line. These are failed regression runs, not
fresh exports; no provenance manifest was rebound. The 24 previously revealed
holdout points cannot be called untouched evidence for a subsequent fix.
