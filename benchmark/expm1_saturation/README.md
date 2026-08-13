# Expm1 saturation probe

Freezes the `PROB_Expm1` subnormal saturation defect before and after the fix.
Both runs are on real VBA at Excel 16.0, 64-bit, Italian locale.

## The defect

`PROB_Expm1` rescales `U - 1` by `X / Log(U)` with `U = Exp(X)`. That ratio is
exactly 1 in exact arithmetic and stays within an ULP of 1 while `U` carries a
full 53-bit significand. Below `X` of about -708.4 the exponential is subnormal
and carries fewer bits, so `Log(U)` no longer recovers `X` and the ratio drifts.
Because `U - 1` has already rounded to exactly `-1` by then, the drift lands
undiluted and the kernel returned values **below -1** — outside the range of
`Exp(X) - 1`.

The probe records `Exp(X)` alongside the result, which makes the mechanism
visible rather than inferred: at `peak_74513` the exponential is
`4.94065645841247E-324`, the minimum positive subnormal, carrying a single
significant bit. `Log()` of a one-bit value cannot recover `-745.13`.

## Error growth, pre-fix

| point | X | `PROB_Expm1(X)` | relative error |
| --- | --- | --- | --- |
| `edge_709` | -709 | -1.0 | 0 |
| `growth_716_1ulp` | -716 | -1.0000000000000002 | 2.22E-16 |
| `growth_724_1e12` | -724.1 | -1.0000000000000595 | 5.95E-14 |
| `growth_731_1e9` | -731.2 | -0.9999999997568725 | 2.43E-10 |
| `growth_738_1e6` | -738.1 | -1.0000003893543672 | 3.89E-07 |
| `growth_742_1e4` | -742.6 | -0.999934945718797 | 6.51E-05 |
| `peak_74513` | -745.13 | **-1.000926774504277** | **9.268E-04** |
| `edge_74514` | -745.14 | -1.0 | 0 |

Below about -745.14 the exponential underflowed hard and the old `U = 0#`
branch took over, which is why the pre-existing assertions at -746 and -1E+300
passed throughout and never surfaced this.

## Public impact, pre-fix

`K_STATS_Exponential_Cumulative` and `K_STATS_Weibull_Cumulative` both compute
`-PROB_Expm1(-z)` and neither clamps, so both returned a **probability above
one**:

| point | Exponential CDF | Weibull CDF |
| --- | --- | --- |
| `growth_716_1ulp` | 1.0000000000000002 | 0.9999999999999999 |
| `growth_724_1e12` | 1.0000000000000595 | 1.00000000000006 |
| `growth_738_1e6` | 1.0000003893543672 | 1.0000003893543674 |
| `peak_74513` | **1.000926774504277** | **1.0009267745042765** |

Four of sixteen points violated the range invariant. **A clamp on the callers
would not have been a fix**: `growth_731_1e9` and `growth_742_1e4` returned
values *below* one while still being wrong by 2.43E-10 and 6.51E-05, so a
`<= 1` guard would have hidden the visible violation and left those two
silently wrong. The defect had to be repaired in the kernel.

`PROB_DS_TryGeometricCDF` reaches the same kernel but clamps to `[0, 1]`
afterwards, so the geometric family was protected by accident rather than by
design.

Two further call sites are unaffected and were checked rather than assumed:
`PROB_CN_TryWeibullLogVarianceFactor` is guarded to `0 < Delta < 0.5`, and
`PROB_LogExpm1` is defined only for `X > 0`.

## The fix

Branch on the difference rather than the exponential:

```vba
U = Exp(X)
V = U - 1#

If U = 1# Then          PROB_Expm1 = X
ElseIf V = -1# Then     PROB_Expm1 = -1#
Else                    PROB_Expm1 = V * X / Log(U)
End If
```

Wherever `V` has rounded to `-1`, `-1` is already the correctly rounded value of
`Exp(X) - 1`, so the compensation is not merely inaccurate there but
unnecessary. The new branch subsumes the old `U = 0#` case, needs no cutoff
constant, and repairs the kernel rather than papering over two callers.

`V` must be a **stored** Double. VBA evaluates a Double expression in wider
precision and rounds only on assignment, so testing `U - 1#` inline would not
ask the question the branch depends on.

## Containment, post-fix

- Every window point returns **exactly -1**; relative error is **0**, correctly
  rounded rather than merely improved.
- Both CDFs return **exactly 1** at every point. Zero range violations.
- The six benign points from -0.025 to -400, and the three past-window points,
  are **bit-identical** between the two runs — same `hi;lo` token, character for
  character.

That last line is the one that mattered. The new branch fires from
`X <= -37.43`, roughly 671 units wider than the defect it targets, so proving it
changes no value outside the window was the point of including the benign
ladder. A Python mirror over 8,582 points outside the window agreed: zero
differences.

## Coverage note

The defect survived a green harness and a green gate because nothing looked.
`Exponential_Cumulative.all.output` and `Weibull_Cumulative.all.output` are
frozen at 5E-15 relative, but the grid carries **three rows each**, the deepest
Exponential row at `Lambda * X = 1.5` — about 500x short of the failing region.
This is a coverage gap, not a missing threshold. Promoting far-tail rows into
the two existing contracts is tracked separately; the permanent CH3 and public
assertions in `tests/M_STATS_PROBDIST_TEST.bas` are what prevent recurrence in
the meantime.

## Scope

Characterization only. Promotes no grid row, claims no threshold and touches no
registry. `promote_grid_rows.py` remains the only sanctioned route for grid
changes.

## Files

| file | purpose |
| --- | --- |
| `M_STATS_PROBDIST_EXPM1PROBE.bas` | export macro; run `Probe_Expm1Saturation` |
| `expm1_saturation_probe_prefix.csv` | measured against the kernel at `a1de93f` |
| `expm1_saturation_probe_postfix.csv` | measured against the repaired kernel |

To reproduce the pre-fix run, remove `M_STATS_PROBDIST_CORE` from the VBE and
import it from `a1de93f`. Importing over an existing module does not replace it,
so confirm with `?PROB_Expm1(-745.13)` in the Immediate Window before running:
it must print `-1.000926774504277`, not `-1`.
