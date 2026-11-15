# P3 guarded stagnation — prospective protocol, 2026-11-12

Freeze model, runner and this protocol before candidate evaluation. P2 remains
rejected. Its two failed traces inform this design; they are fitting data.

P3 uses unchanged P2 orientation, DD continued fraction and 2^-100 stop.
Only before bracket rejection, permit a rounded Newton proposal equal to the
current interior iterate to settle if the density is finite and positive,
both adjacent binary64 neighbors are strictly interior, their forward CDF
residuals are finite and monotone through the current residual, bracket zero,
and the current residual magnitude is no greater than either neighbor's.
Otherwise use the original bracket/iteration/settling rules. This is a
study-only adjacent-float guard, not a proven VBA implementation.

Stage 1: all 45 revealed identities (13 retained, 8 P2 screen, 24 original
fitting). Retained affected and 8 screen rows: F relative error <2^-48;
retained controls additionally <=4x F-0. Original fitting: unchanged extended
fit rules (<=2x F-0 F error, retain clean 2^-48 points; balanced-huge pair,
residual and Newton measures <=2x F-0; fitting median iterations <=baseline+2).
All original failure criteria remain. No waiver near the rounding floor.

Only if all stage-1 rows and median gate pass, execute stage 2: p in {.15,.85}
crossed with {(5e4,8),(8,5e4),(2000,2000),(2e6,2e6)}. These 8 new identities
are disjoint from all previous fitting/screens and original holdout. Require
P3 F error <2^-48 and <=2x F-0, and median Newton count <=baseline+2.

Both stages use the checked CF oracle from run_extended_fit.py with 50/80-dps
reference agreement to 40 relative digits, finite outputs and original caps.
Record exact binary64 identities, full traces, errors, runtime, source hashes
and each guard activation. Failure stops downstream stages, without tuning.
Preserve failed rows; exceptions and unstable references are blockers.

Even a complete pass is SCREEN_PASS_NOT_ADOPTION. Geometry/CF seams,
Beta/Student-t invariance including small-shape capability, independent
holdout and source-bound Excel verification still need separate prospective
gates. No production VBA edit, holdout inspection, merge or closure here.
