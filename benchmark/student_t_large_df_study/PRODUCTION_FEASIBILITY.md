# Production feasibility — finding F1 (#34)

**Status: recorded before any VBA was written.** Excel-free. No `.bas`, contract,
threshold, grid, observation or manifest changes. This finding does not amend
`PREREGISTRATION.md` or the study's results; it measures whether the adopted
route can be built in production as it stands.

## Question

The study adopted the order-5 expansion

    S(t; df) = Q(t) + φ(t) · Σ_{k=1..5} g_k(t) / df^k     for df ≥ 1E5 and |t| ≤ 8

on the strength of its **truncation** error, every term evaluated at 60 to 100
digits. Production VBA evaluates the same formula in binary64 and takes `Q(t)`
from a kernel. The issue audit warned that the study's ~4.18E-16 truncation
bound "must not be presented as measured binary64/VBA implementation
accuracy". This finding measures that implementation accuracy.

## Method

`binary64_route_model.py` models the production arithmetic:

- `vba_upper_tail` is a line-for-line binary64 port of
  `PROB_NormalUpperTailPositive` (`src/M_STATS_PROBDIST_NORMALFAMILY.bas`): Hart
  rational approximation below `z = 7.07106781186547`, sixteen-level Laplace
  continued fraction above. Against every committed Excel observation of
  `NormalStandard_Survival` and `NormalStandard_Cumulative` (12), the port
  matches 9 exactly and the other 3 within 1.25 ulps, which is the last-bit
  difference between VBA `Exp` and the C library's `exp`.
- `route` evaluates the order-5 series in binary64: coefficients rounded to
  Double, each `g_k` by Horner in `t`, the sum by Horner in `1/df`, `φ` as
  `PROB_NormalPDF` forms it.
- Each point is compared with the study's primary oracle, `I_z(df/2, 1/2) / 2`
  at 60 digits, and with `C`, the study's recorded error of the current binary64
  continued fraction at the same point.

The 36 points are the study's forward **fitting** points inside the adopted
region (`t ∈ {0.5, 1, 1.5, 2, 3, 4, 5, 6, 8}`, `df ∈ {1E5, 1E6, 1E7, 1E8}`).
No holdout point is used.

## Result

Relative error. **WORSE** marks a point where the existing-kernel route is less
accurate than today's continued fraction and above the bound `2^-50` (≈ 8.9E-16).

| t | df | current CF | route, existing VBA Q | route, exact Q |
| ---: | ---: | ---: | ---: | ---: |
| 2 | 1E5 | 1.25E-12 | 9.25E-16 | 1.01E-17 |
| 2 | 1E8 | 7.69E-10 | 1.05E-15 | 1.34E-16 |
| 3 | 1E8 | 2.88E-09 | 1.51E-14 | 7.33E-18 |
| 4 | 1E8 | 1.12E-09 | 1.23E-13 | 1.93E-16 |
| 5 | 1E5 | 9.09E-14 | 4.58E-11 **WORSE** | 2.95E-17 |
| 5 | 1E6 | 6.66E-12 | 4.58E-11 **WORSE** | 1.01E-17 |
| 5 | 1E7 | 2.50E-11 | 4.58E-11 **WORSE** | 4.47E-17 |
| 5 | 1E8 | 1.04E-10 | 4.58E-11 | 3.56E-17 |
| 6 | 1E5 | 2.59E-13 | 5.33E-10 **WORSE** | 1.45E-16 |
| 6 | 1E6 | 1.46E-14 | 5.35E-10 **WORSE** | 2.82E-17 |
| 6 | 1E7 | 1.12E-11 | 5.35E-10 **WORSE** | 1.52E-16 |
| 6 | 1E8 | 5.03E-10 | 5.35E-10 **WORSE** | 1.05E-16 |
| 8 | 1E8 | 4.68E-10 | 7.53E-17 | 8.32E-17 |

`python binary64_route_model.py` prints all 36 points.

1. **The series is sound in binary64.** With a correctly rounded `Q(t)`, the
   route meets `2^-50` at all 36 points (worst 5.3E-16).
2. **The existing normal-tail kernel is not.** The route's error equals the
   kernel's: 1.5E-14 at `t = 3`, 1.2E-13 at `t = 4`, 4.6E-11 at `t = 5` and
   5.3E-10 at `t = 6`. Between 3 and 7.07 the Hart approximation loses relative
   accuracy, as `numerical_limitations.csv` already records under
   `SurvivalTailRel`. Above 7.07 the kernel switches to its continued fraction
   and is accurate again.
3. **The adopted predicate therefore fails §8's improvement requirement in
   production.** At 7 of the 36 points the route would be worse than what it
   replaces, by up to about 37,000 times (`t = 6, df = 1E6`). It would still remove
   the #34 counterexample (`t = 2, df = 1E8`: from 7.7E-10 to 1.1E-15), which
   is why a whole-region comparison hides the regression.

## Consequence

The "preregister predicted movement" step of #34 cannot be frozen for the
adopted design: it would preregister a production route that the study's own
rule forbids. The design needs one of the following decisions, followed by a new
preregistration of the production route, as the status note of `PREREGISTRATION.md`
requires when a design proves wrong:

- **A. Repair the normal tail first.** Give `PROB_NormalUpperTailPositive` a
  relative accuracy near 2^-50 for `z ≥ 0`, for example by a continued fraction
  or rational form tuned for `3 ≤ z < 7.07`. With an accurate `Q`, the adopted
  route meets the bound at every fitting point. The same repair would address
  the `SurvivalTailRel` limitation of `Normal_Survival`, `Lognormal_Survival`
  and `NormalStandard_Survival`, whose contracts are currently limited to
  `|z| ≤ 2.75`–`3.25`. It changes their observations and therefore needs its
  own evidence wave.
- **B. Narrow the route to where the existing kernel is accurate.** A bound near
  `|t| ≤ 2` keeps the route under ~1E-15 with today's kernel and still removes
  the `t = 2` counterexample. The CF route's errors above that remain (up to
  2.9E-9 at `t = 3, df = 1E8`), so #34 would stay partly open.
- **C. A Student-t-local accurate tail.** Compute `Q(t)` inside the new route
  by a method accurate on `[0, 8]`, leaving the shared normal kernel unchanged.
  This removes the regression without touching normal-family contracts, at the
  cost of a second normal-tail implementation.

`test_binary64_route_model.py` pins this finding. It fails if the kernel, the
adopted predicate or the measured conclusions change, so the finding cannot
silently become stale.

---

## Amendment 1 — true ULP distance and source binding (2026-11-05, #77)

I added this amendment after review of #70. The record above is unchanged; where
it disagrees with this amendment, this amendment holds.

**The ULP figures were relative epsilon, not ULPs.** The Method section's "the
other 3 within 1.25 ulps", and the fixture's original two-ulp check, divided the
difference by `|observed| · 2^-52`. That is relative error in units of machine
epsilon, not binary64 spacing. Below a power of two it understates the step count
by up to almost 2×, so a three-step mismatch could pass the two-ULP limit: at the
committed `NormalStandard_Survival(1.96)` value, three steps measure 1.875 on that
scale. I re-measured the 12 observations with the ordered-integer distance between
binary64 bit patterns (`model_fidelity.ulp_distance`: +0 and −0 coincide, and
subnormals, binade boundaries and the sign change are ordinary steps):

| Observation | Original scale | True ULPs |
| --- | ---: | ---: |
| `NormalStandard_Cumulative(1)` | 0.59 | 1 |
| `NormalStandard_Survival(1)` | 0.79 | 1 |
| `NormalStandard_Survival(1.96)` | 1.25 | 2 |
| the other 9 | 0 | 0 |

Corrected statement: the port matches 9 of the 12 observations exactly and the
other 3 within 1, 1 and 2 ULPs. The two-ULP limit is unchanged and still holds; no
committed observation exceeds it. The Result table reports relative error against
the oracle, not ULPs, and is unaffected.

**The attribution to `Exp` is withdrawn.** The Method section attributes the
differences to the last bit of VBA `Exp` against the C library's `exp`. That was
not verified. A probe I ran while writing this amendment reproduces both `z = 1`
observations with a one-step-larger `exp`, but no shift of `exp` by up to two steps
reproduces `NormalStandard_Survival(1.96)`. The cause of the differences is not
established; the two-ULP limit is an empirical tolerance on the committed
observations, not a derived bound. The module docstring of `binary64_route_model.py`
repeats the withdrawn sentence. I left that file unchanged because
`../normal_tail_study/results.json` binds its bytes.

**The model is now bound to its VBA source.** The finding relies on the port being
line for line, and the committed observations reach only `|z| = 3`: a change to the
Hart coefficients above 3, to the 7.07 split or to the continued-fraction depth
would not have shown. `source_binding.json` now records the SHA-256 of the text of
`PROB_NormalPDF`, `PROB_NormalUpperTailPositive`, `PROB_NormalCDF`,
`PROB_NormalSurvival` and the constants `PROB_INV_SQRT_TWO_PI`,
`PROB_TAIL_UNDERFLOW_Z`, `PROB_NORMAL_TAIL_CF_TERMS`, `PROB_CDF_SPLIT` and
`PROB_PDF_UNDERFLOW_Z` in `src/M_STATS_PROBDIST_NORMALFAMILY.bas`, and of the
model file.

A digest shows that the source changed, not that the model still mirrors it, so
`model_fidelity.py` also checks the extraction statement by statement. Its
inventory lists all 55 executable statements of the four procedures, in source
order. 51 are cross-checked. One is intentionally unchecked, the density's
underflow guard (4 statements). That guard returns 0 only at
`|Z| ≥ PROB_PDF_UNDERFLOW_Z = 38.6`, and the route evaluates `φ` only for
`|t| ≤ T_MAX = 8`. The binding requires the constant to stay above `T_MAX`. A
statement in neither group fails the check, so any new or changed statement forces
a decision.

The cross-checked statements are translated to Python and compared by syntax tree
with the model function that mirrors each procedure. The comparison keeps
operators, signs (including Python unary signs), literals, loop bounds, branches
and order. The constants are compared the same way. `test_binary64_route_model.py`
now fails when any of these differs. `test_model_fidelity.py` proves this with
mutations: a Hart coefficient, a Hart sign, the split, the continued-fraction depth
and initializer, a loop bound, a predicate, the density, its guard constant, a
comment, the model, and a statement in neither group.

The binding describes the source this finding measured: the `.bas` file last changed
in `a1de93f` (2026-08-13), before this finding was recorded in `1e09309`, and the
model is byte-identical to `1e09309`. With the corrected metric and the binding, the
fixture reproduces every conclusion above: the 36 points, the 7 **WORSE** points, the
`t = 6, df = 1E6` and `t = 2, df = 1E8` comparisons, and the kernel's degradation.

A change to the bound VBA, including the #72 normal-tail repair, makes this finding
stale. It stays stale until the finding is re-measured, this document is amended,
and the binding is re-recorded with
`python3 model_fidelity.py --record --reason TEXT`, which refuses while the model
does not mirror the source.

---

## Amendment 2 — the kernel F1 measured is frozen (2026-11-06, #81)

Option A is being carried out: #81 replaces `PROB_NormalUpperTailPositive` with
the repaired kernel of `benchmark/normal_tail_study` (#72). F1 is a finding about
the kernel it replaces, so it now stays bound to that kernel instead of to the
live source. Nothing above changes.

- **Source.** `f1_frozen/M_STATS_PROBDIST_NORMALFAMILY.bas.frozen` is
  `src/M_STATS_PROBDIST_NORMALFAMILY.bas` as commit `495d0cf` held it, byte for
  byte, with CRLF line endings like the module. `model_fidelity.py` reads the
  bound procedures and constants from it. Their recorded identities are
  unchanged.
- **Observations.** `f1_frozen/normal_observations.csv` holds the 12
  `NormalStandard_Survival` and `NormalStandard_Cumulative` observations of the
  main grid and the holdout at `495d0cf`, the ones Amendment 1 measured.
  `test_binary64_route_model.py` reads them there, because the #81 export
  replaces the grid observations with those of the new kernel.
- **Origin.** `source_binding.json` gains an `origin` block: the commit, the
  copy's Git blob id and the observations' SHA-256. `origin_problems` checks
  both always. Where the commit is in the clone, as in CI, it also rebuilds both
  copies from that commit and compares them byte for byte. A change to either
  copy fails the F1 fixture.

The repaired kernel is bound to its own model by
`benchmark/normal_tail_study/test_vba_kernel_source.py`. The normal-tail study
measured the #34 route with it at all 36 fitting points. A production route for
#34 is designed against that kernel, not this one.
