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
