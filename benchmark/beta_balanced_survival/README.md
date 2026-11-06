# Balanced Beta survival measurement (#163)

**Status: measurement only.** This study changes no `.bas` module, grid,
observation, reference, contract or threshold. It measures the error of the
balanced `Beta_Survival` route against references taken at the binary64
argument, finds the cause, and records what the evidence supports. The
decision between a VBA repair and a contract recalibration is open in #163.

## Question

After #159 re-based the `holdout_older` references to the binary64 argument,
the committed Excel observation of `Beta_Survival(0.82, 2.7, 3.4)` has
relative error 5.45E-15. The contract `Beta_Survival.balanced.output_rel`
allows 5E-15. This study answers two questions. Is 5.45E-15 an isolated point
or typical of the region? Is it a defect in the route or a threshold the route
cannot meet?

## Route and contract

| Item | Where | What it does |
|---|---|---|
| `K_STATS_Beta_Survival` | `src/M_STATS_PROBDIST_CONTINUOUS.bas` | Calls `PROB_TryBetaRegularized(1 - X, X, Beta, Alpha)`: the swapped incomplete beta I_{1-x}(b, a). |
| `K_STATS_Beta_Cumulative` | same | Calls `PROB_TryBetaRegularized(X, 1 - X, Alpha, Beta)`. |
| `PROB_TryBetaRegularized` | `src/M_STATS_PROBDIST_SPECIALFUNCS.bas` | Prefactor Bt = X^A Y^B / B(A, B). Below A + B = 1000 (`PROB_IBETA_LOADER_MIN_SHAPE`) it is `Exp(A Log X + B Log Y - PROB_LogBeta(A, B))`; at and above, the Loader decomposition `PROB_TryBetaLogPdf`. Then the continued fraction: `Bt * CF / A` directly when X < (A + 1) / (A + B + 2), otherwise `1 - Bt * CF / B`. |
| `PROB_TryBetaContinuedFraction` | same | Modified Lentz; stops when abs(Del - 1) <= `PROB_NUM_EPS` (3E-14). |
| `PROB_LogBeta` | same | For ratio >= 0.1: `PROB_LogGamma(A) + PROB_LogGamma(B) - PROB_LogGamma(A + B)`. |
| `PROB_LogGamma` | same | Lanczos g = 7, n = 9, with coefficients written as 15-significant-digit literals (`PROB_LANCZOS_P0` to `P8`). |

`Beta_Survival.balanced.output_rel` is row 63 of `../accuracy_contracts.csv`:
relative error, threshold 5E-15, domain `min(shape)/max(shape) >= 0.1`,
provenance `validated and frozen`, evidence `holdout_older`, note
"Established balanced contract, unchanged". The value has no calibration of
its own. Commit `f6124fd` (2026-07-19) created the regime contracts with
provenance `source claim` and evidence `main grid` (three points); commit
`ecf3e45` (2026-07-21) changed the provenance to `validated and frozen` on
three `holdout_older` points, the worst 4.34E-15 (1.2x margin). That worst was
measured against the decimal-token reference, which is 1.1E-15 away from the
binary64-argument value. `Beta_Cumulative.balanced.output_rel` (2E-14) has
the same history. `Beta_Density.balanced.output_rel` was raised from 5E-15
to 1E-14 for the same reason ("prior 5E-15 marginally overfit").

## Method

**References.** Survival: `_survival_oracle.certify` (#159), I_{1-x}(b, a) at
the exact binary64 arguments, two working precisions agreeing to 32 digits.
Cumulative: I_x(a, b) by the same construction (`cdf_reference`).

**Binary64 model.** `beta_route_model.py` ports, statement by statement, every
procedure in the table above, the Loader branch (`PROB_TryBetaLogPdf`,
`PROB_TryDeviancePart`, `PROB_StirlingError`, `PROB_Log1p`) and
`K_STATS_Beta_Density`. It reads every constant from the source at import, and
`results.json` binds each ported procedure by the SHA-256 of its text. It
refuses (`NotModelled`) what it does not port: shapes below 0.5, ratios below
0.1 with the larger shape at least one (`PROB_LogGammaDelta`), and the exact
0.1 ratio seam.

**Arithmetic semantics, identified.** VBA keeps intermediates in x87 extended
precision (64-bit significand) and rounds to a Double on every store, ByVal
Double argument and return. Of the 57 committed `PROB_LogGamma` observations
with Z >= 0.5, plain binary64 evaluation reproduces 11 and 64-bit
intermediates reproduce all 57. The model uses the second.

**Fidelity.** Every committed Excel observation the model can evaluate is
compared with it, in all grids: 371 observations of `LogGamma`,
`LogGammaHalfDiff`, `PROB_LogBeta`, `StirlingError`, `Beta_Density`,
`Beta_Survival` and `Beta_Cumulative`. **All 371 are reproduced bit for bit
(0 ULP).** They include the six balanced `Beta_Survival` and
`Beta_Cumulative` observations and 134 `Beta_Density` observations, which
validate the Loader branch at small and large shapes. 88 observations are
excluded because they were exported before a later change to a procedure they
reach: the `beta_f_unbalanced` evidence (exported 2026-07-19, before the
Loader dispatch of 2026-07-29).

*2026-11-06 (#81).* Until then three `holdout_older` density and
Stirling-error rows were excluded too: exported on 2026-07-21, before later
`PROB_StirlingError` changes, they were up to 20,367 ULP from the model. The
#81 wave re-exported `holdout_older` (run 42210983927), and the model now
reproduces all three exactly, so they count with the rest. The balanced
`Beta_Survival` and `Beta_Cumulative` observations of that grid did not change.

**Point sets**, all deterministic:

| Sweep | Points | Design |
|---|---:|---|
| `ulp_neighbours` | 16 | x = 0.82 moved by 1 to 8 ULPs either way, (a, b) = (2.7, 3.4) |
| `local` | 650 | x = 0.70..0.95 step 0.01, a = 2.5..2.9 and b = 3.2..3.6 step 0.1 |
| `domain` | 1911 | shapes {0.5, 0.75, 1, 1.5, 2, 2.7, 3.4, 4.3, 5, 7, 10, 14, 19} in every ordered pair with ratio >= 0.1, x in {0.01, 0.05, 0.1, 0.2, ..., 0.9, 0.95, 0.99}; 26 seam points refused |
| `large` | 112 | shapes {20, 30, 50, 100} in every pair, x at the mean and +-1, 2, 3 sd |

Each point is evaluated in four variants of the model, to attribute the error:

| Variant | Change |
|---|---|
| `source` | none: the route as the source has it |
| `exact_lb` | `PROB_LogBeta` replaced by the correctly rounded Log(B(a, b)) |
| `lanczos17` | `PROB_LogGamma` with the 17-digit Lanczos coefficients |
| `loader_prefactor` | the Loader prefactor at every A + B, as if `PROB_IBETA_LOADER_MIN_SHAPE` were 0 |

## Results

### Committed Excel observations in the balanced ratio domain

Every committed `Beta_Survival` and `Beta_Cumulative` observation with ratio
>= 0.1 and both shapes in [0.5, 100). The gate error uses the committed
reference cell; the binary64 error uses the reference recomputed here. Since
the #81 wave re-based the two main-grid copies of `beta_f_unbalanced` rows
below (#159), their gate and binary64 errors agree. 210
rows fall outside: 114 with ratio below 0.1, 84 with a shape of 100 or more
(the `large_shape` and `extreme_shape` contracts) and 12 with both shapes
below 0.5 (`tiny_unbalanced`).

| Grid | Function | (x, a, b) | Regime | Branch | Gate error | Binary64 error | Threshold |
|---|---|---|---|---|---:|---:|---:|
| `probability_accuracy_grid` | Beta_Cumulative | (0.5, 2.0, 2.0) | balanced | reflected | 4.44E-16 | 4.44E-16 | 2E-14 |
| `probability_accuracy_grid` | Beta_Survival | (0.5, 2.0, 2.0) | balanced | reflected | 4.44E-16 | 4.44E-16 | 5E-15 |
| `probability_accuracy_grid` | Beta_Cumulative | (0.3, 2.0, 5.0) | balanced | direct | 3.22E-15 | 3.15E-15 | 2E-14 |
| `probability_accuracy_grid` | Beta_Survival | (0.3, 2.0, 5.0) | balanced | reflected | 4.31E-15 | 4.35E-15 | 5E-15 |
| `probability_accuracy_grid` | Beta_Cumulative | (0.8, 5.0, 1.0) | balanced | reflected | 5.70E-15 | 5.79E-15 | 2E-14 |
| `probability_accuracy_grid` | Beta_Survival | (0.8, 5.0, 1.0) | balanced | direct | 2.78E-15 | 2.82E-15 | 5E-15 |
| `holdout_older` | Beta_Cumulative | (0.35, 2.7, 3.4) | balanced | direct | 5.83E-15 | 5.75E-15 | 2E-14 |
| `holdout_older` | Beta_Survival | (0.35, 2.7, 3.4) | balanced | reflected | 3.04E-15 | 3.00E-15 | 5E-15 |
| `holdout_older` | Beta_Cumulative | (0.82, 2.7, 3.4) | balanced | reflected | 4.81E-17 | 9.72E-17 | 2E-14 |
| `holdout_older` | **Beta_Survival** | **(0.82, 2.7, 3.4)** | balanced | direct | **5.45E-15** | **5.43E-15** | 5E-15 |
| `holdout_older` | Beta_Cumulative | (0.15, 4.3, 1.9) | balanced | direct | 6.79E-16 | 6.68E-16 | 2E-14 |
| `holdout_older` | Beta_Survival | (0.15, 4.3, 1.9) | balanced | reflected | 3.00E-17 | 1.19E-17 | 5E-15 |
| `probability_accuracy_grid` | Beta_Cumulative | (0.13099041533546326, 10.25, 68.0) | unbalanced | direct | 9.76E-15 | 9.85E-15 | - |
| `probability_accuracy_grid` | Beta_Survival | (0.13099041533546326, 10.25, 68.0) | unbalanced | reflected | 1.12E-14 | 1.12E-14 | - |
| `probability_accuracy_grid` | Beta_Cumulative | (0.13564668769716088, 10.25, 68.0) | unbalanced | direct | 1.13E-14 | 1.11E-14 | - |
| `probability_accuracy_grid` | Beta_Survival | (0.13564668769716088, 10.25, 68.0) | unbalanced | reflected | 1.53E-14 | 1.53E-14 | - |
| `beta_f_unbalanced` | Beta_Cumulative | (0.13099041533546326, 10.25, 68.0) | - | direct | 9.92E-15 | 9.85E-15 | - |
| `beta_f_unbalanced` | Beta_Survival | (0.13099041533546326, 10.25, 68.0) | - | reflected | 1.12E-14 | 1.12E-14 | - |
| `beta_f_unbalanced` | Beta_Cumulative | (0.13564668769716088, 10.25, 68.0) | - | direct | 1.11E-14 | 1.11E-14 | - |
| `beta_f_unbalanced` | Beta_Survival | (0.13564668769716088, 10.25, 68.0) | - | reflected | 1.53E-14 | 1.53E-14 | - |
| `excel_comparison` | Beta_Cumulative | (0.3, 2, 3) | - | direct | 2.65E-15 | 2.66E-15 | - |
| `excel_comparison` | Beta_Survival | (0.999, 2, 3) | - | direct | 5.76E-16 | 5.55E-16 | - |

Of the six observations the balanced survival contract governs, one exceeds
5E-15. Two more committed Excel survival observations in the same ratio
domain exceed it by 2x to 3x: the (10.25, 68) rows, ratio 0.15, which the
main grid labels `unbalanced` and so scores against the 2E-10 unbalanced
contract, although that contract's domain is ratio below 0.1. Two balanced
`Beta_Cumulative` observations are above 5E-15 and within their own 2E-14.

### Sweeps, `source` variant

| Sweep | Points | Function | Worst | Median | 99th pct | > 5E-15 | > 2E-14 |
|---|---:|---|---:|---:|---:|---:|---:|
| ulp_neighbours | 16 | Beta_Survival | 6.03E-15 | 5.23E-15 | 6.03E-15 | 13 | 0 |
| ulp_neighbours | 16 | Beta_Cumulative | 1.70E-16 | 1.16E-16 | 1.70E-16 | 0 | 0 |
| local | 650 | Beta_Survival | 6.57E-15 | 3.16E-15 | 6.18E-15 | 108 | 0 |
| local | 650 | Beta_Cumulative | 9.66E-16 | 5.82E-17 | 6.19E-16 | 0 | 0 |
| domain | 1911 | Beta_Survival | 6.70E-14 | 1.07E-15 | 4.47E-14 | 525 | 165 |
| domain | 1911 | Beta_Cumulative | 5.94E-14 | 1.10E-15 | 4.27E-14 | 508 | 144 |
| large | 112 | Beta_Survival | 1.38E-13 | 5.43E-14 | 1.37E-13 | 79 | 64 |
| large | 112 | Beta_Cumulative | 1.34E-13 | 5.39E-14 | 1.30E-13 | 79 | 63 |

By the larger shape, the `domain` survival worst is 1.54E-14 up to 5 (87 of
1053 points above 5E-15), 4.00E-14 up to 10 (193 of 416) and 6.70E-14 up to
19 (245 of 442). 5.45E-15 is therefore typical, not isolated: 13 of the 16
binary64 neighbours of x = 0.82 and 108 of the 650 local points exceed 5E-15.
Across the domain, 27% of the points do, and the error grows with the shapes.
The cumulative fails its own 2E-14 at 144 `domain` points through the same
mechanism; no committed Excel observation shows that yet.

### Attribution, `Beta_Survival` worst (points above 5E-15)

| Sweep | source | exact_lb | lanczos17 | loader_prefactor |
|---|---:|---:|---:|---:|
| ulp_neighbours | 6.03E-15 (13) | 1.17E-15 (0) | 3.02E-15 (0) | 8.60E-16 (0) |
| local | 6.57E-15 (108) | 1.56E-15 (0) | 4.55E-15 (0) | 2.50E-15 (0) |
| domain | 6.70E-14 (525) | 1.42E-14 (23) | 2.43E-14 (146) | 1.11E-14 (54) |
| large | 1.38E-13 (79) | 1.04E-14 (12) | 6.49E-14 (66) | 4.82E-14 (14) |

## Diagnosis

The error is in the prefactor's normalization, `PROB_LogBeta`, not in a
complement and not in the continued fraction.

- **Not a complement.** At (0.82, 2.7, 3.4) the survival takes the direct
  branch, `Bt * CF / A`, with no subtraction. The cumulative at the same point
  is the complement `1 - Bt * CF / B` of the same quantity: the same absolute
  error, divided by a value 47 times larger, gives 9.7E-17. Whichever tail is
  computed directly carries the prefactor's relative error in full; the other
  tail dilutes it. That is why the survival and the cumulative show the same
  error distribution in every sweep.
- **Not the continued fraction.** At the #163 point it converges in 5
  iterations. With `PROB_LogBeta` exact and every other statement unchanged,
  the error at the point falls from 5.43E-15 to 4.1E-16, and the local worst
  from 6.57E-15 to 1.56E-15.
- **`PROB_LogBeta` error.** The relative error of the survival is minus the
  absolute error of the exponent `A Log X + B Log Y - LogBeta`. At (2.7, 3.4)
  `PROB_LogBeta` is 4.82E-15 too large, from the `PROB_LogGamma` errors
  +1.70E-14 at 3.4, +1.41E-14 at 2.7 and +2.71E-14 at 6.1. Two causes add:
  - **Lanczos coefficients truncated to 15 digits.** `PROB_LANCZOS_P2` is
    -1259.1392167224, 2.7E-12 from the 17-digit -1259.1392167224028;
    `P4` is off by 4.3E-13 and `P1` and `P3` by 1.1E-13. `PROB_LogGamma` is
    therefore 1E-14 to 3.5E-14 too large for Z between 2 and 40 (the model
    matches all 57 Excel LogGamma observations bit for bit), and
    LogGamma(a) + LogGamma(b) - LogGamma(a + b) does not cancel that error.
    With the 17-digit coefficients the error at the #163 point falls to
    2.2E-15.
  - **Conditioning of the literal prefactor.** The exponent subtracts terms of
    the size of LogGamma(a + b). Rounding alone gives an absolute error of 1 to
    3 x 2^-53 x (abs(A Log X) + abs(B Log Y) + abs(LogGamma(A)) +
    abs(LogGamma(B)) + abs(LogGamma(A + B))). That is the whole `large` worst,
    1.38E-13 at (50, 100), where LogGamma(150) is 600, and the reason the
    17-digit coefficients leave 2.4E-14 in `domain` and 6.5E-14 in `large`.
- **The Loader prefactor avoids both** (`loader_prefactor`), but it still
  rounds the exponent of the far tails: 54 `domain` points remain above
  5E-15, the worst 1.11E-14, and 4.82E-14 at (100, 100).

The header of `PROB_TryBetaRegularized` states that below A + B = 1000 the
literal form is "at least as accurate" as the Loader form, both near 1E-15 at
A + B = 200. The model does not reproduce that for the regularized
incomplete beta: at (0.410058, 50, 100) it measures the literal survival at
1.38E-13 and the Loader one at 9.0E-15. The committed Excel observations of
`Beta_Survival` and `Beta_Cumulative` at (100, 100) under the `large_shape`
contract already read up to 8.9E-14.

## What the evidence supports

No variant of the route that this study evaluates meets 5E-15 across the
balanced domain: not the source, not an exact `PROB_LogBeta`, not the
17-digit Lanczos coefficients and not the Loader prefactor. The threshold is
therefore the defect: it was carried forward as a source claim and frozen on
three points at 1.2x margin, and the route's error grows with the shapes to
more than ten times that value. A repair can lower the error, but none of the
candidates measured here lets the current contract pass. The evidence supports
a contract recalibration, on these terms; the decision and its review belong
to #163:

- `Beta_Survival.balanced.output_rel`: **2E-13**, with the domain stated as
  ratio >= 0.1 and both shapes in [0.5, 100). The modelled worst is 1.38E-13
  (1.45x margin): 6.70E-14 below shape 20 and 1.38E-13 from 20 to 100. The
  committed Excel worst in the domain is 1.53E-14 (1.55E-14 against the
  committed reference until #81 re-based it).
- `Beta_Cumulative.balanced.output_rel` shares the route and the mechanism and
  is exceeded by the model at 144 `domain` points (worst 5.94E-14; 1.34E-13 at
  shape 100). It needs the same recalibration, under its own issue.
- Before either threshold is frozen, confirm it in Excel on a fresh holdout of
  balanced points chosen where the model puts the worst errors, for example
  the `worst` entries of `results.json` (x = 0.99 at (14, 19), x = 0.7 at
  (19, 14), x = 0.41 at (50, 100)). The model is a measurement, not Excel
  evidence.
- Shapes below 0.5 with ratio >= 0.1 are not measured here, because the model
  does not port the LogGamma series or reflection branches.

If a tighter contract is wanted later, the model points to two repairs, for
validation in Excel: write the Lanczos coefficients at full precision as
split literals (the pattern `PROB_PI` already uses), and use the Loader
prefactor for balanced shapes. The Loader prefactor alone cuts the `domain`
worst from 6.70E-14 to 1.11E-14 and the `large` worst from 1.38E-13 to
4.82E-14. Either change moves `PROB_LogGamma` or `PROB_TryBetaRegularized`
results that other frozen contracts are measured on.

## Reproduce

```
python3 test_beta_route_model.py               # about 6 s; also in test_evidence_tools.py
python3 measure_balanced_survival.py --check   # about 2 min; recomputes results.json
python3 measure_balanced_survival.py           # rewrites results.json
```

## Files

| File | Role |
|---|---|
| `beta_route_model.py` | Binary64 model of the route, constants read from `src/`, procedures bound by hash. |
| `measure_balanced_survival.py` | Fidelity, observations and sweeps; writes `results.json`. |
| `results.json` | The recorded measurement. |
| `test_beta_route_model.py` | CI fixture: binding, fidelity, observations, the #163 point and each sweep's worst points. |
