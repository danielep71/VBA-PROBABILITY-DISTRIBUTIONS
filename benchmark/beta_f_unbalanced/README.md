# Step-6 study: public Beta/F accuracy at unbalanced arguments

Measures the PUBLIC Beta and F worksheet functions directly at strongly
disparate shapes / degrees of freedom, now that `PROB_LogBeta` uses the stable
Lanczos log-gamma difference. This produces the function-level numbers that
freeze the regime-scoped accuracy contracts.

## Why measure the functions, not the LogBeta proxy

Downstream code uses `exp(-LogBeta)`, so to first order the relative error of an
exponentiated quantity is the ABSOLUTE error of `LogBeta`. That means:

- `PROB_LogBeta` should be judged by ABSOLUTE error (done in the delta seam study);
- the public Beta/F functions must be judged by their OWN relative error, because
  the incomplete-beta continued fraction, tail selection and inverse solver damp
  or amplify the normalization error differently for each function.

So this study measures each of `Beta_Density`, `Beta_Cumulative`, `Beta_Survival`,
`F_Cumulative`, `F_Survival` separately. Do NOT infer one common threshold.

## Grid

- **Beta**: strongly disparate `(Alpha, Beta)` — e.g. (0.7, 1000), (2.5, 1E6),
  (10.25, 68), (0.8, 1E4), (1000, 0.8), (1E5, 2.5) — evaluated at X near the
  distribution's mass (the mean and a mode-ish point).
- **F**: strongly asymmetric `(df1, df2)` — (1, 1E4), (2.5, 1E8), (10, 1E10),
  (1E6, 3) — at X = 1 and near the mean.
- 64 points; references are mpmath at 60 digits.

## Files

| File | Role |
|---|---|
| `generate_beta_f_unbalanced.py` | Writes `beta_f_unbalanced_grid.csv` (64 rows, 60-digit refs). |
| `beta_f_unbalanced_grid.csv` | The grid; `arg1 = X`, `arg2 = Alpha/df1`, `arg3 = Beta/df2`. |
| `M_STATS_PROBDIST_BETAF_UNBAL.bas` | Standalone macro `Export_BetaF_Unbalanced` (calls the 5 public functions; handles CVErr). |
| `analyze_beta_f_unbalanced.py` | Per-function worst-case relative error + suggested frozen thresholds. |

## How to run

1. Import `M_STATS_PROBDIST_BETAF_UNBAL.bas` into the workbook and `Debug > Compile`.
2. Run `Export_BetaF_Unbalanced`; select `beta_f_unbalanced_grid.csv` when prompted.
3. Commit the filled CSV.
4. Analysis (done for you): `python3 analyze_beta_f_unbalanced.py`.

## Freezing the contracts (completion criteria 5-7)

The analyzer prints, per function, the measured worst-case relative error and a
suggested frozen threshold (worst measured, rounded up with headroom). Then:

1. Set SEPARATE balanced and unbalanced contracts in `accuracy_contracts.csv`
   (the balanced claim stays tight; the unbalanced claim uses the measured value).
2. Replace the broad `known_limitation` with the precise scoped contracts.
3. Retain a `known_limitation` ONLY where a measured public function still exceeds
   its revised contract.

The measured VBA numbers are the authority; freeze only after running against the
final module.
