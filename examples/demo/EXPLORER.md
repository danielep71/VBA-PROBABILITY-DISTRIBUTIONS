# Distribution Explorer

Development implementation for #254, dependent on the shared shell in #253.
The Explorer offers 16 distributions: Normal, Lognormal, Student t, Chi-square,
F, Gamma, Beta, Exponential, Weibull and Uniform (continuous), and Binomial,
Poisson, Geometric, Negative binomial, Hypergeometric and Discrete uniform
(discrete). Each shows its density or mass, cumulative probability and
survival probability P(X > x) at a probe point, the quantile at a probability
p, and plots the density and the cumulative. Moments are not offered here.

It needs desktop Excel with dynamic arrays (Microsoft 365 or Excel 2024) and
macros: every number is a call to the library's public `K_STATS_` functions.

## How the workbook is made

The package starts from the template
[`src/workbook/PROBDIST_Template.xlsx`](../../src/README.md#the-workbook): the
workbook the user formatted by hand, as the workflow last verified it, without
its VBA project, pinned by its SHA-256 in `.github/binary-allowlist.json`. Its
layout, styles, charts, Ribbon and navigation bar are the presentation
baseline and are not redrawn. `tools/demobuild/explorer_template.py` makes a
new template from a verified workbook when the layout changes.

1. `python tools/demobuild/explorer_build.py` copies every part of the
   template, writes the Ribbon from `src/workbook/customUI14.xml`, adds a VBA
   project built from the sources of the commit (the library of `src/modules/`, the
   modules of `src/workbook/` and the workbook events of
   `src/workbook/ThisWorkbook.cls`), and writes `explorer-plan.json`.
2. In desktop Excel, `tools/demobuild/Evolve-Explorer.ps1 -PackageDirectory
   test-results/demo` applies the plan: the family table, the workbook names,
   the formulas, the validation and conditional formats, and the two charts.
3. `ci/Test-Explorer.ps1`, `ci/Test-ExplorerStates.ps1`,
   `ci/Test-ExplorerUI.ps1` and `ci/Test-DemoPackage.ps1` check the result;
   `ci/Test-SyncWorkbook.ps1` builds a second workbook from the template with
   `tools/Sync-PROBDIST-VBA.ps1` and checks it.
   The `Distribution Explorer` workflow runs all of it and keeps the workbook,
   a PDF preview and the JSON evidence.

Final workbook bytes are not deterministic, because Excel writes the file:
bind acceptance to the digest in the test reports.

Everything about a family is data in `explorer-families.json`: its functions,
parameters with limits and defaults, probe point, plot window and support
text. `explorer_plan.py` turns it into formulas, so adding a distribution is
one entry there and a reference case in `explorer_references.py`.

## How the sheet works

Inputs are D6:D16: distribution, up to four parameters, probe point x or k,
probability p, plot lower and upper bounds, number of points and number of
decimals. The status (D19), the macros cell (D20) and the results (D21:D25)
sit below them, the notes in F5:I17 and the grid from F20. The
family table lives in the hidden columns from LZ, read through the `Explorer*`
workbook names. Picking a distribution in D6 loads its default parameters,
probe point and plot window; editing a parameter refits the window.

Outputs are dynamic arrays, one formula each:

| Cell | Spills | Formula |
| --- | --- | --- |
| `F20` | the x column | `SEQUENCE` between the bounds, or every integer for a discrete family |
| `G20`, `H20`, `I20` | density or mass, survival, cumulative | `MAP` over the x column, one `SWITCH` on the distribution |
| `D21` | the three results | `VSTACK` of density, cumulative, survival at the probe |
| `D24` | the quantile | the smallest x with CDF(x) >= p |

The names `ExplorerX`, `ExplorerPDF`, `ExplorerSurvival` and `ExplorerCDF`
refer to those spills (`F20#`, `G20#`, `H20#`, `I20#`); both charts take their
series from them, so the grid
and the charts follow the plot window and the number of points with no fixed
rows. The `K_STATS_` functions do not accept arrays, hence `MAP` over the x
column rather than passing it whole.

Status (D19) says `OK` or names the first parameter that is out of range, the
unmet constraint (for example, a hypergeometric sample larger than its
population), the probe point, the plot window or the number of points. Until
it says `OK` the outputs are `#N/A`. The quantile depends only on p as well:
a p outside 0 to 1 leaves it `#N/A` and the other results as they are.

The survival column is the library's own survival function, not 1 - CDF, so
the right tail keeps its digits where the cumulative is 1 to double precision.
The number of decimals (0 to 15, a list in D16) formats the results and the
grid: the Explorer module sets their number format when D16 changes, so it
needs macros like the rest of the sheet. A discrete x column stays in whole
numbers. Conditional formats, one per value, were tried first and rejected:
each rule over the spilled grid took Excel 40 to 80 seconds to add. A density that is undefined at a grid
point (a pole) is `#N/A` by the library's design.

Every input cell has an input message and an error alert, in the style of
the K-SACCR template: a title and a text that say what the cell holds, its
unit and its limits, and an error that says what to enter instead. The
parameter cells D7:D10 take the messages of the selected family's parameters,
written in `explorer-families.json` (`help`, and `note` for a constraint
between parameters); the Explorer module sets them when a family loads and
when the workbook opens.

Parameters are validated by named rules, one per slot, that read the limits
from the family table. Limits: continuous grids up to 1,000 points; discrete
grids up to 1,001 integers; probe point and bounds within +/-10,000,000. A
discrete family plots its mass as columns and its cumulative as points; a
continuous one plots lines. Geometric and negative binomial count failures,
not trials, and say so in their support text. Hypergeometric takes the sample
size, the successes in the population and the population size.

The Explorer sheet is unprotected, as in the template; the result cells are
locked and the inputs are not. Reset inputs on the Ribbon asks for confirmation and
restores Normal(0,1), probe 0, p = 0.95, 201 points and 6 decimals. Cancelling changes nothing.
On every sheet with inputs the button resets that sheet only; on Home and README it is
disabled. [RIBBON.md](RIBBON.md) documents every control of the Ribbon.

## When macros are not running

With macros off nothing is live: the file shows the values it was saved with,
and an edit turns the results into `#NAME?`. A downloaded file opens this way
until it is unblocked (file Properties, Unblock) or its content is enabled.
The Macros cell (D20) reads `ENABLED` or `DISABLED`, and a red line above the
tables appears when it is disabled. `Workbook_Open` recalculates and redraws
the charts, keeping the Saved flag, so a file saved with macros off is right
as soon as macros run.

## Evidence

`explorer_references.py` regenerates `explorer-references.json` with mpmath at
60 digits, from the definitions: closed-form densities, incomplete gamma and
beta functions, finite sums, and complement functions for the tails. It needs
mpmath; without it the unit tests check the stored file only. The file holds
80 quantile cases and 143 probe cases (several parameter sets per family, with absolute tolerance
1e-12 plus relative 1e-9) and, for each family at its defaults, the plot
window, the number of grid rows and five grid points with density, survival
and cumulative.

`ci/Test-Explorer.ps1` checks, in desktop Excel, every probe and quantile case;
the decimals; every
family's default window, spilled row count, five grid points and chart series
and type; that an edit moves the grid and the charts; the parameter rules; an
invalid input; and the reset. `ci/Test-ExplorerStates.ps1` does each family
through single edits with events on, the state with macros off, and the
recovery of a file saved in that state, while watching for any dialog.
`ci/Test-ExplorerUI.ps1` drives the real Ribbon through Windows UI Automation
in an owned Excel process: Recalculate, Home, the Go to destinations, the
navigation toggle, Reset inputs disabled on Home and README, the reset dialog,
Cancel and OK, a renamed tab and a deleted sheet.
`ci/Test-HandlerErrors.ps1` induces two errors in the Explorer's input handler with a
second workbook open, one before and one after the sheet is unprotected, and checks that
events, screen updating, calculation, the sheet's protection and the other workbook are
as they were, with one message shown (#482).

These checks passed on source `45289de7b6faca8c6cd21d1ed6231556016b2f3f` in
[run 44264560770](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/actions/runs/44264560770),
Excel 16.0 build 20710, x64. This is demonstration coverage, not a new
accuracy certification, and native UI automation, not a human usability
review. Not checked: Office architectures and regional settings other than
the runner's, and typing an out-of-range value into a validated cell (a write
through COM bypasses data validation, so the rules are checked as formulas).
Do not close #254 solely from package checks.

## The Probabilities sheet

The Probabilities sheet (#255) reads the Explorer's family table, so the 16 distributions, their limits and their input messages are defined once. Its plan is `tools/demobuild/probabilities_plan.py`, applied by `Evolve-Probabilities.ps1` after the Explorer's; its VBA is `src/workbook/M_STATS_PROBDIST_PROBABILITIES.bas`. Inputs are D6:D14 (distribution, up to four parameters, the points a and b, the probability p, the number of decimals); the status is D17, the macros cell D18 and the results D20:D28: P(X <= a), P(X > a), P(X <= b), P(X > b), P(a < X <= b), the quantile Q(p), F(Q(p)), the round-trip difference and its verdict. The grid is F20:I1020: x (or the integer k), the density or mass, and two columns that keep it where x <= a and where a < x <= b and hold 0 elsewhere, for the shaded chart areas.

A discrete family truncates a and b to whole numbers, P(X <= a) includes a, and its quantile is the smallest k with F(k) >= p. The interval is S(a) - S(b) when F(a) is above one half and F(b) - F(a) otherwise. The round trip compares F(Q(p)) with p, relative to min(p, 1 - p), within 1E-8; a discrete family has no exact round trip and the verdict states the definition. The library's quantile contract covers p from 0.01 to 0.999 and the deep lower tail (`benchmark/accuracy_contracts.csv`); above 0.999 the verdict may say OUTSIDE.

`probabilities_references.py` regenerates `probabilities-references.json` with mpmath (112 cases: each family at its defaults with three probabilities, both tails, fractional points for the discrete families and probabilities at the ends of the support). `ci/Test-Probabilities.ps1` sets each case through the sheet's cells with events on and compares the results, then checks the chart of each kind, the shaded columns, an invalid input and the reset.

## The Parameter Effects sheet

The Parameter Effects sheet (#256) reads the Explorer's family table too. Its plan is `tools/demobuild/parameters_plan.py` (data: `examples/demo/parameter-effects.json`), applied by `Evolve-ParameterEffects.ps1`; its VBA is `src/workbook/M_STATS_PROBDIST_PARAMETERS.bas`. Inputs are D6:D16: the distribution, up to four base parameters, the parameter to vary (D11, a list of that family's parameters), the four values it takes (D12:D15) and the number of decimals. Choosing a family loads its base parameters, the parameter it varies first and that parameter's four values; choosing another parameter loads its four. The status is D19, the macros cell D20; the four curves are in C22:I25: the curve and its value, and its mean, standard deviation, skewness and P(X > mean + 3 sd); the line on what the parameter does is C27:D32.

A curve is the family at the base parameters with the varied parameter set to its value, so every formula of curve k opens with `LET(pa, IF(slot = 1, value, $D$7), ...)` and the Explorer's dispatch (one SWITCH over the distribution) is reused with pa to pd in place of the input cells. The curves are drawn from hidden columns (x in R, the four densities in S:V, the four cumulative functions in W:Z, the window and the legend texts in AX:AZ, the table of the families from AB), through the `ParameterEffects*` workbook names; the two charts show the union of the four plot windows. The moments are written in the parameters (`parameter-effects.json`) and need no library function; where one does not exist (a Student t mean at 1 degree of freedom, a variance at 2, an F variance at 4 denominator degrees of freedom) the cell reads undefined or infinite. The tail column is the library's survival function at mean + 3 sd.

`parameter_effects_references.py` regenerates `parameter-effects-references.json` with mpmath: for every family and parameter the four default values, and for the first parameter varied four values typed between them. The mean, variance and skewness come from raw moments that are integrated from the density (continuous) or summed from the mass function (discrete), not from the closed forms of the sheet; whether a moment exists is a stated rule of the distribution. `ci/Test-ParameterEffects.ps1` sets each case through the sheet's cells in Excel and compares the table, the shared window and five points of the grid of each curve; it also checks both chart kinds, the choice of another parameter, an invalid curve value, a paste over several cells and the reset.

## The Comparison sheet (Excel vs VBA vs Python)

The Comparison sheet (#257) puts Excel's own functions, the library and SciPy side by side on the 29 cases of the [Excel comparison study](../../benchmark/excel_comparison/README.md): a body point of every family, and the deep tails, large shapes and missing functions where a difference is expected. Its plan is `tools/demobuild/comparison_plan.py` (data: `examples/demo/comparison-cases.json`), applied by `Evolve-Comparison.ps1`; its VBA is `src/workbook/M_STATS_PROBDIST_COMPARISON.bas`.

Inputs are D6:D9: the case shown in detail, the two grade thresholds in correct digits (12 and 6, the study's) and an optional cell for an =PY() formula. The status is D19, the macros cell D20. Rows 21 to 35 show the selected case: the Excel formula and value, the library formula and value, the reference with its 50 digits, the Python code and its stored SciPy value, a line for each implementation against the reference (difference, relative difference, correct digits, grade), why the case is in the study, the Excel that computed the sheet (version, build, bitness) and the provenance. The table C37:O66 holds every case: the reference, the three values, their correct digits and grades; the chart on the right shows the correct digits of every case.

- The Excel and library columns are live formulas, written as the study lists them; where Excel has no equivalent function the cell reads "none". The Excel formulas are the study's, `1 - CDF` included: they are not always the best Excel formulation (the study's README names better ones), and the sheet does not claim they are. The detail rows show the two formulas as texts stored beside the table, which the tests compare with the live formulas: FORMULATEXT would be recalculated at every opening and leave the workbook marked as changed.
- Each value is graded against the mpmath reference only (`benchmark/excel_comparison/generate_excel_comparison.py`, 50 digits): Excel is one of three implementations, not the yardstick. Correct digits are -log10 of the relative difference, from 0 to 16; a zero reference has no relative difference, so there only an exact 0 counts. The grades are the study's: reference from 12 digits, report from 6, inadequate below, wrong at 0, "no function" and "error" where there is no number. A grade below reference is a difference from this reference, not a certified defect: the library's contracts are in `docs/`.
- Where Excel has no usable value, the sheet says so in colour: the *Excel grade* column of the table and the Excel line of the selected case are red for "wrong" and amber for "no function" (a conditional format; the other grades are not coloured, so the eye goes to where Excel cannot be used).
- The SciPy values are stored by `comparison_references.py`, with the versions of Python, NumPy, SciPy and mpmath, so the sheet needs no Python. The code of each case is the SciPy call with the same inputs, parameterization and tail, written so it can be pasted into an =PY() formula: in D9 its value is graded like the others; empty, the stored value stands in.

`ci/Test-Comparison.ps1` checks in Excel that every Excel and library formula is the study's, that the digits shown are the ones computed from the same values, that the grades follow the thresholds, that the library and SciPy are reference grade on every case, and the detail of every case, the =PY() cell with a number and with text, other and contradictory thresholds and the reset. Its report, `comparison-report.json`, records what that Excel returned for every case, with its version and build.

## The Tails sheet (Extreme Tails and Numerical Limits)

The Tails sheet (#258) shows what two different things do to a small probability: cancellation and underflow. Its plan is `tools/demobuild/tails_plan.py` (data: `examples/demo/tails-cases.json`, written by `tails_references.py`), applied by `Evolve-Tails.ps1`; its VBA is `src/workbook/M_STATS_PROBDIST_TAILS.bas`.

A scenario is a family at fixed parameters and a ladder of 10 to 27 points that runs from the body of the distribution to far beyond 2.2E-308, the smallest positive number an Excel cell holds. There are 13: the upper tail of the Standard Normal, the Exponential, the Weibull, the Gamma, the Chi-square, the Student t, the Poisson and the Binomial (the library's survival function against `1 - CDF`), and the log-mass of the Binomial, the Poisson, the Geometric, the Negative binomial and the Hypergeometric (`LogPMF` against `LN(PMF)`).

Inputs are D6:D8: the scenario and the two grade thresholds in correct digits (12 and 6, the Comparison sheet's). The status is D19, the macros cell D20. Rows 21 to 35 describe the scenario: its parameters and units, the two paths as formulas, the reference and its provenance, the release guarantee of the function, where the alternative stops being usable, the deepest point where the direct path is at reference grade, the representation limit, and the Excel that computed the sheet. The ladder (C37:M..) lists each point with its reference, the two values, their correct digits and statuses and the *depth* of the point, -log10 of the probability. Two charts show the correct digits against the depth, and log10 of the values against it, each with the limit of the Double range drawn.

- **Cancellation is not underflow.** The statuses say which: *zero: cancellation* is `1 - CDF` equal to 0 although the true tail can be stored (about 1E-16 and below); *zero: underflow* is a true value below 2.2E-308, where every method returns 0; *PMF underflow* is `LN` of a mass that underflowed to 0, where `LogPMF` is still finite. Wrong, inadequate, report grade and reference grade follow the thresholds.
- **The reference is never `1 - CDF`.** It is mpmath at 60 digits, stored with 34, as text (a tail below 1E-308 is not a Double), with its depth, which is a Double. The unit tests recompute every reference by another method: exact integers for the discrete families, a Decimal sum for the Poisson tail, closed forms for the Exponential and the Weibull, `math.erfc` for the Normal where it does not underflow and SciPy's `logsf` where it is installed.
- **Supported domain, stated.** Only the discrete families have a log function: there is no log-survival and no log-CDF, so a continuous tail below 2.2E-308 is not available from the library. The release guarantee of each function is quoted from the user manual; the far-tail points of a ladder are a demonstration, measured on this run, and do not extend it. An error or a zero shown on the sheet is what this Excel returned.
- The scenario is chosen from a list because Excel cannot compute a 50-digit reference: the references are stored, and the sheet works without Python. The calls of the two paths are live formulas of the library.

`ci/Test-Tails.ps1` checks in Excel, for every scenario, the ladder the sheet spills, both paths against the formula written out in a scratch cell, the digits and statuses recomputed from the values, the colour Excel displays for each status, the two charts, other and contradictory thresholds, an unknown scenario and the reset. Its report, `tails-report.json`, records what that Excel returned at every point, with its build.

## The Discrete sheet (Binomial and Poisson counts)

The Discrete sheet (#259) reads a count two ways a reader can get wrong: a threshold, which a whole number makes inclusive or exclusive, and a quantile, which is a step and not an inverse. Its plan is `tools/demobuild/discrete_plan.py` (cases: `examples/demo/discrete-cases.json`, written by `discrete_references.py`), applied by `Evolve-Discrete.ps1`; its VBA is `src/workbook/M_STATS_PROBDIST_DISCRETESHEET.bas` (named apart from the library's `M_STATS_PROBDIST_DISCRETE`).

Inputs are D6:D12: the model, the trials n and the probability p of the Binomial, the event rate and the exposure of the Poisson (whose mean is rate x exposure), the threshold k and the probability q. The inputs of the model not chosen are greyed and ignored. The default is a simplified count of defaults: 100 loans that default independently with probability 2% each, and, as a Poisson, 0.02 defaults per loan-year over 100 loan-years, the same mean of 2. The status is D19, the macros cell D20.

Rows 21 to 31 give the mean, the standard deviation, P(X = k), the four readings of the threshold, the quantile Q(q), F(Q - 1), F(Q) and the check of F(Q - 1) < q <= F(Q) with the numbers in it; beside each result, in F:I, its formula as written. Rows 32 to 35 state the assumptions of the model, the Excel that computed the sheet and the provenance. The table (C37:I..) lists every count of a window around the mean (mean - 5 sd to mean + 5 sd, widened to hold k and Q, at most 1000 rows) with its mass and the four readings; the count k is marked in amber and the quantile in green. Two charts draw the mass as columns with the counts from k up in front, and the cumulative function as columns with the quantile's column in front and q as a line.

- **Four readings of a threshold.** At most k is P(X <= k); fewer than k is P(X <= k - 1); at least k is P(X > k - 1); more than k is P(X > k). Each is its own call, a survival value is never `1 - CDF`, and at k = 0 fewer than k is 0 and at least k is 1 without a call (the library refuses a count below the support). In the table, a row's *fewer than* is the row above's *at most*, and its *at least* the row above's *more than*.
- **A quantile is a step.** Q(q) is the smallest k with F(k) >= q, so F(Q - 1) < q <= F(Q): there is in general no k with F(k) = q, and F(Q) is above q unless q falls exactly on a step. The step-boundary case puts q = 0.74 just below the step F(1) = 0.75 of a Binomial with n = 2 and p = 0.5, so Q = 1. A q exactly on a step is not used: in binary64 the computed tail at that step can differ from q in the last bit, and the comparison then picks the next count.
- **Supported domain, stated.** n is a whole number from 1 to 10,000, p and q are strictly between 0 and 1, the rate is above 0 and at most 1,000, the exposure above 0 and at most 1,000,000, rate x exposure from 0.001 to 5,000, k a whole number from 0 (at most n for the Binomial). Every input has an input message and an error alert; a value outside these limits makes the status name it and blanks the results.

`ci/Test-Discrete.ps1` checks in Excel nine cases (the default example and its Poisson counterpart, a fair coin, a step boundary, k = 0, k = n, a mean of 500, rare events in 5000 trials, a quantile beyond mean + 5 sd): every result against an mpmath reference within 5E-8 relative, the loosest contract of these functions; the formula shown beside each result evaluated again in a scratch cell; the table row by row; the charts and the colours; then eleven invalid inputs one at a time, the Poisson inputs, a table too wide and the reset. Its report, `discrete-report.json`, records what that Excel returned for every case, with its build. The unit tests recompute the references by exact fractions (the Binomial, from the exact value of the Double p) and a Decimal series (the Poisson).

## The Risk sheet (value at risk, expected shortfall, tails and default counts)

The Risk sheet (#260) puts the library's distributions to four uses of risk management: a value at risk and an expected shortfall, a Normal against a Student t with the same variance, the probability of a loss above a threshold, and the count of defaults of a portfolio of loans. Its plan is `tools/demobuild/risk_plan.py` (cases: `examples/demo/risk-cases.json`, written by `risk_references.py`), applied by `Evolve-Risk.ps1`; its VBA is `src/workbook/M_STATS_PROBDIST_RISK.bas`.

Inputs are D6:D15: the position value V, the mean daily return, the daily volatility, the horizon in days, the confidence level, the Student-t degrees of freedom nu, a loss threshold, and for the defaults the number of loans, the default probability and the loss per default. The default example is a position of 1,000,000 with a daily volatility of 1% and a mean daily return of 0.03% over 10 days at 99%, nu = 4, a threshold of 50,000, and 100 loans that default independently with probability 2% and a loss of 10,000 each. The status is D19, the macros cell D20.

Rows 21 to 33 give the mean and the standard deviation of the loss over the horizon, the Normal VaR and ES, the Student-t scale, VaR and ES, P(L > threshold) under both models, the expected defaults, the default-count VaR Q, the loss it implies and P(defaults > Q); beside each result, in F:I, its formula as written. Row 34 reads the two models side by side; rows 35 to 38 state the assumptions, the Excel that computed the sheet and the provenance. The table (C40:I141) gives 101 losses from the mean - 4 sd up to the mean + top sd, where top is 6 or more, enough to hold the larger expected shortfall, with the two densities, the two exceedance probabilities and their ratio. Two charts draw the two densities with a dashed line at each VaR, and the two exceedance probabilities on a log scale with a line at 1 - confidence.

- **Losses, units, horizon.** A loss is L = -V R, positive when money is lost, in the currency of V. Daily returns are independent with the same mean and volatility, so over h days the mean loss is m = -V mu h and its standard deviation sd = V sigma sqrt(h).
- **The same variance.** The Student t is scaled by s = sd sqrt((nu - 2) / nu), so its variance is sd^2 like the Normal's. That needs nu > 2: with nu <= 2 the variance is not finite and the status says so. At the same variance the t has a narrower body and heavier tails, so at a moderate confidence its VaR can be the lower one; far enough out its VaR, ES and P(L > x) are larger.
- **VaR and ES.** VaR is the confidence-level quantile of the loss (`K_STATS_Normal_InverseCumulative`, or m + s times `K_STATS_StudentT_InverseCumulative`). ES is the mean loss beyond the VaR, from the closed forms m + sd phi(z) / (1 - a) and m + s (nu + t^2) / (nu - 1) f(t) / (1 - a); it is never below the VaR.
- **Default counts.** The defaults are a Binomial of independent loans: the count's VaR is its quantile, the smallest Q with P(defaults <= Q) >= confidence (`K_STATS_Binomial_InverseCumulative`), so P(defaults > Q) <= 1 - confidence. Real defaults are correlated; independence makes the tail of the count thinner than a real portfolio's.
- **Supported domain, stated.** V above 0 and at most 1E12; the mean daily return from -0.1 to 0.1; the volatility above 0 and at most 1; the horizon a whole number from 1 to 250; the confidence level from 0.5 to 0.9999; nu above 2 and at most 1000; the threshold between -1E13 and 1E13; the loans a whole number from 1 to 10,000; the default probability strictly between 0 and 1; the loss per default from 0 to 1E12. Every input has an input message and an error alert; a value outside these limits makes the status name it and blanks the results.

`ci/Test-Risk.ps1` checks in Excel seven cases (the default example, ES at 97.5%, nu = 2.5 near the variance limit, nu = 200 close to the Normal, one day with a zero mean, a positive mean loss at 90%, 99.99% with nu = 3): every result against an mpmath reference within 1E-9 relative (of the standard deviation for an amount), the formula shown beside each result evaluated again in a scratch cell, the reading, the table row by row and the charts; then fifteen invalid inputs one at a time, a Normal VaR of 0 (zero mean at confidence 0.5, where the reading gives the two VaRs instead of a ratio) and the reset. Its report, `risk-report.json`, records what that Excel returned for every case, with its build. The unit tests recompute the references by another method: the expected shortfall by numerical integration, the Student-t tail and quantile by integrating the density, the Normal tail by the complementary error function, the Binomial by a direct sum of masses, and SciPy.

## The Diagnostics sheet (errors and diagnostics)

The Diagnostics sheet (#261) shows what the library returns on valid input, at the edge of a support and on input it refuses, and how to read the reason. Its plan is `tools/demobuild/diagnostics_plan.py` (cases: `examples/demo/diagnostics-cases.json`, written by `diagnostics_references.py`), applied by `Evolve-Diagnostics.ps1`; its VBA is `src/workbook/M_STATS_PROBDIST_DIAGNOSTICS.bas`.

Inputs are D6:D12: a function from a list of 28 public functions, its arguments 1 to 4 and the range over which the chart sweeps argument 1. The arguments take any value, a probability of 1, a negative shape, a text or an empty cell included: what the library does with it is what the sheet shows. Choosing a function loads its example arguments and sweep; the label of each argument names the parameter. The default is `K_STATS_Gamma_InverseCumulative(0.95, 2, 1.5)`, with the probability swept from 0 to 1. The status is D19, the macros cell D20.

Rows 21 to 33 give the call with its values, the value returned (and beside it, in F:I, its formula), the kind of that value (a number, `#NUM!`, `#VALUE!`), the Status the library wrote, what they mean and what to do, the domain of the function as chapter 6 of the manual states it, the arguments Excel converted before the call (an empty cell passed as 0, a text or an error value that stops the call), the sweep in one line, the case table in one line, what the sheet does not show, the Excel that computed it and the provenance. One chart draws the function across the sweep as a line with a gap at each error: the default shows the probabilities 0 and 1 refused at both ends. It has one series: a second series of markers at the errors, in another chart type, made Excel stop with an access violation in the chart macros of other sheets (three runs out of four on this branch; none without it).

- **Reading Status from a worksheet.** A worksheet formula cannot read the `Status` argument. `DiagnosticsStatus`, a function of the sheet's VBA module, makes the same call from VBA and returns the text the library wrote, unchanged: empty when the call succeeds. When VBA cannot pass an argument as a number (a text that is not a number, an error value) the function never runs, and `DiagnosticsStatus` says so in a line that starts "Not called".
- **The case table (C40:O70).** Thirty calls, each with the outcome the documentation gives it and the rule that gives it: valid calls; natural limits (0 or 1 outside a continuous support, the top of a discrete support, a count truncated toward zero, the Discrete Uniform at a real X and outside its support); values outside the Double range (a density and a Gamma quantile returned as 0, a cumulative probability rounded to 1, the same tail kept by `Survival`); calls refused on purpose with `#NUM!` (a probability of 0, 1 or 95 in an inverse, a zero standard deviation, a negative shape, bounds in the wrong order, a negative count, a count above the trials, a success probability above 1, an unbounded density, the logarithm of a zero mass, degrees of freedom above 1E8, a location of 1E100, a mean beyond the largest Double); and what Excel does to an argument (a text gives `#VALUE!`, an empty cell is read as 0). For each call: the call as written, what it returns, the documented outcome, a check that reads As documented (green) or UNEXPECTED (red), the Status and the cause with its correction. An error shown on purpose reads As documented; only a call that departs from its documented outcome reads UNEXPECTED.
- **Not shown.** Non-convergence returns `#NUM!` with a Status that says so, but no input of the library's survey of bad input produces it. `K_STATS_NormalStandard_InverseCumulativeFast` validates nothing and is left out; so are the differences between functions recorded as not settled in #418, and more than one wrong argument at a time.

`ci/Test-Diagnostics.ps1` checks in Excel every call of the case table against its documented outcome (a number within 1E-12 relative of an mpmath reference, or the documented error), its green check, its Status (empty for a number; for an error, a fragment of the library's message where the case names one), and the call as written entered again in a scratch cell. It types eight trials into the try-it area with events on: the default, a probability of 1, a negative shape, a text, an empty X, an unbounded density, an impossible count of four arguments and a discrete quantile; for each, the example the function loads, the value, its kind, the Status, the meaning, the arguments Excel converted, the formula beside the value and the chart. Then the default sweep refused at 0 and 1 only, five invalid inputs, the Excel state and the navigation after them, and the reset. Its report, `diagnostics-report.json`, records what that Excel returned for every call, with its build. The unit tests check that every rule a case cites is a section of the manual, that every fragment of a Status is in the library's messages, and recompute the numbers with SciPy and closed forms.
