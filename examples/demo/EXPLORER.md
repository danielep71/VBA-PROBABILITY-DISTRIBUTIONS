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
locked and the inputs are not. Reset on the Ribbon asks for confirmation and
restores Normal(0,1), probe 0, p = 0.95, 201 points and 6 decimals. Cancelling changes nothing.

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
navigation toggle and the reset dialog, Cancel and OK.

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
