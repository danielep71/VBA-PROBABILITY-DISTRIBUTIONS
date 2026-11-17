# Distribution Explorer

Development implementation for #254, dependent on the shared shell in #253.
The selected user workbook remains the presentation baseline. This first
Explorer exposes Normal (mean and standard deviation), Exponential (rate,
not scale), and Poisson (expected count). Other families and operations are
not offered in this view. The capability map is checked against the public
functions in the target source tree.

## Build and review

1. Run `python tools/demobuild/build_demo.py --explorer`.
2. In desktop Windows Excel, run
   `powershell -File tools/demobuild/Build-Explorer.ps1 -PackageDirectory test-results/demo`.
3. Run `ci/Test-Explorer.ps1` and `ci/Test-DemoPackage.ps1` with the same
   `-PackageDirectory`. The `Distribution Explorer` workflow performs these
   steps and retains the workbook, PDF preview and JSON evidence.

The optional variant embeds the Explorer callbacks; the default shell build
retains its original Ribbon and VBA behavior. The native builder uses Excel
to create formulas, validation and charts, copying the shell's styles and
retaining its navigation. It does not import VBA through Trust Center or
change production numerical modules. Final native workbook bytes are not
deterministic: bind acceptance to the output digest in the test report, not
to the initial shell manifest. That manifest describes the pre-Excel package.

## Interaction and interpretation

Inputs are D6:D11. Changing them recalculates the sheet and refreshes its two
charts. Unselected-family parameters are retained and ignored. Formula cells
are protected; select them to inspect their visible formulas in the formula
bar. Reset on the Ribbon asks for confirmation and restores only these six
cells to Normal(0,1), rate 1, Poisson mean 4 and probe x=0. Cancelling leaves
the inputs unchanged. The non-interactive reset entry point is used by CI.

The plot uses a bounded educational domain: Normal mean +/- 4 standard
deviations, Exponential 0 to 8/rate, and Poisson integer counts through
ceil(mean + 8 sqrt(mean) + 10). The right-tail probability omitted is shown
using the direct survival function. Normal also omits its left tail. These
are illustrative domains, not quantile guarantees. Continuous curves use a
regular grid; Poisson uses integer bars and unconnected cumulative points.
Invalid or pasted out-of-range inputs display `Check inputs` and suppress
numerical output with `#N/A`. These chart limits are narrower than the API.

Input bounds: Normal mean +/- 1,000,000, standard deviation 0.001 to 10,000;
Exponential rate 0.001 to 1,000; Poisson mean 0.01 to 100; probe x +/-
10,000,000. Fractional Poisson probe values follow the API's truncation rule.
Macros are required for editing/recalculation. With macros disabled, the
saved results and navigation remain available; changing inputs then does not
produce a validated updated chart.

## When macros are not running

Every result on the sheet is a call to the embedded VBA, and the charts follow
the inputs through a workbook event. With macros off nothing is live: the file
opens showing the values it was saved with, and the first edit turns the
results into `#NAME?` while the charts stay as they were. A file that came by
download or mail opens this way until it is unblocked (file Properties,
Unblock) or its content is enabled.

The sheet states this itself. The Macros cell (G6) calls one library function
and reads `Enabled` or `DISABLED`; a red line above the tables appears when it
is disabled. The cell depends on the six inputs, so it is evaluated again at
the first edit, which is when the results break; it is not volatile, so
opening the workbook does not mark it as changed.

`Workbook_Open` recalculates the sheet and redraws the charts, keeping the
Saved flag, so a file saved with macros off, or saved while its charts showed
another family, is right as soon as macros run. An edit redraws the charts
only when it touches the inputs D6:D11, and a failure to redraw is reported in
a message instead of stopping in the VBA debugger. Grid rows beyond a Poisson
domain show `#N/A`, not `#VALUE!`. Chart axes use the General number format.

`ci/Test-ExplorerStates.ps1` checks these states in desktop Excel: every
family through single edits with events on, the statement with macros off, and
the recovery of a file saved in that state.

## Evidence and remaining acceptance

`ci/Test-ExplorerUI.ps1` exercises the real Ribbon through Windows UI
Automation, scoped to the HWND/process of a newly created Excel instance.
It invokes Recalculate, Home and every Go to menu destination, toggles the navigation bar, and handles
the real reset dialog twice: Cancel must preserve all inputs; OK must restore
the defaults. Dialog selection checks the owned process and exact prompt.
Excel exposes its dialog buttons as native child controls on the tested host;
the adapter uses their matched dialog HWND and button IDs when InvokePattern
is unavailable. No global mouse or keyboard input is sent. A runtime failure remains a
failure; callback-only tests are not substituted for this UI evidence.

`explorer_references.py` regenerates `explorer-references.json` using Python
standard-library closed forms: Normal erfc, Exponential exp/expm1 and the
finite defining Poisson sum. No expected value comes from the VBA being
tested. CI compares PDF/PMF and CDF values with absolute tolerance 1e-12,
changes inputs, checks chart values/domain/type, rejects an invalid sigma,
and verifies reset isolation. This is demonstration coverage, not a new
accuracy certification.

The native Ribbon and reset-dialog checks passed on source
`d46c1f6d8e7f2646c6b547e6746fbf698f6b7479` in
[run 44137767519](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/actions/runs/44137767519),
Excel 16.0 build 20710, x64. All six inputs were changed before each dialog:
Cancel preserved them all; OK restored every default. Every Go to destination,
Home, Recalculate, navigation toggle, reference cases and save/reopen passed.
The input workbook SHA-256 was
`93fa6263c1b54cee262ca5ecab3e763ad0477426e39b9e85eb6af2f770ccda9d`.

This is native UI automation evidence, not a human usability review.
Visual review of every family, missing-sheet callback handling and the shared
shell's remaining acceptance are still open. The Normal print preview was
reviewed separately. Do not infer compatibility with untested Office
architectures or close #254 solely from package checks.
