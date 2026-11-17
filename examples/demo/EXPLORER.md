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

## Evidence and remaining acceptance

`explorer_references.py` regenerates `explorer-references.json` using Python
standard-library closed forms: Normal erfc, Exponential exp/expm1 and the
finite defining Poisson sum. No expected value comes from the VBA being
tested. CI compares PDF/PMF and CDF values with absolute tolerance 1e-12,
changes inputs, checks chart values/domain/type, rejects an invalid sigma,
and verifies reset isolation. This is demonstration coverage, not a new
accuracy certification.

Actual Ribbon clicks, reset cancellation, screenshots and visual/print
acceptance remain distinct from automation. Do not infer compatibility with
untested Office architectures or close #254 solely from package checks.
