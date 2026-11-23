# The Ribbon of the demo workbook

The workbook has one Ribbon tab, **Probability Distributions** (keytip `PD`).
Its source is `src/workbook/customUI14.xml`. The build writes that file into the
workbook, and `src/workbook/PROBDIST_Template.xlsx` holds the same part.

The callbacks are in two modules of `src/workbook/`:

- `M_STATS_PROBDIST_EXPLORER.bas`: `ExplorerAction` and `ExplorerEnabled`;
- `M_STATS_PROBDIST_DEMO.bas`: `RibbonOnLoad`, `RibbonGoTo`, `RibbonToggle`
  and `RibbonGetPressed`.

The callbacks hold no numerical logic. They act only on this workbook, and
never on another open workbook.

## Controls

| Control (id) | Group | Kind | Callbacks | Enabled | What it does |
| --- | --- | --- | --- | --- | --- |
| Recalculate (`pdRecalculate`), keytip `R` | Calculation | button | `ExplorerAction`, `ExplorerEnabled` | when the active sheet is built; on Home and README, when the Explorer is built | Recalculates every sheet of this workbook, then refreshes the charts and the texts of each built sheet. No other workbook is recalculated. |
| Reset inputs (`pdResetInputs`), keytip `A` | Calculation | button | `ExplorerAction`, `ExplorerEnabled` | on the eight sheets with inputs, when that sheet is built; never on Home and README | Puts the inputs of the active sheet back to its example after a confirmation (OK or Cancel). Cancel changes nothing, and the other sheets are not changed. |
| Go to (`pdGo`), keytip `G` | Navigate | split button | `RibbonGoTo` | always | The button opens Home. The arrow opens the menu below. |
| Go to menu (`pdGoMenu`) | Navigate | menu | `RibbonGoTo` | always | One entry per sheet, under section separators, in the order of `examples/demo/sheets.json`. |
| Navigation bar (`pdNavigationBar`), keytip `B` | Navigate | toggle button | `RibbonToggle`, `RibbonGetPressed`, `ExplorerEnabled` | as Recalculate | Hides or shows column A and the navigation bar on every sheet, for this session. The workbook is always saved with the bar shown. |

`RibbonOnLoad` keeps the Ribbon object, so `RefreshRibbon` can ask Excel to read
the states again. `Workbook_SheetActivate` calls it whenever another sheet
becomes active, so Reset inputs follows the active sheet. After a reset of the
VBA state the Ribbon object is lost: the tab keeps the states it shows until the
workbook is reopened.

## Go to entries

Each entry opens a sheet by its CodeName (`RibbonSheet`), not by its tab name.
A renamed tab is still reached. An entry whose sheet is missing shows the message
*This workbook has no sheet for the Ribbon entry …* and changes nothing.

| Section | Entry (id) | Sheet (CodeName) |
| --- | --- | --- |
| Start | Home (`pdGoHome`) | `shHome` |
| Explore | Explorer (`pdGoExplorer`) | `shExplorer` |
| Explore | Probabilities (`pdGoProbability`) | `shProbability` |
| Explore | Parameters (`pdGoParameters`) | `shParameters` |
| Accuracy | Comparison (`pdGoComparison`) | `shComparison` |
| Accuracy | Tails (`pdGoTails`) | `shTails` |
| Applications | Discrete (`pdGoDiscrete`) | `shDiscrete` |
| Applications | Risk (`pdGoRisk`) | `shRisk` |
| Reference | Diagnostics (`pdGoDiagnostics`) | `shDiagnostics` |
| Reference | README (`pdGoReadme`) | `shReadme` |

The bar on the left of every sheet and the links on Home use the same map. They
target defined names (`nav_<CodeName>`), so they work with macros disabled. The
Ribbon commands need macros.

## Reset inputs, sheet by sheet

Each sheet's module asks its own question and restores only its own inputs:

| Sheet | Confirmation |
| --- | --- |
| Explorer | Reset only Explorer inputs to Normal(0,1)? |
| Probabilities | Reset only the Probabilities inputs to Normal(0,1)? |
| Parameters | Reset only the Parameter Effects inputs to Normal(0,1)? |
| Comparison | Reset only the Comparison inputs to the first tail case and the study's grades? |
| Tails | Reset only the Tails inputs to the Standard Normal scenario and the study's grades? |
| Discrete | Reset only the Discrete inputs to the default example (100 loans, default probability 2%)? |
| Risk | Reset only the Risk inputs to the default example (1,000,000 over 10 days at 99%, 100 loans)? |
| Diagnostics | Reset only the Diagnostics inputs to the default example (the Gamma quantile at 0.95)? |

Each reset turns Excel's events off while it writes, and turns them back on, on
success and on error.

## How it is checked

`ci/Test-ExplorerUI.ps1` drives the real Ribbon through Windows UI Automation, in
an Excel process it owns, with no global keyboard or mouse input. It checks:

- Recalculate;
- the Navigation bar toggle, off and on;
- Go to, and every entry of the menu;
- Reset inputs disabled on Home and README and enabled on Probabilities and the
  Explorer;
- the Explorer's confirmation dialog, Cancel then OK;
- a renamed tab, still reached;
- a deleted sheet: the message is shown and the active sheet does not change.

The renamed tab and the deleted sheet are changed only in a copy opened
read-only, which is never saved. `ci/Test-DemoPackage.ps1` follows every Home
link with macros disabled, and again after a tab is renamed. Each sheet's own
test checks its reset.

The Reset dialogs of the other seven sheets are checked through their modules,
not by clicking the Ribbon.
