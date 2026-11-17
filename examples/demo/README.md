# Demo workbook shell

Development work for #253. The workbook is not a release asset. Automated
Excel checks cover the operations listed below; interactive acceptance remains
incomplete. The thematic sheets are explicitly marked as under
development; #254–#261 supply their content. No library kernel is changed.

## User-supplied baseline

On 2026-11-17 the user supplied `Probability-Distributions-Demo.xlsm` and
instructed us to use it. A fresh build at commit
`2da91db696da6fa45509eeb5ed21ee1385d5886e` is byte-identical to that file.
The exact digest and comparison record are in
[reference/user-workbook.json](reference/user-workbook.json).

Preserve this workbook's presentation, sheet structure and navigation when
adding the thematic examples. The supplied file is already represented by
the committed sources; no binary replacement or reverse-engineered override
is needed. This selection does not close the outstanding runtime checks.

## Build

From the repository root, with Python 3.11 or later:

```sh
python tools/demobuild/build_demo.py
python -m unittest discover -s tests -p test_demo_shell.py -v
```

The `Demo shell package` workflow uploads a `demo-shell-unverified` artifact
on pull requests affecting the shell. Download it for the desktop acceptance
checklist below; a successful package job does not certify Excel behavior.

The builder writes `test-results/demo/Probability-Distributions-Demo.xlsm`,
its SHA-256 manifest and the generated RibbonX source there. Outputs are
ignored. It uses the existing MS-OVBA/MS-CFB writer, without Office automation,
external dependencies or p-code. The package acceptance test executes an
embedded library function in Excel; #110's certified CI-workbook execution is
separate work.

`examples/demo/sheets.json` is the single map of sections, tab names, document
CodeNames, tooltips and issues. The Go to menu of the Ribbon, the navigation
bar of every sheet, the table of links on Home and the defined names of the
link targets are all written from it. The workbook embeds the current
production modules plus `M_STATS_PROBDIST_DEMO.bas`. README, the last sheet,
shows the source commit and the library version. The manifest is an inventory,
not certification.

## Reference: the K-SACCR workbook template

Reference inspected: K-SACCR commit `400871e`, `src/workbook/SACCR_Template.xlsx`,
`customUI14.xml`, `M_Ribbon.bas`, `M_Main.bas` (navigation bar),
`tools/navigation_bar.py`, `docs/workbook/RIBBON.md` and
`docs/workbook/sheet-style.json`. The demo takes these from it:

| Element | In the demo |
| --- | --- |
| Style table and theme | The template's own `xl/styles.xml` and `xl/theme/theme1.xml`, kept in `reference/` and written unchanged. |
| Page | Black band with the project name, blue band with the sheet title, subtitle, section headings and bordered two-column tables: the template's cell styles, row heights and column widths. Rows 1 to 3 and column A frozen, no gridlines, landscape print setup. |
| Navigation bar | Column A of every sheet: linked shapes in the sections of the Go to menu, the first section after Start in the frozen rows, the others below with Start last. The current sheet is highlighted and not linked; each button has a strip in its sheet's tab colour; the shapes are not printed. Same sizes, colours and shape names (`NavBarTop`, `NavBar`). |
| Ribbon | One tab with a keytip; a Navigate group with a Go to split button (the button opens Home, the menu lists every sheet under section separators, with tooltips) and the Navigation bar toggle. Callbacks named as in the template: `RibbonOnLoad`, `RibbonAction`, `RibbonGoTo`, `RibbonGetEnabled`, `RibbonToggle`, `RibbonGetPressed`, with `RibbonSheet` mapping an entry to a CodeName. |
| Home | First tab, last entry of the bar. Three tables: what the workbook shows, how to use it, and a link to every sheet with its section. |
| Tab names and colours | Plain names without numbers; one colour per role, as the template's sheet-style contract gives them. |
| Protection | Every sheet protected without a password; workbook structure locked. |
| Saving | The workbook is always saved with the bar shown (`Workbook_BeforeSave`, `Workbook_AfterSave`). |

Not taken, because the demo has nothing for them yet: the Calculation split
button with its menu, CSV import, evidence export, the Status labels, and the
stored Ribbon pointer that lets the template refresh its tab after a reset of
the VBA state. After such a reset the demo's Navigation bar toggle keeps
working, but its pressed state is refreshed only when the workbook is reopened.

Adapted from K-SACCR, maintained by Daniele Penza, under MPL-2.0.
Source: https://github.com/danielep71/K-SACCR/tree/400871e/src/workbook
The source license is retained with the extracted parts. No SA-CCR cell
content, command or code is included.

The old `examples/STATS-Distributions demo.xlsm` was removed intentionally by
commit `6d919e97765b79085bfb6b2ffb23d0a8b753127f` for v1.0.0: its embedded VBA
was not checked against the released modules. The new shell replaces that
entry point rather than restoring uncertified embedded code.

## Controls and ownership

- **Go to:** `RibbonGoTo` activates a sheet of this workbook by CodeName. A
  missing target produces a clear message.
- **Recalculate:** `RibbonAction` calls `RecalculateWorkbook`: worksheet
  calculation in ThisWorkbook only.
- **Reset inputs:** disabled through `RibbonGetEnabled`. There are no editable
  example inputs in the shell. Future implementations must define bounded
  input ranges, confirmation and restoration of any owned Excel state.
- **Navigation bar:** `RibbonToggle` hides or shows column A and the two shape
  groups on every sheet, for the session. The state is read from the workbook
  (column A of the first sheet), not from a variable.
- The module keeps the Ribbon object and one flag for a save in progress.
  Screen updating and the Saved flag are put back on success and on failure.
  No timer or shortcut is set.
- Sheets are protected against unintended edits, without a password, and the
  structure is locked so that tabs cannot be renamed under the links. This is
  an editing aid, not a security boundary. All cells are currently locked.
- The links of the bar and of Home target defined names (`nav_<CodeName>`), so
  they work with macros disabled.

## Desktop Excel acceptance

Record the exact commit, workbook digest, Excel build and Office bitness.

1. Open without repair messages and compile the embedded project.
2. Exercise Recalculate, every Go to entry and every link of the bar and of Home.
3. Press Navigation bar: column A and the bar disappear on every sheet; press
   again: they return. Hide the bar, save, close and reopen: the bar is shown.
4. Open a second workbook and confirm all actions stay in the demo.
5. Open with macros disabled: the bar and the Home links must remain usable.
6. Save, close and reopen: verify RibbonX, links and protection persist.
7. Inspect every sheet at normal zoom and print preview for clipped text; the
   bar must not print.
8. Record tested environments only; do not infer 32-bit compatibility.

The automated test covers opening, embedded VBA execution, mapped targets,
navigation visibility, two-workbook isolation and save events. Actual Ribbon
and hyperlink clicks, renamed/missing targets, visual/print review and
screenshots remain pending. The shell does not close #253. Integrated demo
validation remains #262; screenshots remain #250.

## Visual review

The package was rendered to PDF with LibreOffice and compared, sheet by sheet,
with the same rendering of the template. This verifies a rendered preview, not
the desktop Excel Ribbon, the toggle or the save cycle.

## Excel-open check

Every sheet and drawing is written as plain Open XML, with no optional
revision or compatibility metadata of its own; the template's style table
keeps its `mc:Ignorable` list, whose prefixes it declares. Package tests check
namespace-valued compatibility attributes in every part. The `excel-open` job
opens the XLSM with macros disabled, saves a copy, and reopens it using desktop
Excel. The evidence records Excel version/build and input hash. This open test
alone does not validate Ribbon callbacks or numerical execution; the functional
phase below additionally exercises the embedded VBA.

## Automated functional acceptance

The Excel package test also executes an embedded library function, checks every
Ribbon destination and defined-name target against the sheet map, hides the
navigation groups while a second workbook is active, verifies protection and
state restoration, and exercises the workbook save events. Reopening with macros
disabled must show the navigation again. This checks the production routines
behind the controls; actual Ribbon clicks and interactive visual acceptance
remain distinct manual checks. The original artifact is tested in a disposable
runner directory and the input SHA-256 is retained in the report.
