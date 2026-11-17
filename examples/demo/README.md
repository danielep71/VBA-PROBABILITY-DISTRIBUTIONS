# Demo workbook shell

Development work for #253. The workbook is not a release asset and has not
been validated in Excel. The thematic sheets are explicitly marked as under
development; #254–#261 supply their content. No library kernel is changed.

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
external dependencies or p-code. Excel compilation of this new package remains
unverified; #110's certified CI-workbook execution is separate work.

`examples/demo/sheets.json` is the single map for tab order, document CodeNames,
Ribbon destinations, defined-name hyperlinks and current-sheet highlighting.
The workbook embeds the current production modules plus `M_STATS_PROBDIST_DEMO.bas`.
Start Here shows the source commit and the library version separately from the
development workbook version. The manifest is an inventory, not certification.

## Reference and migration

Reference inspected: K-SACCR commit
`963695e2a862647d9013814b5bc63df3489c387e`, especially
`src/workbook/SACCR_Template.xlsx`, `customUI14.xml`, `M_Ribbon.bas` and
`docs/workbook/RIBBON.md`. This implementation is original; it adapts the
interaction contract and visual conventions, not the SA-CCR commands or code.

| Reference element inspected | Adaptation |
| --- | --- |
| Aptos typography, blue `265A99` title band | Aptos body and blue title/section bands |
| Home navigation at left, narrow separator | Persistent left links on every sheet, teal current destination |
| Home label and description columns | Labels in C, wrapped guidance in D |
| Input/output/reference conventions | Yellow input legend; explicit result/reference sections |
| Header panes and hidden gridlines | Freeze first three rows and navigation columns; hide gridlines |
| Large Home description column | Narrower wrapped description column to fit the demo |
| Print layout | Landscape A4, one-page fit requested |
| Ribbon CodeName navigation | Callback resolves CodeName only inside ThisWorkbook |
| Macro-disabled navigation | Defined-name links and visible tabs |
| Reset and calculation commands | Workbook-only recalculation; reset disabled until input contracts exist |
| Numerical tables, formats and charts | Deferred to the thematic sheets, not invented by the shell |

The old `examples/STATS-Distributions demo.xlsm` was removed intentionally by
commit `6d919e97765b79085bfb6b2ffb23d0a8b753127f` for v1.0.0: its embedded VBA
was not checked against the released modules. Its historical workbook was
inspected: HOME, SETUP, NORMAL, LOGNORMAL and BENCHMARK_NORMALFAMILY. The new
shell replaces that entry point rather than restoring uncertified embedded
code. Relevant examples can be reviewed separately in the thematic issues.

## Controls and ownership

- **Go to / Help:** `DemoGoTo`, with a sheet CodeName in the control tag.
  Missing targets produce a clear message. Renamed tabs do not change CodeNames.
- **Recalculate:** `DemoRecalculate`, worksheet calculation in ThisWorkbook only.
- **Reset inputs:** statically disabled. There are no editable example inputs
  in the shell. Future implementations must define bounded input ranges,
  confirmation/cancellation and restoration of any owned Excel state.
- No application state, timers, shortcuts or Ribbon pointers are retained.
- Sheets are protected against unintended edits, without a password. This is
  an editing aid, not a security boundary. All cells are currently locked.
- Defined names target sheet cells, so Excel can update the destinations on
  tab rename. Names provide macro-disabled navigation independently of VBA.

## Acceptance still required in desktop Excel

Record the exact commit, workbook digest, Excel build and Office bitness.

1. Open without repair messages and compile the embedded project.
2. Exercise every Ribbon target, Help and worksheet link.
3. Rename a tab, then repeat Ribbon and hyperlink navigation.
4. In a disposable copy remove a sheet: its Ribbon target must report failure
   without navigating to another workbook.
5. Open a second workbook and confirm all actions stay in the demo.
6. Open with macros disabled: all tabs and left links must remain usable.
7. Save, close and reopen: verify RibbonX, hyperlinks and protection persist.
8. Inspect every sheet at normal zoom and print preview for clipped text.
9. Record tested environments only; do not infer 32-bit compatibility.

These checks are pending. The shell does not close #253. Integrated demo
validation remains #262; screenshots remain #250.
