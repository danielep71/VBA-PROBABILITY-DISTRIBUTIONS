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

The initial approximation was rejected in review. The builder now uses extracted
Open XML parts from that exact template: the complete style table and theme,
the Home sheet cell styles, row heights, column widths, panes and print settings,
and the original navigation button shape as a prototype. Captions, sheet targets
and project branding are adapted. The original dark sidebar and black/blue
header bands are preserved. These parts live in `reference/` and are covered
by the build manifest and template-fidelity tests.

Adapted from K-SACCR, maintained by Daniele Penza, under MPL-2.0.
Source: https://github.com/danielep71/K-SACCR/tree/963695e2a862647d9013814b5bc63df3489c387e/src/workbook
The source license is retained with the extracted parts. The layout XML contains
no SA-CCR cell content; navigation captions are replaced during generation.

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

## Visual review

The revised package was rendered to PDF with LibreOffice and every worksheet
was inspected. The navigation overlap found in that rendering was corrected.
This verifies a rendered preview, not the desktop Excel Ribbon or save cycle.
The reference retains its multi-page landscape print layout.
