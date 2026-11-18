# Source

The VBA source of the library and of the Explorer workbook. The layout
follows the K-SACCR repository: the library modules, and a `workbook/`
folder with what one workbook needs besides them.

```text
src/
├─ M_STATS_PROBDIST_CORE.bas           # the six library modules: import these
├─ M_STATS_PROBDIST_SPECIALFUNCS.bas   #   into any workbook to use the
├─ M_STATS_PROBDIST_NORMALFAMILY.bas   #   K_STATS_ worksheet functions
├─ M_STATS_PROBDIST_TFAMILY.bas
├─ M_STATS_PROBDIST_CONTINUOUS.bas
├─ M_STATS_PROBDIST_DISCRETE.bas
└─ workbook/                           # the Explorer workbook
   ├─ PROBDIST_Template.xlsx           # its sheets, formulas, names, validation, charts; no code
   ├─ customUI14.xml                   # its Ribbon
   ├─ ThisWorkbook.cls                 # its workbook events
   ├─ M_STATS_PROBDIST_DEMO.bas        # navigation bar and Ribbon callbacks
   ├─ M_STATS_PROBDIST_EXPLORER.bas    # the Explorer's inputs and charts
   └─ M_STATS_PROBDIST_REGISTER.bas    # Function Wizard descriptions
```

## The library

The six modules of `src/` are the library. They are what the
[installation guide](../INSTALLATION.md) imports, what the release bundle
ships, and what the Excel regression and the accuracy evidence bind
(`.github/excel-evidence-policy.json`, `benchmark/observation_manifest.json`).
The worksheet functions are listed in [docs/PUBLIC_API.txt](../docs/PUBLIC_API.txt).

## The workbook

`workbook/` holds the Explorer workbook (`examples/demo/EXPLORER.md`). Its
modules compute nothing: every number on the sheets is a formula that calls
the library. They are not part of the library, so no regression or evidence
workbook imports them and no evidence binding records them; a public
`K_STATS_` function may not be declared here (`tools/check_vba_source.py`).

- `PROBDIST_Template.xlsx` is the workbook without its VBA project, made by
  `tools/demobuild/explorer_template.py` from a workbook that the
  `Distribution Explorer` workflow verified, and pinned by its SHA-256 in
  `.github/binary-allowlist.json`. Update it when the layout of the sheets
  changes; a change to the code needs no new template.
- `customUI14.xml` is the Ribbon. The build writes it into the workbook, and
  the template holds the same part (`tests/test_demo_explorer.py`).
- `ThisWorkbook.cls` is the code of the ThisWorkbook module, exported as the
  VBE exports it. The sheet modules hold no code, so they have no file.
  A document module keeps the name Excel gives it (ThisWorkbook, or a sheet
  code name starting with `sh`) and lives only here.

## Two ways to a workbook

- `python tools/demobuild/explorer_build.py` puts a VBA project built from
  these sources into the template, without Excel; `Evolve-Explorer.ps1` and
  the checks of the `Distribution Explorer` workflow follow
  (`examples/demo/EXPLORER.md`).
- `tools/Sync-PROBDIST-VBA.ps1` imports these sources into a workbook in
  desktop Excel: into a new workbook saved from the template
  (`-FromTemplate`), or into a workbook you already have. Double-click
  `tools/Sync-PROBDIST-VBA.cmd` to build
  `test-results/Probability-Distributions-Development.xlsm`
  ([tools/README.md](../tools/README.md#vba-synchronization)).
