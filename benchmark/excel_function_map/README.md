# #330: from Excel's distribution functions to the library's

The table of [docs/PUBLIC_API.md](../../docs/PUBLIC_API.md#coming-from-excels-functions)
is generated here, from a list of functions and from what Excel returned.

| File | What it is |
| --- | --- |
| `function_map.py` | The source: each Excel function, the library expression that replaces it, what differs, one example; the probes; the two lists of functions without a counterpart |
| `references.py` | The value of a library expression from 60-digit arithmetic, written from the definitions with mpmath and independent of the VBA |
| `build_grid.py` | Writes `excel_function_map_grid.csv`: one row per example and probe, with its reference |
| `M_STATS_PROBDIST_XFM.bas` | Evaluates both formulas of every row in Excel, as a cell would, and records each result as the bits of its Double (study export `xfm_examples`) |
| `render_map.py` | Writes the section of `docs/PUBLIC_API.md` between its two markers; `--check` fails when the page is not current |
| `test_map.py` | The map against the public surface, the references against known values, the renderer on made-up observations |
| `floor_build.py` | #493: writes `excel_floor_grid.csv`, the library call and every Excel formulation at each comparable point of the main grid and the two holdouts; `--check` fails when the committed grid is not current |
| `floor_check.py` | #493: judges the exported floor grid against the rule of `CONTRIBUTING.md` ("Excel as the floor") and writes `floor_results.json` |
| `floor_dispositions.json` | #493: the inferior points not fixed now, each against an issue, and the functions excluded by decision; `floor_check.py` fails on an inferior point without one and on a disposition that is no longer needed |
| `test_floor.py` | The floor tools on made-up rows, and the committed grid against its generator |

## What is checked and what is not

- **Checked for every example:** the library call exists in
  `docs/PUBLIC_API.txt` and takes that many arguments; Excel evaluated it; its
  value agrees with the reference within 1E-12 relative, or the page is not
  written.
- **Measured, not asserted:** whether Excel's own function returns the same
  number, and what either side does with a fractional degree of freedom or
  count. The page says what was returned on Excel 16.0 build 20710, 64-bit;
  another build may differ.
- **Not checked:** the legacy names in the "also" lines. They are listed from
  Microsoft's documentation and were not evaluated. Nor is each "What
  differs" note tested beyond its one example.
- **One example per function** shows the mapping, not the accuracy. The
  accuracy evidence is the main grid and its contracts; where Excel loses a
  tail, `../excel_comparison/` measures it.

## Changing the map

Edit `function_map.py`, delete the grid, run `build_grid.py --write`, export
`xfm_examples` in Excel, apply the export, then `render_map.py --write`.
