# Excel comparison

Measures Excel's native statistical functions and this library **against the
same 50-digit references**, so the README's accuracy claim can be shown rather
than asserted.

## What is and is not being claimed

Excel's statistical functions are good across the range most users work in.
This is **not** a list of Excel bugs, and the grid is not a highlight reel: it
deliberately includes points where Excel is expected to do just as well
(`NORM.S.DIST(-8, TRUE)` reads the lower tail directly; `F.DIST.RT(1, 1E6, 1E6)`
sits at the median). A comparison that reported only the wins would not be
evidence, and the analyzer prints a warning if any point favours Excel so it
cannot be quietly dropped.

The points where a difference is expected are structural, not accidental:

- **No direct survival function.** Excel has no standard-normal survival, so the
  upper tail must be written `1 - NORM.S.DIST(z, TRUE)`. That subtraction
  destroys the tail: by z = 8 most of the significant digits are gone, and by
  z = 15 the result is exactly zero. This library computes the tail directly.
- **Large shape or df**, where the underlying series and continued fractions
  lose accuracy.
- **Inputs where one side errors and the other returns a value.**

## Grid format

The grid is **pipe-delimited**, not comma-delimited. Every interesting field
here contains commas — `normal survival, z = 8`, `CHISQ.DIST(1E6,1E6,TRUE)` —
and the VBA exporter splits on the delimiter without quote handling, so a comma
delimiter silently wrote observations over the reference column. A delimiter
that cannot occur in the data is simpler and safer than teaching VBA to parse
quoted CSV; the generator asserts no field contains a pipe.

## Method

Both columns are evaluated through `Application.Evaluate`, so each is exercised
through the worksheet layer exactly as a user would write it — the library is
not given an in-VBA advantage over the native functions. Values are written as
a hi;lo pair to preserve full Double precision through the CSV.

## Files

| file | role |
| --- | --- |
| `generate_excel_comparison.py` | writes the grid with 50-digit references |
| `excel_comparison_grid.csv` | the grid; both observed columns filled by the macro |
| `M_STATS_PROBDIST_XLCMP.bas` | `Export_ExcelComparison` fills both columns |
| `analyze_excel_comparison.py` | prints the comparison and the README table |

## Procedure

1. `python generate_excel_comparison.py`
2. Import `M_STATS_PROBDIST_XLCMP.bas`, run `Export_ExcelComparison`
3. `python analyze_excel_comparison.py`
4. Paste the emitted table into the root README
