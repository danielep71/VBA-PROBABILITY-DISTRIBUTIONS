# Excel comparison

Measures Excel's native statistical functions and this library **against the
same 50-digit references**, so the README's accuracy claim can be shown rather
than asserted.

## What is and is not being claimed

This is a comparison of 29 specific formulas and recorded observations, not an exhaustive assessment of Excel or a current-source certification. Body, tail and extreme-parameter cases are included; results favouring Excel remain visible.

The `1 - CDF` tail expressions can lose significance or return zero. They are not always the best Excel formulations: normal symmetry gives `NORM.S.DIST(-z,TRUE)`, and Exponential and Weibull survival probabilities have elementary formulas. The library offers a consistent direct-tail API; this study does not benchmark every possible Excel alternative.

## Grid format

The grid is **pipe-delimited**, not comma-delimited. Every interesting field
here contains commas — `normal survival, z = 8`, `CHISQ.DIST(1E6,1E6,TRUE)` —
and the VBA exporter splits on the delimiter without quote handling, so a comma
delimiter silently wrote observations over the reference column. A delimiter
that cannot occur in the data is simpler and safer than teaching VBA to parse
quoted CSV; the generator asserts no field contains a pipe.

## Coverage

29 cases across all 15 families. Each family contributes a body point where
Excel is expected to match exactly, and where one exists a deep-tail or
extreme-parameter point.

## Scope of the formula comparison

The grid uses direct right-tail Excel functions where listed and `1 - CDF` for several other cases. Rows marked `NONE` mean no dedicated counterpart was included for that surface; they do not rule out an equivalent formula built from other Excel functions.

At the 2026-09-30 audit, rerunning the analyzer on the committed observations gives 15 cases with both sides at reference grade, 12 with an absent/unusable listed Excel formula and a usable library result, and 2 with both usable but at different grades (Beta survival and large-df Student t survival). This reproduces existing measurements; it is not a new Excel run.

## How the results are judged

Two columns, because "which is more accurate" and "is this fit for use" are
different questions and merging them misleads:

| column | question |
| --- | --- |
| **Closer** | which implementation is nearer the reference - often true and meaningless |
| **Fit for use** | whether each side clears the study's stated digit threshold |

The bar has two levels, both stated so a reader can disagree with them
explicitly rather than infer them:

| grade | correct digits | meaning |
| --- | --- | --- |
| **reference** | >= 12 | study threshold of at least 12 correct significant digits; downstream error still depends on conditioning and use. |
| **report** | >= 6 | study threshold of at least 6 correct significant digits; suitability depends on the application. |
| **inadequate** | < 6 | could change what a user reports. |
| **wrong** | 0 | not an approximation of the answer at all. |

Being "closer" while both sides are reference grade is a fact about the two
implementations, not a reason to prefer either.

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
4. Reconcile documentation against the emitted table and preserve all outcomes.

The generator rewrites the comparison grid. Use a disposable checkout when reproducing references; preserve the committed observations until a real replacement export is available.
