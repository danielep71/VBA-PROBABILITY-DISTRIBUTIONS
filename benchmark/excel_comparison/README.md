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

## Coverage

29 cases across all 15 families. Each family contributes a body point where
Excel is expected to match exactly, and where one exists a deep-tail or
extreme-parameter point.

## The structural difference: Excel has no upper tail for most families

Only **three** of the fifteen families have a direct upper-tail function in
Excel: `T.DIST.RT`, `CHISQ.DIST.RT`, `F.DIST.RT`. For Normal, Lognormal, Gamma,
Beta, Exponential, Weibull, Binomial, Poisson, NegativeBinomial and
Hypergeometric the tail must be written `1 - CDF`, and that subtraction is what
destroys it: at z = 8 nothing correct survives, and by z = 15 the expression
returns exactly zero against a true 3.7E-51.

Uniform and DiscreteUniform have no Excel function at all, as does log-mass for
any discrete family. Those cases carry `-` as the Excel formula and record
`NONE`, because "no function exists" is the most consequential difference in
the table and omitting it would understate the case rather than overstate it.

## How the results are judged

Two columns, because "which is more accurate" and "is this fit for use" are
different questions and merging them misleads:

| column | question |
| --- | --- |
| **Closer** | which implementation is nearer the reference - often true and meaningless |
| **Fit for use** | whether each side clears an absolute bar - the column that should drive decisions |

The bar has two levels, both stated so a reader can disagree with them
explicitly rather than infer them:

| grade | correct digits | meaning |
| --- | --- | --- |
| **reference** | >= 12 | safe as a *building block*: the value can be fed into iteration, root-finding or accumulation without its error becoming visible. This is the standard the library's own accuracy contracts are written to. |
| **report** | >= 6 | safe as a *directly reported* statistic: no published p-value or critical value carries more than about six significant figures. |
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
4. Paste the emitted table into the root README
