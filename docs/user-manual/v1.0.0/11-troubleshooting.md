# 11. Troubleshooting

## 11.1 Installation problems

| Symptom | What to check |
|---|---|
| Cell shows `#NAME?` | Modules missing from this workbook, or macros disabled |
| Compile error, missing procedure | Not all six modules imported, or versions mixed |
| "Ambiguous name" on compile | A module imported twice, or an old copy remains |
| Functions lost after saving | Workbook saved as `.xlsx`, which does not store macros |
| Security warning on opening | Source origin, hash, trusted location, macro settings |

- `#NAME?` comes from Excel: it found no function with that name. Check the spelling against chapter 13 and check that the six modules are in the VBA project of the same workbook.
- For a compile error or an ambiguous name, remove all `M_STATS_PROBDIST_` modules and import the six files again from one release (section 4.2). Do not paste new source beside old components.

## 11.2 A function returns an error value

Section 9.2 lists the causes of `#NUM!`, section 9.3 those of `#VALUE!`, and section 9.6 gives the steps for finding the cause from a worksheet.

If a formula is rejected as you type it, check the separators: some regional settings use a semicolon between arguments and a decimal comma.

## 11.3 A result differs from another tool

Differences are almost always caused by conventions, not by arithmetic. Check in this order.

| Check | Typical mismatch |
|---|---|
| Parameterization | Gamma scale versus rate; Exponential rate versus mean |
| Lognormal parameters | Parameters of $\ln X$ versus arithmetic moments |
| Counted quantity | Failures versus trials (Geometric, Negative Binomial) |
| Tail | $P(X \le x)$ versus $P(X > x)$ |
| Inequality | $P(X > k)$ versus $P(X \ge k)$ for counts |
| Discrete quantile | Smallest $k$ reaching the probability, or another rule |
| Argument order | Differs between tools |
| Displayed digits | The other tool may round its output |

If the conventions agree and the results still differ:

- compare more digits than the cell format shows;
- check whether the input lies in a region with a looser guarantee or a recorded limitation (chapter 10);
- make sure the other value is an independent reference and was not produced by subtracting from 1 in a far tail;
- record the library version, the inputs, both results and the tolerance you applied.

## 11.4 A small tail probability comes out as zero

You probably computed one minus a cumulative probability. Use the `Survival` function (section 7.3). If `Survival` itself returns 0, the true value is smaller than the smallest Double. For a discrete distribution, `LogPMF` still gives the size of a single probability mass.

## 11.5 A discrete result is unexpected

- **The quantile does not round-trip.** `Cumulative` applied to the result of `InverseCumulative` is usually larger than the probability you started from. This is the convention of section 7.5.
- **A non-integer count.** Counts are truncated toward zero: `7.9` is treated as `7`, not rounded to `8`. Round explicitly before you pass a computed count.
- **Off by one.** `Survival` at $k$ excludes $k$ (section 7.2). Geometric and Negative Binomial functions count failures, not trials (section 7.6).

## 11.6 The worksheet recalculates slowly

- Beta and F inverses, and Gamma, Chi-square and Poisson calculations in particular regions, take much longer per call than the other functions (section 10.5).
- Reduce the number of such cells, or compute a value once and refer to it.
- For standard normal draws in a VBA loop, consider `K_STATS_NormalStandard_InverseCumulativeFast` if its lower accuracy is acceptable (section 5.8).

## 11.7 Behaviour differs between machines

- Confirm that both machines use the same library version and the same six modules.
- Confirm the Excel version and the Office bitness. Only 64-bit Office on Windows is covered by the retained test evidence (chapter 2).
- Run the regression harness on the machine in question (section 4.5).

## 11.8 If you cannot resolve a problem

Reduce the problem to a single formula with numeric arguments in an otherwise empty workbook. Note the library version, the Excel version and bitness, the formula, the result you got and the result you expected. Do not include client, personal or confidential data in a report. Report suspected security problems privately, as described in `SECURITY.md`.
