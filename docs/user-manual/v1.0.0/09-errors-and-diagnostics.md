# 9. Errors and diagnostics

## 9.1 The return contract

| Situation | Result |
|---|---|
| The calculation succeeds | A number |
| The request cannot be answered | `#NUM!` |
| An unexpected run-time failure occurs | `#VALUE!` |

The functions do not raise VBA run-time errors for invalid input, do not show message boxes and, as delivered, do not write to the Excel status bar.

Invalid input is never repaired silently. A probability of 1.2 is not treated as 1, and a negative standard deviation is not replaced by its absolute value. The single exception is `K_STATS_NormalStandard_InverseCumulativeFast` (section 5.8).

## 9.2 When a function returns #NUM!

`#NUM!` means that the library understood the request and declined to answer it.

| Cause | Examples |
|---|---|
| Invalid parameter | Zero or negative scale, shape, rate or d.f. parameter |
| | Uniform bounds not in increasing order |
| | Interval with the upper bound below the lower bound |
| | `ProbSuccess` outside its allowed range (section 6.4) |
| Invalid count | Negative or impossible count (section 7.8) |
| Invalid probability | `Probability` of 0, 1 or 1.5 in an inverse function |
| Unsupported magnitude | Degrees of freedom above the limits of section 6.2 |
| | Discrete arguments above the limits of section 6.5 |
| | Location, scale or shape of size 1E100 or more, where checked |
| Result too large | A quantile or moment beyond the largest Double |
| Unbounded density | The points listed in section 6.3 |
| No convergence | An iterative calculation did not settle |
| Undefined logarithm | `LogPMF` of an outcome with probability zero |

The library does not return a substitute value in any of these cases: no largest representable number, no partial result of an iteration, and no large negative number in place of an undefined logarithm.

## 9.3 When a function returns #VALUE!

`#VALUE!` from the library means that an unexpected VBA run-time failure occurred during the calculation. It is not the normal response to bad input. If you can reproduce it with numeric arguments, treat it as a possible defect and report it with the exact formula and the library version.

Excel also shows `#VALUE!` when an argument is not a number, for example a text cell. Check the arguments first.

## 9.4 Results that are not errors

| Result | Reason |
|---|---|
| 0 for a density or PMF far in a tail | The true value is too small for a Double |
| 0 or 1 outside a continuous support | The natural limit (section 7.8) |
| 0 for a Gamma quantile | The true quantile is below the Double range |
| 1 for a cumulative probability | The upper tail is below Double resolution |

For the size of a very small probability mass, use `LogPMF`. For a very small upper tail, use `Survival`.

## 9.5 Reading the diagnostic message from VBA

When a function returns an error value, it writes a message into its `Status` argument. The message names the argument or the condition that caused the failure.

```vba
Dim Result As Variant
Dim Status As String

Result = K_STATS_Binomial_Cumulative(12#, 20#, 1.35, Status)

If IsError(Result) Then
    Debug.Print Status
End If
```

This prints `ProbSuccess must be between 0 and 1`. Other typical messages:

| Message | Meaning |
|---|---|
| `Probability must be strictly between 0 and 1` | Invalid probability in an inverse |
| `NumberSuccesses must not exceed Trials` | Count outside the Binomial support |
| `Log-mass is undefined: the outcome has probability zero` | `LogPMF` of an impossible outcome |
| `Normal quantile overflows the supported Double range` | Result too large |
| `Unexpected error in ...` | The `#VALUE!` case, with the VBA error text |

Display or log the text, but do not write code that depends on a particular wording.

## 9.6 Finding the cause from a worksheet

A worksheet formula cannot read `Status`. To find the reason for an error in a cell:

1. Check each argument against chapter 6: sign, range and order. The evaluation point comes first, then the parameters.
2. Check that probabilities are between 0 and 1 and not entered as percentages such as `5`.
3. Check the size limits in sections 6.2 and 6.5.
4. If the cause is still unclear, call the same function from a short macro and print `Status`, as in section 9.5.

Do not hide errors with `IFERROR` by default. An error from this library means that the request was outside the supported domain or could not be completed.
