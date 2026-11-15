# 5. Calling the functions from worksheets and VBA

## 5.1 How the functions are named

A function name has three parts: the prefix `K_STATS_`, the distribution and the kind of result, as in `K_STATS_Gamma_Survival`.

| Kind | Returns |
|---|---|
| `Density` | Density of a continuous distribution |
| `PMF` | Probability mass of a discrete distribution |
| `LogPMF` | Natural logarithm of the probability mass |
| `Cumulative` | $P(X \le x)$ |
| `Survival` | $P(X > x)$ |
| `InverseCumulative` | Quantile for a cumulative probability |
| `InverseSurvival` | Quantile for an upper-tail probability |
| `Mean`, `Variance`, `StdDev` | Moments |

Chapter 13 shows which kinds exist for each distribution.

## 5.2 Argument order

1. The evaluation point comes first: a value, a count, or a `Probability` for an inverse function.
2. The distribution parameters follow.
3. An optional `Status` argument comes last (section 5.6).

Moment functions take the parameters only.

```text
=K_STATS_Gamma_Cumulative(4.5,3,2)
=K_STATS_Gamma_InverseCumulative(0.99,3,2)
=K_STATS_Gamma_Mean(3,2)
```

Here `3` is the shape and `2` the scale.

## 5.3 Optional parameters

Only the seven general Normal functions (`K_STATS_Normal_Cumulative` and its siblings) have optional parameters: `Mean` defaults to 0 and `StdDev` to 1. These two formulas return the same value:

```text
=K_STATS_Normal_Cumulative(1.5)
=K_STATS_Normal_Cumulative(1.5,0,1)
```

Every other distribution parameter is required.

## 5.4 Cell references

Arguments can be numbers, cell references or other formulas. Each argument is a single number. To evaluate a column of values, fill the formula down.

```text
=2*K_STATS_StudentT_Survival(ABS(B2),C2)
```

## 5.5 The function that returns two values

`K_STATS_Lognormal_ParametersFromMeanStdDev` converts the arithmetic mean and standard deviation of a lognormal variable into the parameters used by the other Lognormal functions. It returns one row of two values: `MeanLog`, then `StdDevLog`. Pick each with `INDEX`:

```text
=INDEX(K_STATS_Lognormal_ParametersFromMeanStdDev(100,30),1,1)
=INDEX(K_STATS_Lognormal_ParametersFromMeanStdDev(100,30),1,2)
```

Section 8.5 uses it in a complete example.

## 5.6 The Status argument

When a function fails, it writes a short explanation into `Status`, for example `Probability must be strictly between 0 and 1`. On success, `Status` is empty. `Status` is useful from VBA; in a worksheet formula you normally leave it out.

## 5.7 Calling from VBA

Call the functions by name from any module in the same project. They return a `Variant` that holds a Double on success and an Excel error value on failure. Test it with `IsError` before you use it as a number.

```vba
Public Sub Example_PoissonQuantile()

    Dim Result As Variant
    Dim Status As String

    Result = K_STATS_Poisson_InverseCumulative(0.99, 25#, Status)

    If IsError(Result) Then
        Debug.Print "Calculation failed: " & Status
        Exit Sub
    End If

    Debug.Print "Poisson quantile: "; CDbl(Result)

End Sub
```

This example prints `37`.

- Declare the result as `Variant`, not as `Double`.
- The functions do not raise run-time errors for invalid input and do not show message boxes.
- A discrete quantile is returned as a Double that holds a whole number.

## 5.8 The fast inverse normal helper

`K_STATS_NormalStandard_InverseCumulativeFast` is meant for VBA loops that turn many uniform random numbers into standard normal values. It differs from all other functions:

- it returns a `Double` and has no `Status` argument;
- it never returns an error value: a probability at or below 0, or at or above 1, is moved slightly inside the interval;
- it is less accurate than `K_STATS_NormalStandard_InverseCumulative` (section 10.3).

Use it only when your code has already validated the probability.
