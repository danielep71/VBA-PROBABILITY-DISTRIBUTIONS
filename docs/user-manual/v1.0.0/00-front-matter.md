# About this manual

This manual is for people who use VBA Probability Distributions v1.0.0 from an Excel worksheet or from their own VBA code. It explains how to install the library, how to call its functions, which conventions it follows and where its numerical limits are. It does not explain the algorithms, the test tooling or the release procedure.

The exact declaration of every function is in the Public API reference: [`docs/PUBLIC_API.md`](../../PUBLIC_API.md) and its inventory [`docs/PUBLIC_API.txt`](../../PUBLIC_API.txt). Chapter 13 gives a compact overview.

The manual describes v1.0.0 only. For another version, read its changelog first.

## Conventions used

- Worksheet formulas appear in blocks like this one:

```text
=K_STATS_NormalStandard_Cumulative(1)
```

- Formulas use a comma between arguments and a point as the decimal separator. If your Excel uses a semicolon and a decimal comma, type them instead.
- Function and argument names are written exactly as declared, for example `K_STATS_Gamma_Survival` and `ScaleParam`.
- `1E-13` means $10^{-13}$.
- Expected results are shown to at most 10 significant digits.
- $P(X \le x)$ is the probability that the variable $X$ is at most $x$.
- "Relative error" is the error divided by the size of the true value.
- A "Double" is the 64-bit floating-point number used by Excel and VBA.

Other terms are explained in chapter 12.
