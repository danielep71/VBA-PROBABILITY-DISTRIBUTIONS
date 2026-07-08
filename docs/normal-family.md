# 📘 Normal Family Distributions

```text
==============================================================================
📘 NORMAL FAMILY DISTRIBUTIONS
------------------------------------------------------------------------------
Module:  M_STATS_PROBDIST_NORMALFAMILY
Scope:   Standard normal, general normal, and lognormal distributions
Status:  Baseline module for the first release of VBA-Probability-Distributions
==============================================================================
```

## 🎯 Purpose

`M_STATS_PROBDIST_NORMALFAMILY` provides robust Excel VBA routines for the normal family of probability distributions.

The module is designed for:

- 📈 finance and risk modelling;
- 🧮 worksheet formulas;
- ⚙️ VBA numerical routines;
- 🎲 Monte Carlo simulation;
- 🧪 model validation and testing;
- 🎓 teaching and audit walkthroughs.

The library intentionally provides transparent VBA alternatives and extensions to selected native Excel worksheet functions. Native Excel already covers many of these distributions, but VBA calls to `Application.WorksheetFunction` can be awkward in production code because invalid inputs raise runtime errors and repeated calls introduce worksheet-function marshalling overhead.

---

## 🧱 Module Design

```text
==============================================================================
DESIGN LAYERS
------------------------------------------------------------------------------
Public API       K_STATS_<Distribution>_<Operation>
Private kernels  PROB_<TechnicalKernelName>
Error contract   Variant / CVErr for public worksheet-facing functions
Fast routines    Double -> Double for validated numerical callers
==============================================================================
```

### Public API convention

Public worksheet-facing routines use:

```vba
K_STATS_<Distribution>_<Operation>
```

Examples:

```vba
K_STATS_Normal_Cumulative
K_STATS_NormalStandard_InverseCumulative
K_STATS_Lognormal_ParametersFromMeanStdDev
```

### Private kernel convention

Private numerical routines use:

```vba
PROB_<TechnicalKernelName>
```

Examples:

```vba
PROB_NormalPDF
PROB_NormalCDF
PROB_NormalInvCDF
PROB_TryExp
```

---

## 📦 Public Functions

```text
==============================================================================
PUBLIC FUNCTION MAP
==============================================================================
```

### 🔵 Standard Normal Distribution

| Function | Purpose | Excel equivalent |
|---|---|---|
| `K_STATS_NormalStandard_Density` | Standard normal density φ(z) | `NORM.S.DIST(z, FALSE)` |
| `K_STATS_NormalStandard_Cumulative` | Standard normal cumulative Φ(z) | `NORM.S.DIST(z, TRUE)` |
| `K_STATS_NormalStandard_InverseCumulative` | Standard normal quantile Φ⁻¹(p) | `NORM.S.INV(p)` |
| `K_STATS_NormalStandard_IntervalProbability` | Probability between two z-values | CDF difference |
| `K_STATS_NormalStandard_InverseCumulativeFast` | Fast inverse standard normal for Monte Carlo | None |

### 🔷 General Normal Distribution

| Function | Purpose | Excel equivalent |
|---|---|---|
| `K_STATS_Normal_Density` | Normal density with mean and standard deviation | `NORM.DIST(x, mean, stddev, FALSE)` |
| `K_STATS_Normal_Cumulative` | Normal cumulative with mean and standard deviation | `NORM.DIST(x, mean, stddev, TRUE)` |
| `K_STATS_Normal_InverseCumulative` | Normal quantile with mean and standard deviation | `NORM.INV(p, mean, stddev)` |
| `K_STATS_Normal_ZScore` | Standardized score `(x - mean) / stddev` | `STANDARDIZE(x, mean, stddev)` |
| `K_STATS_Normal_IntervalProbability` | Probability between two normal values | CDF difference |

### 🟣 Lognormal Distribution

| Function | Purpose | Excel equivalent |
|---|---|---|
| `K_STATS_Lognormal_Density` | Lognormal density | `LOGNORM.DIST(x, meanlog, stddevlog, FALSE)` |
| `K_STATS_Lognormal_Cumulative` | Lognormal cumulative probability | `LOGNORM.DIST(x, meanlog, stddevlog, TRUE)` |
| `K_STATS_Lognormal_InverseCumulative` | Lognormal quantile | `LOGNORM.INV(p, meanlog, stddevlog)` |
| `K_STATS_Lognormal_Mean` | Arithmetic mean of lognormal variable | None |
| `K_STATS_Lognormal_Variance` | Arithmetic variance of lognormal variable | None |
| `K_STATS_Lognormal_StdDev` | Arithmetic standard deviation of lognormal variable | None |
| `K_STATS_Lognormal_ParametersFromMeanStdDev` | Convert arithmetic mean/stddev into log-space parameters | None |

---

## 🧮 Distribution Definitions

```text
==============================================================================
MATHEMATICAL DEFINITIONS
==============================================================================
```

### 🔵 Standard Normal

For a standard normal variable `Z ~ N(0, 1)`:

```text
Density:
φ(z) = exp(-0.5 z²) / sqrt(2π)

Cumulative:
Φ(z) = P(Z ≤ z)

Inverse cumulative:
Φ⁻¹(p) = z such that Φ(z) = p
```

### 🔷 General Normal

For `X ~ N(μ, σ²)` with `σ > 0`:

```text
Z = (X - μ) / σ

Density:
f(x) = φ((x - μ) / σ) / σ

Cumulative:
F(x) = Φ((x - μ) / σ)

Inverse cumulative:
F⁻¹(p) = μ + σ Φ⁻¹(p)
```

### 🟣 Lognormal

For `X ~ Lognormal(μ, σ²)`, where `Log(X) ~ N(μ, σ²)` and `σ > 0`:

```text
Density:
f(x) = φ((ln(x) - μ) / σ) / (x σ),  x > 0

Cumulative:
F(x) = Φ((ln(x) - μ) / σ),  x > 0

Inverse cumulative:
F⁻¹(p) = exp(μ + σ Φ⁻¹(p))
```

Arithmetic moments:

```text
Mean:
E[X] = exp(μ + 0.5 σ²)

Variance:
Var[X] = (exp(σ²) - 1) exp(2μ + σ²)

Standard deviation:
StdDev[X] = sqrt(Var[X])
```

Parameter conversion from arithmetic mean `m` and arithmetic standard deviation `s`:

```text
σ² = ln(1 + s² / m²)
σ  = sqrt(ln(1 + s² / m²))
μ  = ln(m) - 0.5 σ²
```

---

## 🧾 Error Policy

```text
==============================================================================
ERROR POLICY
------------------------------------------------------------------------------
Public worksheet-facing functions return Variant so they can return CVErr.
No public worksheet-facing function raises MsgBox.
==============================================================================
```

| Situation | Return |
|---|---|
| Success | `Double` value inside `Variant` |
| Invalid numeric domain | `CVErr(xlErrNum)` |
| Unexpected runtime error | `CVErr(xlErrValue)` |

Examples of invalid numeric domains:

- probability not strictly inside `(0, 1)` for inverse CDF functions;
- standard deviation less than or equal to zero;
- non-positive `x` for lognormal density;
- exponential overflow in public lognormal moment or inverse routines.

The optional `Status` argument is intended mainly for VBA callers. Worksheet users should normally omit it.

---

## ⚙️ Fast Monte Carlo Function

```text
==============================================================================
FAST NUMERICAL ENTRY POINT
==============================================================================
```

`K_STATS_NormalStandard_InverseCumulativeFast` is intentionally different from the worksheet-facing inverse CDF.

It:

- returns `Double`, not `Variant`;
- does not return `CVErr`;
- does not write diagnostic status;
- clips endpoint probabilities defensively;
- uses the raw Acklam approximation for speed.

Use it in simulation loops such as:

```vba
Z = K_STATS_NormalStandard_InverseCumulativeFast(U)
```

where `U` is expected to be a validated uniform random number strictly inside `(0, 1)`.

---

## 🧪 Testing

```text
==============================================================================
TEST MODULE
------------------------------------------------------------------------------
Module: M_STATS_PROBDIST_TEST
Runner: Test_STATS_PROBDIST_NormalFamily_RunAll
==============================================================================
```

The normal-family test set should cover:

- known standard normal density values;
- known standard normal cumulative values;
- inverse-normal benchmark values;
- inverse/CDF round trips;
- symmetry checks;
- general normal scaling and z-score behavior;
- interval probabilities;
- fast inverse-normal smoke tests;
- lognormal density, cumulative and inverse values;
- lognormal mean, variance and standard deviation;
- lognormal parameter conversion;
- invalid-domain error behavior;
- overflow behavior.

Recommended command from the VBA Immediate Window:

```vba
Test_STATS_PROBDIST_NormalFamily_RunAll
```

Future test families can be added to the same test module:

```vba
Test_STATS_PROBDIST_Discrete_RunAll
Test_STATS_PROBDIST_Continuous_RunAll
Test_STATS_PROBDIST_Bivariate_RunAll
Test_STATS_PROBDIST_RunAll
```

---

## 🧭 Edge-Case Behaviour

```text
==============================================================================
EDGE-CASE POLICY
==============================================================================
```

### Probability inputs

Inverse CDF functions require:

```text
0 < Probability < 1
```

Invalid probabilities return `CVErr(xlErrNum)`.

### Standard deviations

Normal and lognormal routines require strictly positive standard deviations:

```text
StdDev > 0
StdDevLog > 0
```

Invalid standard deviations return `CVErr(xlErrNum)`.

### Lognormal support

The lognormal density is defined only for:

```text
X > 0
```

For the lognormal cumulative distribution, the mathematically correct left-tail value is returned:

```text
If X <= 0, F(X) = 0
```

This differs from native Excel `LOGNORM.DIST`, which may return `#NUM!` for non-positive `X`.

### Overflow

Public lognormal inverse and moment functions should return `CVErr(xlErrNum)` when the final result would overflow `Double`.

---

## 🏗️ Repository Placement

Recommended file layout:

```text
VBA-Probability-Distributions/
│
├─ src/
│  └─ M_STATS_PROBDIST_NORMALFAMILY.bas
│
├─ tests/
│  └─ M_STATS_PROBDIST_TEST.bas
│
├─ docs/
│  └─ normal-family.md
│
├─ examples/
│  └─ ...
│
├─ README.md
├─ CHANGELOG.md
├─ LICENSE
└─ .gitignore
```

---

## 📌 Example Worksheet Usage

```excel
=K_STATS_NormalStandard_Cumulative(1.96)
```

```excel
=K_STATS_Normal_InverseCumulative(0.975, 0, 1)
```

```excel
=K_STATS_Normal_IntervalProbability(-1.96, 1.96, 0, 1)
```

```excel
=K_STATS_Lognormal_Cumulative(100, 4.5, 0.2)
```

```excel
=K_STATS_Lognormal_ParametersFromMeanStdDev(100, 20)
```

---

## 📌 Example VBA Usage

```vba
Dim Status As String
Dim Result As Variant

Result = K_STATS_Normal_Cumulative( _
    1.5, _
    0#, _
    1#, _
    Status)

If IsError(Result) Then
    Debug.Print "Error: " & Status
Else
    Debug.Print "Probability: " & CDbl(Result)
End If
```

Monte Carlo use:

```vba
Dim U As Double
Dim Z As Double

U = Rnd()
If U <= 0# Then U = 0.000000000000001
If U >= 1# Then U = 1# - 0.000000000000001

Z = K_STATS_NormalStandard_InverseCumulativeFast(U)
```

---

## 🚧 Future Enhancements

Potential additions to the normal family:

- `K_STATS_NormalStandard_Survival`
- `K_STATS_Normal_Survival`
- `K_STATS_Lognormal_Survival`
- `K_STATS_Lognormal_Median`
- `K_STATS_Lognormal_Mode`
- `K_STATS_Lognormal_CoefficientOfVariation`
- `K_STATS_Lognormal_ParametersFromMeanCV`

Potential numerical improvement:

- tail-stable normal survival probability;
- tail-stable normal interval probability;
- log-space implementation for extreme lognormal variance and standard deviation.

---

## ✅ Current Baseline

```text
==============================================================================
BASELINE
------------------------------------------------------------------------------
Source module:  M_STATS_PROBDIST_NORMALFAMILY.bas
Test module:    M_STATS_PROBDIST_TEST.bas
Public runner:  Test_STATS_PROBDIST_NormalFamily_RunAll
==============================================================================
```

This module is the first distribution family in the `VBA-Probability-Distributions` repository.
