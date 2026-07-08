# VBA Probability Distributions

Excel VBA library of probability distribution functions for finance, risk modelling, teaching, and model validation.

## Purpose

This repository provides documented, auditable and worksheet-friendly probability distribution routines written in Excel VBA.

The library is designed for:

- finance and risk modelling;
- derivatives and Monte Carlo examples;
- credit risk modelling;
- teaching and academic material;
- model validation and audit walkthroughs.

## Design principles

- Clear function names.
- Explicit input validation.
- Worksheet-facing functions return `Variant` and `CVErr` on failure.
- Fast numerical kernels are separated from public worksheet functions.
- No `MsgBox` inside worksheet-facing functions.
- Diagnostic messages are written to optional `Status` arguments and to the Excel status bar.
- Source code is exported as plain `.bas` modules for version control.

## Current modules

| Module | Description |
|---|---|
| `M_ProbabilityDistributions_NormalFamily.bas` | Standard normal, general normal and lognormal distributions |

## Current functions

### Standard normal

- `K_StandardNormal_Density`
- `K_StandardNormal_Cumulative`
- `K_StandardNormal_InverseCumulative`

### General normal

- `K_Normal_Density`
- `K_Normal_Cumulative`
- `K_Normal_InverseCumulative`
- `K_Normal_ZScore`
- `K_Normal_InverseCumulativeFast`

### Lognormal

- `K_Lognormal_Density`
- `K_Lognormal_Cumulative`
- `K_Lognormal_InverseCumulative`
- `K_Lognormal_Mean`
- `K_Lognormal_Variance`
- `K_Lognormal_ParametersFromMeanStdDev`

## Naming convention

Public worksheet-facing routines use the prefix:

```vba
K_
