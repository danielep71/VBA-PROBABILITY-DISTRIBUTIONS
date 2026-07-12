# 🤝 Contributing to VBA-PROBABILITY-DISTRIBUTIONS

<p align="left">
  <img alt="Contributions" src="https://img.shields.io/badge/Contributions-Welcome-217346">
  <img alt="Language" src="https://img.shields.io/badge/Language-Excel_VBA-blue">
  <img alt="Style" src="https://img.shields.io/badge/Style-House_Conventions-6f42c1">
  <img alt="Tests" src="https://img.shields.io/badge/Tests-Test__STATS__PROBDIST__RunAll-orange">
  <img alt="Verification" src="https://img.shields.io/badge/Verification-Independent_Reference-00A3E0">
  <img alt="License" src="https://img.shields.io/badge/License-MIT-green">
</p>

Thank you for your interest in improving the library.

This is a numerical project. It therefore prioritizes:

- small and reviewable changes;
- explicit statistical parameterization;
- documented numerical methods;
- predictable failure contracts;
- independently calculated reference values;
- regression tests that protect the corrected behavior.

A contribution is not complete merely because it compiles. It must also explain and
verify the numerical behavior it changes.

---

## 💬 Before you start

<p align="left">
  <img alt="Step" src="https://img.shields.io/badge/Step-Open_an_Issue_First-217346">
</p>

Open an issue before beginning non-trivial work so the intended API, numerical
method, and scope can be agreed in advance.

Good issues include:

- a reproducible wrong result or wrong worksheet-error code;
- a numerical-accuracy problem supported by an independent reference;
- a focused new distribution, moment, tail, or interval function;
- a convergence, overflow, underflow, or performance concern;
- a documentation gap or inaccurate parameter description.

Tiny corrections such as typographical fixes, comment corrections, or obvious
one-line defects may go directly to a pull request.

Suspected security vulnerabilities must be reported privately under
[SECURITY.md](SECURITY.md), not through a public issue.

---

## 🧰 Project layout

<p align="left">
  <img alt="Source" src="https://img.shields.io/badge/Source-Exported_VBA-217346">
  <img alt="Tooling" src="https://img.shields.io/badge/Workflow-Import_Edit_Export-blue">
  <img alt="External runtime" src="https://img.shields.io/badge/External_Runtime-None-lightgrey">
</p>

The repository stores exported VBA source rather than a binary production
workbook.

```text
src/M_STATS_PROBDIST_CORE.bas
    Shared constants, predicates, guarded arithmetic, Log1p, Expm1,
    the raw inverse-normal seed, and diagnostic status support.

src/M_STATS_PROBDIST_SPECIALFUNCS.bas
    Log-gamma, log-beta, stable log-combination, regularized incomplete
    beta and gamma functions, continued fractions, series expansions,
    and safeguarded inverse special functions.

src/M_STATS_PROBDIST_NORMALFAMILY.bas
    Standard Normal, Normal, and Lognormal worksheet functions.

src/M_STATS_PROBDIST_TFAMILY.bas
    Student t, Chi-square, and F worksheet functions.

src/M_STATS_PROBDIST_CONTINUOUS.bas
    Gamma, Beta, Exponential, Weibull, and continuous Uniform functions.

tests/M_STATS_PROBDIST_TEST.bas
    Consolidated assertions, family suites, reference values, and
    regression tests.
```

`M_STATS_PROBDIST_CORE` and `M_STATS_PROBDIST_SPECIALFUNCS` are the shared
numerical foundation. Distribution-family modules consume those routines rather
than maintaining private duplicate implementations.

---

## 🔁 Edit and export workflow

1. Import the required `.bas` files into a macro-enabled Excel workbook.

   Recommended dependency order:

   ```text
   CORE → SPECIALFUNCS → distribution-family modules → TEST
   ```

2. Make the change in the VBA Editor.

3. Compile with:

   ```text
   Debug → Compile VBAProject
   ```

4. Run the relevant family suite and then the complete regression harness.

5. For any numerical change, compare the result with an independent reference.

6. Re-export each changed module over the matching file under `src/` or `tests/`.

7. Review the textual diff before committing.

Do not commit the workbook used to edit or test the source.

---

## 🧱 Coding standards

<p align="left">
  <img alt="Option Explicit" src="https://img.shields.io/badge/Option_Explicit-Required-217346">
  <img alt="Procedure headers" src="https://img.shields.io/badge/Procedure_Headers-Required-blue">
  <img alt="Naming" src="https://img.shields.io/badge/Naming-Namespaced-6f42c1">
</p>

### Module visibility

- Every module must use `Option Explicit`.
- `M_STATS_PROBDIST_CORE` and `M_STATS_PROBDIST_SPECIALFUNCS` use
  `Option Private Module`.
- Distribution-family modules must remain worksheet-visible and therefore must
  not use `Option Private Module`.
- The consolidated test module exposes only its intended public runner procedures.

### Procedure headers

Every module and procedure uses a structured banner.

Include the fields relevant to the routine:

```text
PURPOSE
WHY
INPUTS
RETURNS
BEHAVIOR
ERROR POLICY
DEPENDENCIES
NOTES
CALLED FROM
UPDATED
```

Where relevant, preserve specific sections such as:

```text
ALGORITHM PROVENANCE
DESIGN PRINCIPLES
ACCURACY
```

Provenance and accuracy statements must be specific and supportable. Do not use
an unqualified phrase such as “high accuracy.”

### Body structure

Use only the sections a routine needs, generally in this order:

```text
DECLARE
INITIALIZE
VALIDATE INPUTS
COMPUTE
RETURN SUCCESS
FAIL - NUMERIC
ERROR HANDLER
```

Place explanatory comments above the code they describe. Use inline comments
primarily for declarations.

Use the established literal and string forms where applicable:

```vba
0#
1#
vbNullString
```

---

## 🧩 Naming and callable contracts

The prefix identifies the intended scope, but **does not by itself imply one
universal VBA return type**.

| Naming pattern | Scope | Typical callable contract |
|---|---|---|
| `K_STATS_*` | Worksheet-facing statistical API | Normally `Variant`, returning either a numerical value or `CVErr` |
| `K_STATS_NormalStandard_InverseCumulativeFast` | Specialized fast public helper | `Double`; intentionally lighter validation and error contract |
| `PROB_Is*` | Project-scoped predicates | `Boolean` |
| `PROB_Try*` | Guarded arithmetic or iterative numerical operations | `Boolean`, with results returned through `ByRef` arguments |
| Other `PROB_*` functions | Project-scoped numerical helpers and kernels | Usually `Double`, but may be another explicitly documented type |
| `PROB_SetStatus` | Project-scoped diagnostic writer | `Sub` |
| `Test_STATS_PROBDIST_*` | Public test-harness entry points | Argument-less `Sub` |

Do not document all `PROB_*` routines as returning `Double`. The current
project-scoped API deliberately includes Boolean predicates, Boolean Try
contracts, Double-valued kernels, and a status-writing Sub.

---

## 🧯 Public error-handling contract

For ordinary worksheet-facing UDFs:

- use `On Error GoTo Err_Handler`;
- invalid domains, predictable overflow, density poles, and non-convergence route
  to `Fail_Num:`;
- `Fail_Num:` returns `CVErr(xlErrNum)` (`#NUM!`);
- unexpected runtime errors route to `Err_Handler:`;
- `Err_Handler:` returns `CVErr(xlErrValue)` (`#VALUE!`);
- mathematically valid exponential underflow returns zero;
- no numerical UDF displays a `MsgBox`;
- the optional `Status` argument is the detailed diagnostic channel.

The fast inverse-normal helper is a documented exception with a deliberately
lighter `Double`-returning contract.

---

## 🧠 Numerical-kernel discipline

### Domain responsibility

Do not apply one blanket rule to every `PROB_*` routine.

- Predicates such as `PROB_IsFinite` and
  `PROB_IsValidProbabilityOpen` exist specifically to test domains.
- Public distribution wrappers own end-user parameter validation and worksheet
  error mapping.
- Low-level numerical kernels generally assume that their documented
  preconditions have already been checked.
- `PROB_SetStatus` is the deliberate project-scoped status writer; ordinary
  numerical kernels should not write user-facing diagnostics directly.

### Guarded arithmetic

Use:

```text
PROB_TryExp
PROB_TryAdd
PROB_TryMultiply
PROB_TryDivide
```

when an intermediate result may overflow or when failure classification matters.

Use:

```text
PROB_Log1p
PROB_Expm1
```

for cancellation-sensitive expressions such as:

```text
Log(1 + x)
Exp(x) - 1
```

Do not reconstruct a small upper tail as `1 - CDF` when a direct survival kernel
is available.

### VBA Boolean operators

VBA does not short-circuit `And` and `Or`.

This is unsafe:

```vba
If Denominator <> 0# And Numerator / Denominator > Limit Then
```

Use separate guards:

```vba
If Denominator = 0# Then
    GoTo Fail_Num
End If

If Numerator / Denominator > Limit Then
    GoTo Fail_Num
End If
```

---

## 🔬 Numerical verification

<p align="left">
  <img alt="Reference" src="https://img.shields.io/badge/Reference-SciPy_%2F_mpmath-217346">
  <img alt="Precision" src="https://img.shields.io/badge/Precision-40--60_Digits-blue">
  <img alt="Test values" src="https://img.shields.io/badge/Expected_Values-Independently_Calculated-red">
</p>

Every new or changed expected value must be calculated independently.

Do not:

- estimate a value from a chart;
- round a remembered value;
- copy a value produced by the same VBA code path being tested;
- insert a placeholder literal and weaken the tolerance around it.

For a numerical change, verify as applicable:

- density;
- CDF;
- survival;
- inverse CDF;
- moments;
- support boundaries;
- central-region behavior;
- lower and upper tails;
- round-trip identities;
- overflow and valid underflow;
- exact worksheet-error classification.

Suitable independent references include SciPy, mpmath, R, Julia, MATLAB, or
authoritative published tables and formulas.

State in the pull request:

- the reference system and version;
- the precision used;
- the points tested;
- the maximum observed absolute or relative error;
- the tolerance justified by that evidence.

Cross-family identities are valuable, but they should exercise independent
calculation paths rather than compare two wrappers around the same kernel.

---

## 🧪 Testing

Import `tests/M_STATS_PROBDIST_TEST.bas` and run:

```vba
Test_STATS_PROBDIST_RunAll
```

Family runners:

```vba
Test_STATS_PROBDIST_RunCore
Test_STATS_PROBDIST_RunNormalFamily
Test_STATS_PROBDIST_RunTFamily
Test_STATS_PROBDIST_RunContinuous
```

Passing assertions are silent. Failures print detailed output to the Immediate
Window, followed by a consolidated verdict.

Requirements:

- all existing suites must pass;
- changed behavior must be covered by the relevant family suite;
- a corrected numerical defect must receive a regression test;
- use exact error-code assertions where `#NUM!` versus `#VALUE!` is contractual;
- do not weaken a valid test merely to make a changed implementation pass.

---

## 📚 Documentation expectations

Documentation is part of the pull request.

Update the relevant pages when changing:

- **public UDFs**  
  [API Reference](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/wiki/API-Reference)

- **architecture or module boundaries**  
  [Architecture](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/wiki/Architecture)  
  [Module Reference](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/wiki/Module-Reference)

- **special-function kernels or algorithms**  
  [Special Functions and Numerical Kernels](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/wiki/Special-Functions-and-Numerical-Kernels)  
  [Numerical Accuracy and Design](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/wiki/Numerical-Accuracy-and-Design)

- **error behavior**  
  [Error Handling and Diagnostics](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/wiki/Error-Handling-and-Diagnostics)

- **test behavior or regression cases**  
  [Testing and Regression Harness](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/wiki/Testing-and-Regression-Harness)

Also update the README when a change affects:

- distribution coverage;
- function counts;
- installation;
- parameterization;
- public examples;
- the repository structure.

---

## 📦 What not to commit

Do not commit:

- the workbook used to edit or test the source;
- Excel lock or owner files such as `~$*`;
- local scratch files;
- generated output;
- machine-specific paths or settings;
- client, personal, or confidential data;
- manually edited `.bas` files that were not re-exported and checked against the
  VBE version.

See `.gitignore` for the current exclusion policy.

---

## 🚀 Submitting a pull request

1. Fork the repository and create a focused branch.
2. Keep one logical change per pull request.
3. Complete every applicable section of the pull-request template.
4. State the numerical method and independent reference.
5. Confirm the project compiles.
6. Run the relevant family suite.
7. Run `Test_STATS_PROBDIST_RunAll`.
8. Re-export changed modules.
9. Update the documentation.
10. Review the final text diff.

The maintainer may adopt, adapt, defer, or decline a contribution to preserve the
coherence of the library.

---

## 📄 License

By contributing, you agree that your contribution is licensed under the
project’s [MIT License](LICENSE).

---

## 👤 Maintainer

Maintained by **Daniele Penza**.
