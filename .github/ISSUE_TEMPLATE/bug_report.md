---
name: 🐞 Bug report
about: Report a wrong result, wrong error code, performance defect, or crash
title: "[Bug]: "
labels: bug
---

## 🐞 Description

Provide a clear description of the defect and its practical impact.

> [!IMPORTANT]
> Do not use this public template for a suspected security vulnerability.
> Use the repository’s private vulnerability-reporting channel described in
> `SECURITY.md`.

## 🔖 Version and source state

Identify the exact source you tested.

```text
Release tag:     <e.g. v1.0.0, or N/A>
Commit SHA:      <full 40-character SHA if using main or another snapshot>
Source obtained: <official repository / tagged source archive / other>
```

Do not write only “latest.”

## 🔢 Exact call and result

Provide the exact function, arguments, and returned value or worksheet error.

```text
Function: K_STATS_...(...)
Returned: <value / #NUM! / #VALUE! / runtime error>
Expected: <value or error code>
Status:   <optional ByRef Status text, when available>
```

For a VBA call, include the smallest runnable example:

```vba
Option Explicit

Public Sub ReproduceIssue()

    Dim Status As String
    Dim Result As Variant

    Result = K_STATS_...(Status:=Status)

    Debug.Print Result
    Debug.Print Status

End Sub
```

## 🔬 Independent reference

Explain how the expected result was obtained.

Examples:

- SciPy, including the exact function and version;
- mpmath, including the precision used;
- R, Julia, MATLAB, or another numerical library;
- an authoritative published table or paper;
- a mathematically exact identity.

```text
Reference system:
Reference function:
Version:
Precision:
Expected result:
```

Include enough digits to evaluate the reported discrepancy.

## 🔁 Steps to reproduce

1.
2.
3.

## 🧪 Environment

```text
Excel version:
Office bitness: 32-bit / 64-bit
Operating system:
Use context: worksheet formula / VBA call
Workbook type: .xlsm / .xlsb / other
```

## ✅ Regression-harness result

Run the most relevant suite when possible.

```text
Test_STATS_PROBDIST_RunAll            →
Test_STATS_PROBDIST_RunCore           →
Test_STATS_PROBDIST_RunNormalFamily   →
Test_STATS_PROBDIST_RunTFamily        →
Test_STATS_PROBDIST_RunContinuous     →
```

Paste the relevant Immediate Window output.

## 📐 Numerical region

Check any that apply:

- [ ] Central distribution body
- [ ] Lower tail
- [ ] Upper tail
- [ ] Probability close to `0`
- [ ] Probability close to `1`
- [ ] Support boundary
- [ ] Very small parameter
- [ ] Very large parameter
- [ ] Overflow or underflow region
- [ ] Inverse round-trip
- [ ] Moment calculation
- [ ] Error-code classification
- [ ] Performance or apparent non-convergence

## 📎 Additional context

Add screenshots, formulas, or links that help reproduce the issue.

Do not attach workbooks containing confidential, personal, client, or production
data.
