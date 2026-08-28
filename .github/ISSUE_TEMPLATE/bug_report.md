---
name: 🐞 Bug report
about: Report a wrong result, a spurious error, a wrong error code, a performance defect, or a crash
title: "[Bug]: "
labels: bug
assignees: danielep71
---

## 🐞 Description

Describe the defect and its practical impact.

> [!IMPORTANT]
> Do not use this public template for a suspected security vulnerability.
> Use the repository's private vulnerability-reporting channel described in
> `SECURITY.md`.

## 🚦 Defect class

Tick exactly one. This is the single most important field: it determines
severity independently of how large the numerical discrepancy is.

- [ ] **Silent wrong value** — a plausible number is returned and it is wrong
- [ ] **Spurious error** — a valid input is refused, though the correct answer
      is representable
- [ ] **Wrong error code** — correctly refused, but the wrong `CVErr` is returned
- [ ] **Unhandled runtime error** — a VBA error escapes to the caller
- [ ] **Accuracy shortfall** — a value is within its documented envelope but
      misses its accuracy contract
- [ ] **Performance or non-convergence**

A silent wrong value outranks everything else here, because the caller has no
way to detect it. A spurious error is visible and self-announcing.

## 🔖 Version and source state

```text
Release tag:      <e.g. v1.0.0, or N/A>
Commit SHA:       <full 40-character SHA if using main or another snapshot>
Source obtained:  <official repository / tagged source archive / other>
```

Do not write only "latest."

## 🔢 Exact call and result

```text
Function: K_STATS_...(...)
Returned: <value / #NUM! / #VALUE! / runtime error number and description>
Expected: <value or error code>
Status:   <optional ByRef Status text, when available>
```

### Argument bit-exactness

**Required whenever any argument is subnormal, denormal, near an overflow
boundary, or beyond about 15 significant digits.** A VBA source literal cannot
be relied on to denote such a value, and `Val()` cannot be relied on to parse
one, so a decimal written here may not be the number that was actually tested.

Give each such argument as a two-part sum `hi;lo` (summing to the exact Double),
or state how it was constructed by arithmetic:

```text
Argument:     <name>
hi;lo:        <e.g. 5.56268464626800E-309;4.94065645841247E-324>
Constructed:  <e.g. halve 1# 1024 times, or 1E-300 / 1E+20, or N/A - literal>
```

### Smallest runnable example

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

Build extreme arguments by arithmetic inside this Sub rather than pasting a
decimal literal, so the example reproduces the exact Double that was tested.

## 🎯 Boundary localisation

*Optional, but the most valuable evidence you can supply for a numerical
defect.* If the defect appears below or above some threshold, narrow it to the
two adjacent Doubles that straddle it — the last input that works and the first
that fails.

```text
Last good input:   <hi;lo or construction>   Returned: <value>
First bad input:   <hi;lo or construction>   Returned: <value / error>
Adjacent Doubles?  yes / no / unknown
Predicted boundary and why: <e.g. 1/DoubleMax, because the expression forms 1/N>
```

A boundary pinned to adjacent Doubles turns "it breaks somewhere down there"
into a specific, testable claim about a specific expression.

## 🔬 Independent reference

Explain how the expected result was obtained.

Examples: SciPy with the exact function and version; mpmath with the precision
used; R, Julia or MATLAB; an authoritative published table or paper; a
mathematically exact identity.

```text
Reference system:
Reference function:
Version:
Precision:
Expected result:
```

Include enough digits to evaluate the reported discrepancy. If the expected
value is subnormal or beyond 15 significant digits, give it as `hi;lo` too.

## 🔁 Steps to reproduce

1.
2.
3.

## 🧪 Environment

```text
Excel version:
Office build:
Office bitness:   32-bit / 64-bit
Operating system:
Locale:           <decimal separator matters for parsing and Format$>
Use context:      worksheet formula / VBA call
Workbook type:    .xlsm / .xlsb / other
```

## ✅ Regression-harness result

Run the most relevant suite when possible and paste the Immediate Window output.

```text
Test_STATS_PROBDIST_RunAll            →
Test_STATS_PROBDIST_RunCore           →
Test_STATS_PROBDIST_RunNormalFamily   →
Test_STATS_PROBDIST_RunTFamily        →
Test_STATS_PROBDIST_RunContinuous     →
Test_STATS_PROBDIST_RunDiscrete       →
```

**Does the harness detect this defect?**

- [ ] Yes — a suite fails, named above
- [ ] No — every suite passes and the defect is still present

If the harness passes, say so plainly. A defect the harness cannot see is a gap
in the evidence set as well as a defect in the code, and both need fixing.

## 📋 Contract coverage

*Optional but strongly encouraged.* This library is contract-driven: a region
that no active contract claims is scored by nothing and reported by nothing.

```text
Governing contract:  <contract_id from benchmark/accuracy_contracts.csv, or none>
Threshold:           <as frozen, or N/A>
Failing region covered by an active contract?  yes / no / unsure
Grid rows in the failing region:               <count, or none>
```

If the answer is *no*, note it here — the missing coverage is a separate
finding from the defect itself, and it is usually the reason the defect
survived.

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
- [ ] Subnormal or denormal argument
- [ ] Overflow or underflow region
- [ ] Parameter-validation boundary
- [ ] Cancellation-prone region
- [ ] Inverse round-trip
- [ ] Moment calculation
- [ ] Error-code classification
- [ ] Performance or apparent non-convergence

## 🌐 Affected surface

*Optional.* Which public functions can reach the defect, and which provably
cannot. Naming what is **not** affected is as useful as naming what is: it
bounds the remediation and gives the reviewer a movement prediction to check
against.

```text
Reachable from:
Provably unaffected:
Reason the unaffected paths cannot reach it:
```

## 🛠️ Proposed fix

*Optional.* If you have one, give the replacement expression and say what
evidence supports it — particularly whether it changes results anywhere the
current code already works.

```text
Proposed change:
Measured worst error:
Bit-identical to current where current works?  yes / no / partly
Expected contract impact:
```

Do not propose loosening a frozen threshold to accommodate a defect. If a
contract is breached, that is a regression to investigate first.

## 📎 Additional context

Add screenshots, formulas, benchmark artifacts, or links that help reproduce
the issue.

Do not attach workbooks containing confidential, personal, client, or
production data.
