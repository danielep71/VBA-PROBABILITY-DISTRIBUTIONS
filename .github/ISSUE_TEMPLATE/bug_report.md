---
name: 🐞 Bug report
about: Report a wrong result, wrong error code, or crash
title: "[Bug]: "
labels: bug
---

## 🐞 Description

A clear description of what is wrong.

## 🔢 The call

The exact function and arguments, and what came back.

```text
Function:  K_STATS_...(...)
Returned:  <value or error code, e.g. 0.1337 / #NUM! / #VALUE!>
Expected:  <value or error code>
```

## 🔬 Reference for the expected value

How did you obtain the expected value? (e.g. SciPy `scipy.stats...`, mpmath at N
digits, a textbook table). Independent references make a numerical report
actionable.

```text

```

## 🔁 Steps to reproduce

1.
2.
3.

## 🧪 Environment

- Library version: <!-- e.g. v1.0.0 -->
- Use context: <!-- worksheet formula / VBA call -->
- Excel version: <!-- e.g. Excel 2021 / Microsoft 365 -->
- Office bitness: <!-- 32-bit / 64-bit -->
- OS: <!-- e.g. Windows 11 / macOS -->

## 🖥️ Diagnostics

Paste any runtime error and relevant **Immediate window** output. If you can run
the harness, note whether the matching suite reproduces it:

```text
Test_STATS_PROBDIST_Run... →
```

## 📎 Additional context

Screenshots or anything else that helps. Please do **not** attach workbooks
containing sensitive data.
