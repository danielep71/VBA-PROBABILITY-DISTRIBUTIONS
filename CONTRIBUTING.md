# 🤝 Contributing to VBA-PROBABILITY-DISTRIBUTIONS

<p align="left">
  <img alt="Contributions" src="https://img.shields.io/badge/Contributions-Welcome-217346">
  <img alt="Language" src="https://img.shields.io/badge/Language-Excel_VBA-blue">
  <img alt="Style" src="https://img.shields.io/badge/Style-House_conventions-6f42c1">
  <img alt="Tests" src="https://img.shields.io/badge/Tests-Test__STATS__PROBDIST__RunAll-orange">
  <img alt="Verified against" src="https://img.shields.io/badge/Verified-SciPy_%2F_mpmath-00A3E0">
  <img alt="License" src="https://img.shields.io/badge/License-MIT-green">
</p>

Thanks for your interest in improving the library. This is a numerical project,
and it values **well-documented and numerically verified**
changes that match the existing conventions over large rewrites. This guide
explains how to work with the VBA source, and — just as importantly — what a
change to a distribution or kernel has to prove before it can be merged.

---

## 💬 Before you start

<p align="left">
  <img alt="Step" src="https://img.shields.io/badge/Step-Open_an_issue_first-217346">
</p>

Please **open an issue before starting non-trivial work** so the approach can be
agreed up front. Good issues to raise:

- a clear bug with reproduction steps (Excel version, 32/64-bit, OS, the exact
  `K_STATS_*` call and arguments, and the returned value or error code)
- a numerical-accuracy report (observed value, expected value, and the
  independent reference you computed it against)
- a focused enhancement with a concrete use case (a new distribution, moment,
  or tail routine)
- a documentation gap or inaccuracy

Tiny fixes (typos, comment corrections, obvious one-line bugs) can go straight to
a pull request without a prior issue.

---

## 🧰 Project layout and toolchain

<p align="left">
  <img alt="Source" src="https://img.shields.io/badge/Source-Exported_VBA-217346">
  <img alt="Tool" src="https://img.shields.io/badge/Tool-GitHub_Desktop-blue">
  <img alt="Add-in" src="https://img.shields.io/badge/Add--in-None-lightgrey">
</p>

The repository stores **exported VBA source**, not a binary workbook. There is
**no add-in, no installer and no external DLL** — the modules are imported into
the VBE and used directly.

```text
src/M_STATS_PROBDIST_CORE.bas          # shared constants, predicates, safe Exp/Log, inverse-normal seed, status writer
src/M_STATS_PROBDIST_SPECIALFUNCS.bas  # reusable special-function kernels (gamma, regularized beta/gamma, erf, ...)
src/M_STATS_PROBDIST_NORMALFAMILY.bas  # Standard Normal, Normal, Lognormal
src/M_STATS_PROBDIST_TFAMILY.bas       # Student t, Chi-square, F
src/M_STATS_PROBDIST_CONTINUOUS.bas    # Gamma, Beta, Exponential, Weibull, Uniform + cross-family identities
tests/M_STATS_PROBDIST_TEST.bas        # consolidated regression harness (counters, assertions, suites)
docs/                                  # code review and supporting notes
examples/                              # usage examples
```

The two lower layers (`CORE`, `SPECIALFUNCS`) are the numerics foundation; the
family modules consume them without keeping private duplicate copies. Read a few
procedures in `M_STATS_PROBDIST_CORE.bas` and `M_STATS_PROBDIST_NORMALFAMILY.bas`
before contributing — they are the reference for both style and structure.

You do not need the git command line. The maintainer works through **GitHub
Desktop**, and that is the recommended workflow for contributors too.

---

## 🔁 Edit and export workflow

<p align="left">
  <img alt="Flow" src="https://img.shields.io/badge/Flow-Import_Edit_Export-217346">
</p>

Because the source lives as exported files, the working loop is:

1. Import the relevant `.bas` files into an Excel workbook through the VBE
   (`File → Import File...`). The dependency order is `CORE` → `SPECIALFUNCS` →
   family modules → `TEST`.
2. Make your change in the VBE and **compile** (`Debug → Compile VBAProject`)
   until it is clean.
3. Run the regression harness (see below) and, for anything numerical, validate
   against an independent reference (see **Numerical verification**).
4. **Re-export** each changed module (`File → Export File...`) back over the
   matching file in `src/` (or `tests/`), preserving the existing layout.
5. Commit the changed text files only.

Do not commit the host workbook you used for editing.

---

## 🧱 Coding standards (house style)

<p align="left">
  <img alt="Explicit" src="https://img.shields.io/badge/Option-Explicit_required-217346">
  <img alt="Banners" src="https://img.shields.io/badge/Doc-Banner_per_procedure-blue">
  <img alt="Prefixes" src="https://img.shields.io/badge/Naming-Prefix_namespaced-6f42c1">
</p>

New code must match the existing conventions exactly.

**Module hygiene**

- `Option Explicit` at the top of every module.
- `Option Private Module` on the **numerics layers** (`CORE`, `SPECIALFUNCS`) so
  their `PROB_*` surface stays project-visible but invisible to the worksheet.
- Do **not** add `Option Private Module` to the distribution-family modules —
  their `K_STATS_*` functions must remain worksheet-visible.

**Banners**

Every module and every procedure carries a banner doc-block. Module banners open
with a full rule of 78 `=` characters after a single leading quote; sub-rules use
`-`. Procedure banners use these labels:

```vb
'==============================================================================
' PROCEDURE_NAME
'------------------------------------------------------------------------------
' PURPOSE
'   ...
' INPUTS
'   ...
' RETURNS
'   ...
' BEHAVIOR
'   ...
' ERROR POLICY
'   ...
' DEPENDENCIES
'   ...
' UPDATED
'   YYYY-MM-DD
'==============================================================================
```

Where relevant, keep the `WHY THIS EXISTS`, `ALGORITHM PROVENANCE` and
`DESIGN PRINCIPLES` sections — provenance statements must be **honest and
specific** (e.g. "Acklam's rational approximation, ~1.15E-9, used as a
root-finder seed, not a final answer"), never a vague "high accuracy".

**Body structure**

Order the body with sectioned sub-banners in this sequence, using only those a
routine needs:

```text
DECLARE / INITIALIZE / VALIDATE INPUTS / COMPUTE / RETURN SUCCESS / FAIL - NUMERIC / ERROR HANDLER
```

Put a short intent comment **above** each meaningful statement. Align
declarations `Dim Name As Type  'comment`. Use the literal forms `0#`, `1#` and
`vbNullString`.

**Naming and prefixes**

| Prefix | Scope | Return type |
| --- | --- | --- |
| `K_STATS_` | public worksheet-facing UDFs | `Variant` (so they can return `CVErr`) |
| `PROB_` | project-scoped private numerical kernels and helpers | `Double` |
| `Test_STATS_PROBDIST_` | regression-harness entry points | — |

**Error-handling contract**

The public error policy is a contract the tests enforce — match it exactly:

- `On Error GoTo Err_Handler` at the top of the executable body.
- A predictable numerical failure (invalid domain, non-finite input, density
  pole, overflow, non-convergence) does `GoTo Fail_Num`, and the `Fail_Num:`
  block returns `CVErr(xlErrNum)` (`#NUM!`).
- The `Err_Handler:` block returns `CVErr(xlErrValue)` (`#VALUE!`) for
  unexpected runtime errors only.
- Underflow of an exponential is a **valid zero**, not an error.
- Reserve `On Error Resume Next` for genuinely best-effort `Try*` primitives, not
  for normal logic.

**Kernel discipline**

- `PROB_*` kernels never validate their callers' domains and never write Status.
- Use `PROB_TryExp` / `PROB_TryMultiply` / `PROB_TryAdd` / `PROB_TryDivide` for
  overflow-aware arithmetic rather than raw operators where range is a concern;
  guard overflow with `PROB_DOUBLE_MAX`.
- Use `PROB_Log1p` / `PROB_Expm1` for left-tail precision (Exponential, Weibull)
  instead of `Log(1 + x)` / `Exp(x) - 1`.
- Beware non-short-circuit `And`/`Or` in guards: VBA evaluates both sides, so a
  left condition does **not** protect a right-side division. Split the guard.

---

## 🔬 Numerical verification (required)

<p align="left">
  <img alt="Reference" src="https://img.shields.io/badge/Reference-SciPy_%2F_mpmath-217346">
  <img alt="Precision" src="https://img.shields.io/badge/Precision-40--60_digits-blue">
  <img alt="Literals" src="https://img.shields.io/badge/Test_literals-Computed_not_typed-red">
</p>

This is the part that makes this project different from a thin wrapper. **Every
expected value used to assert correctness must be computed independently — never
estimated, rounded from memory, or typed by hand.**

- Verify new or changed results against an independent reference, typically
  Python with **SciPy** and **mpmath** at 40–60 significant digits, and check the
  density, CDF, survival, and inverse across the body **and** both tails.
- Cross-family identities (e.g. χ²(2) ≡ Exponential, F ≡ Beta relationships) are
  welcome as genuine kernel-level checks — but make sure they route through
  independent code paths rather than reducing to a tautology that both sides
  satisfy trivially.
- Any reference constant that appears in an `AssertClose` / `AssertRelClose` must
  be the computed value, to the precision the assertion needs. A placeholder
  literal is a defect, not a TODO.
- State the accuracy you are claiming, and how you verified it, in the PR.

---

## 🧪 Testing

<p align="left">
  <img alt="Harness" src="https://img.shields.io/badge/Harness-M__STATS__PROBDIST__TEST-217346">
  <img alt="Entry" src="https://img.shields.io/badge/Entry-Test__STATS__PROBDIST__RunAll-blue">
</p>

Import `tests/M_STATS_PROBDIST_TEST.bas` and run the suite from the VBE
**Immediate window** (`Ctrl+G`) before submitting. All entry points are
argument-less public subs, so they also run via `F5` or `Alt+F8`:

```vb
Test_STATS_PROBDIST_RunAll            ' full suite
Test_STATS_PROBDIST_RunCore           ' constants, predicates, primitives, inverse-normal seed
Test_STATS_PROBDIST_RunNormalFamily   ' Standard Normal, Normal, Lognormal
Test_STATS_PROBDIST_RunTFamily        ' Student t, Chi-square, F
Test_STATS_PROBDIST_RunContinuous     ' Gamma, Beta, Exponential, Weibull, Uniform
```

Results are written with `Debug.Print`: passing assertions are **silent**, each
failure prints one detailed line, and a consolidated PASS/FAIL summary closes the
run.

- All existing suites must still pass.
- If you change or add behavior, extend the matching suite with the existing
  `Assert*` helpers. Assertions accept `Variant` and reject unexpected worksheet
  errors explicitly; use the error-code assertions to pin `#NUM!` vs `#VALUE!`
  where the numerical contract requires a predictable failure.
- Never weaken a genuine assertion merely to make a failing implementation look
  green.

---

## 📚 Documentation expectations

<p align="left">
  <img alt="Docs" src="https://img.shields.io/badge/Docs-Kept_in_sync-6f42c1">
</p>

Documentation is part of the change, not a follow-up. When your change affects a
user-facing surface, update the matching pages in the same pull request:

- **public UDF surface** changed → the README function table **and** the
  [API-Reference](../../wiki/API-Reference) wiki page
- **kernels / algorithms** changed → [Special-Functions-and-Numerical-Kernels](../../wiki/Special-Functions-and-Numerical-Kernels)
  and [Numerical-Accuracy-and-Design](../../wiki/Numerical-Accuracy-and-Design)
- **error behavior** changed → [Error-Handling-and-Diagnostics](../../wiki/Error-Handling-and-Diagnostics)
- **new distribution / family** → the relevant family page and the
  [Module-Reference](../../wiki/Module-Reference)

---

## 📦 What not to commit

<p align="left">
  <img alt="Excluded" src="https://img.shields.io/badge/Excluded-Workbooks_and_Locks-red">
</p>

- The Excel workbook you used to edit or run the source.
- Excel lock / owner files (`~$*`).
- Local scratch sheets, generated outputs, or machine-specific paths.

See `.gitignore` for the full list.

---

## 🚀 Submitting changes

<p align="left">
  <img alt="PR" src="https://img.shields.io/badge/PR-Small_and_focused-217346">
</p>

1. Fork the repository and create a branch for your change.
2. Keep the pull request **small and focused** — one logical change per PR.
3. In the PR description, state the problem, the approach, **the independent
   reference you validated against**, and which suites you ran.
4. Confirm: project compiles cleanly, `Test_STATS_PROBDIST_RunAll` passes,
   banners / naming / error policy follow the house style, numerical results are
   verified, and docs are updated.

The maintainer reviews changes selectively and may adopt, adapt, or decline a
contribution to keep the library coherent. Clear, well-scoped, numerically
justified PRs are the most likely to be merged.

---

## 📄 License

By contributing, you agree that your contributions are licensed under the
project's **MIT License**.

---

## 👤 Maintainer

Maintained by **Daniele Penza**. For anything that is not a code change — design
questions, larger proposals, or general feedback — open an issue to start the
conversation.
