<!--
Thank you for contributing.

Keep the pull request small and focused: one logical numerical, API, testing,
or documentation change per PR.

Read CONTRIBUTING.md before completing this template.
-->

## 📋 Summary

Describe the change and why it is needed.

## 🔗 Related issue

```text
Closes #
```

Remove this section when no issue applies.

## 🧩 Type of change

- [ ] 🐞 Numerical or functional bug fix
- [ ] 🎯 Numerical-accuracy improvement
- [ ] ✨ New distribution, moment, tail, interval, or helper
- [ ] ♻️ Refactor with no intended behavioral change
- [ ] 🧪 Test-harness or regression-test change
- [ ] 📚 Documentation-only change
- [ ] 🔧 Repository or maintenance change

## 📐 Public behavior and parameterization

Describe any change to:

- function names or signatures;
- parameter order;
- rate versus scale conventions;
- support boundaries;
- endpoint behavior;
- worksheet-error classification;
- optional `Status` diagnostics;
- worksheet visibility.

```text
Public behavior changed:
Backward compatible:
Parameterization:
```

Write `No public behavior change` when applicable.

## 🧠 Numerical method

For numerical changes, explain:

- the formula or algorithm;
- why the previous implementation was insufficient;
- cancellation, overflow, underflow, or convergence risks;
- tail orientation;
- algorithm provenance;
- expected accuracy.

```text
Method:
Previous failure mode:
Stability treatment:
Provenance:
Accuracy target:
```

Write `Not applicable` for documentation-only or repository-only changes.

## 🔬 Independent verification

Identify the independent reference used.

```text
Reference system:
Version:
Precision:
Functions or formulas:
Parameter grid:
Central-region checks:
Lower-tail checks:
Upper-tail checks:
Maximum observed error:
```

Expected values in the VBA test harness must be calculated independently, not
copied from the implementation being tested.

Write `Not applicable` only when the change has no numerical effect.

## 🧪 Testing performed

```text
Debug → Compile VBAProject             →
Test_STATS_PROBDIST_RunCore             →
Test_STATS_PROBDIST_RunNormalFamily     →
Test_STATS_PROBDIST_RunTFamily          →
Test_STATS_PROBDIST_RunContinuous       →
Test_STATS_PROBDIST_RunAll              →
```

Include the relevant Immediate Window output for a failure fix or substantial
numerical change.

## ✅ Contract checklist

### Source and compilation

- [ ] The VBA project compiles cleanly.
- [ ] Changed modules were re-exported to `src/` or `tests/`.
- [ ] No editing workbook, lock file, or confidential data is included.
- [ ] The textual diff contains only intended changes.

### Naming and return contracts

- [ ] Worksheet-facing statistical UDFs use the established `K_STATS_*` naming.
- [ ] Ordinary worksheet-facing UDFs return `Variant` so they can return `CVErr`.
- [ ] Any specialized `Double`-returning public helper has an explicit documented
      exception contract.
- [ ] `PROB_Is*` predicates return `Boolean`.
- [ ] `PROB_Try*` routines return `Boolean` and provide results through `ByRef`.
- [ ] Other `PROB_*` routines use the explicitly documented type appropriate to
      their purpose; they are not assumed universally to return `Double`.
- [ ] Test entry points follow `Test_STATS_PROBDIST_*` and are argument-less Subs.

### Error and numerical behavior

- [ ] Invalid domains and predictable numerical failures return `#NUM!`.
- [ ] Unexpected runtime failures return `#VALUE!`.
- [ ] Valid exponential underflow returns zero where mathematically appropriate.
- [ ] No numerical UDF displays a `MsgBox`.
- [ ] Direct survival calculations are used where `1 - CDF` would lose precision.
- [ ] Overflow-sensitive arithmetic uses the shared guarded helpers.
- [ ] Non-short-circuit `And` or `Or` does not expose a protected division or
      otherwise faulting expression.
- [ ] Iterative kernels never publish a non-converged partial result.

### Documentation and tests

- [ ] Every changed or new procedure has an accurate structured header.
- [ ] Algorithm provenance and accuracy claims are specific and supportable.
- [ ] A numerical bug fix includes a regression test.
- [ ] New expected values were calculated independently.
- [ ] All existing suites still pass.
- [ ] README and Wiki pages were updated where relevant.
- [ ] Public counts, examples, and parameter tables remain synchronized.

## 📚 Documentation updated

Check all that apply:

- [ ] README
- [ ] API Reference
- [ ] Architecture
- [ ] Module Reference
- [ ] Distribution-family page
- [ ] Special Functions and Numerical Kernels
- [ ] Numerical Accuracy and Design
- [ ] Error Handling and Diagnostics
- [ ] Testing and Regression Harness
- [ ] No documentation change required

## 📎 Reviewer notes

Describe trade-offs, known limitations, or follow-up work.
