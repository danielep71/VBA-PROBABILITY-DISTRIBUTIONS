<!--
  Thanks for contributing! Keep PRs small and focused — one logical change each.
  See CONTRIBUTING.md for the house style, the error-policy contract, and the
  numerical-verification requirement.
-->

## 📋 Summary

What does this change do, and why?

## 🔗 Related issue

Closes #<!-- issue number, if any -->

## 🧩 Type of change

- [ ] 🐞 Bug fix (wrong value / wrong error code)
- [ ] 🎯 Numerical-accuracy improvement
- [ ] ✨ New distribution, moment, or routine
- [ ] ♻️ Refactor (no behavior change)
- [ ] 📚 Documentation
- [ ] 🧪 Tests / harness

## 🔬 Numerical verification

Independent reference used (e.g. SciPy / mpmath at N digits), and the points
checked (body + both tails, density / CDF / survival / inverse):

```text

```

## 🧪 How it was tested

Which suites did you run, and what was the result?

```text
Test_STATS_PROBDIST_RunAll          →
Test_STATS_PROBDIST_RunContinuous   →   (if a continuous distribution changed)
Test_STATS_PROBDIST_RunTFamily      →   (if t / chi-square / F changed)
Test_STATS_PROBDIST_RunNormalFamily →   (if normal / lognormal changed)
Test_STATS_PROBDIST_RunCore         →   (if a kernel / primitive changed)
```

## ✅ Checklist

- [ ] Project compiles cleanly (`Debug → Compile VBAProject`)
- [ ] `Test_STATS_PROBDIST_RunAll` passes
- [ ] New/changed procedures have a banner doc-block and follow the naming and
      error-handling contract (`K_STATS_*` → `Variant`/`CVErr`, `PROB_*` →
      `Double`; `Fail_Num:` → `#NUM!`, `Err_Handler:` → `#VALUE!`)
- [ ] Every new expected value in an assertion is **computed independently**, not
      typed by hand
- [ ] Changed modules were **re-exported** to `src/` / `tests/` (no editing
      workbook committed)
- [ ] Docs updated where relevant (README function table, `API-Reference`,
      `Numerical-Accuracy-and-Design`, `Error-Handling-and-Diagnostics`)
- [ ] No workbooks or lock files committed (`~$*`)

## 📎 Notes for the reviewer

Trade-offs, provenance of an algorithm, or follow-ups that need attention.
