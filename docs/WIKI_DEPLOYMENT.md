# Documentation Deployment Guide

This package was prepared from the current `main` repository state retrieved through the authenticated GitHub connector.

- Repository: `danielep71/VBA-PROBABILITY-DISTRIBUTIONS`
- Source branch: `main`
- Source commit reviewed: `c632897590782aca1f05776073a25ff969fcad52`
- Documentation date: `2026-07-22`

## Main README

Replace the root `README.md` with the packaged file.

```bash
git clone https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS.git
cd VBA-PROBABILITY-DISTRIBUTIONS
cp /path/to/package/README.md README.md
git add README.md
git commit -m "Refresh README for complete discrete family"
git push origin main
```

Review the Markdown rendering before pushing directly to `main`. A documentation branch and pull request can be used instead.

## GitHub Wiki

GitHub stores the Wiki in a separate Git repository.

```bash
git clone https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS.wiki.git
cd VBA-PROBABILITY-DISTRIBUTIONS.wiki
cp /path/to/package/wiki/*.md .
git add .
git commit -m "Refresh probability library Wiki"
git push origin master
```

Some Wiki repositories use `main` rather than `master`; check the cloned default branch before pushing.

The package preserves the existing page slugs referenced by the repository README:

```text
Home
Getting-Started
Architecture
Module-Reference
API-Reference
Normal-and-Lognormal-Family
StudentT-ChiSquare-and-F-Family
Continuous-Distributions
Discrete-Distributions
Special-Functions-and-Numerical-Kernels
Numerical-Accuracy-and-Design
Benchmarking-and-Accuracy-Contracts
Repository-Structure
Error-Handling-and-Diagnostics
Testing-and-Regression-Harness
Excel-VBA-CI
Troubleshooting
```

`_Sidebar.md` and `_Footer.md` are also included.

## Browser-based Wiki update

For each Wiki page:

1. open the page;
2. choose **Edit**;
3. replace the body with the matching Markdown file;
4. save with a clear change message.

Create missing pages using the exact filename stem to preserve links.

## Review checklist

```text
[ ] README image and badges render
[ ] Wiki sidebar appears
[ ] All internal Wiki links resolve
[ ] Discrete catalogue lists all six families
[ ] Discrete Uniform parameterization is correct
[ ] Test suite lists RunDiscreteSuite
[ ] CI documentation distinguishes Excel CI from the hosted accuracy gate
[ ] Benchmark documentation states that hosted CI checks committed observations
[ ] Multivariate distributions are marked out of scope
[ ] Random and array APIs are clearly marked as roadmap items
[ ] Source commit or release tag is recorded in the documentation change
```
