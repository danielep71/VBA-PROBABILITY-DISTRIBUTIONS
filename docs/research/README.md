# Numerical research and evaluation

Review date: **2026-11-08**. Tracking: [#265](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/issues/265).

This directory records relevant publications, their possible application to the
library, and the evidence needed before adoption. It contains original summaries
and links, not copies of publisher articles or third-party implementations.

| Document | Purpose |
| --- | --- |
| [Bibliography](BIBLIOGRAPHY.md) | Twenty-six source records (2000–2026), publication dates, access depth and transfer limitations |
| [Evaluation methodology](METHODOLOGY.md) | Reproducible comparison and adoption protocol |
| [Issue map](ISSUE_MAP.md) | Existing implementation owners, roadmap links and feasibility work |
| [Article note template](ARTICLE_NOTE_TEMPLATE.md) | Fields for future reviews and decisions |

## Scope and status

The search covers publications and author preprints from 2000–2026 on incomplete
beta/gamma functions, inverse distributions, extreme tails, discrete distributions,
bivariate probabilities and compound frequency-severity models. Publication years
and preprint revisions are recorded separately. This is a selected review, not an exhaustive systematic review.
Publisher pages, author manuscripts and institutional sources were preferred.
Access depth is recorded per article; discovery does not imply full-text review.

The current design uses shared incomplete-beta/gamma kernels, direct tails and
safeguarded inverses. The papers suggest experiments; they do not establish that
the current VBA implementation is defective or that a replacement is superior.
The issue map records known defects separately from exploratory improvements.

All twenty-six records are **reviewed for relevance, not adopted or validated in VBA
by this review**. No benchmark, runtime measurement or new Excel run was performed
for this documentation. No contract threshold or frozen preregistration changes.

## Extension to 2000–2020

R18–R26 add nine references without duplicating the 2020 records already listed.
For Gamma P/Q precision work, prioritize R18 (formulas and dispatch regions) and
R19 (two-Double primitive error analysis), with R20/R22 as arithmetic comparisons
and R21 as an alternative integral representation. These are research inputs,
not a claim that any current PR meets numerical acceptance.

## Maintaining the review

1. Assign a stable R-number and record the exact source/version consulted.
2. Separate the publication's claims from the project's hypothesis and results.
3. Check open and closed issues before adding work; link the existing owner.
4. Use the methodology before experiments, preserving existing frozen protocols.
5. Record adopt, defer or reject decisions with evidence and limitations.

Full text and implementation code require their own redistribution/license
assessment before being copied. Open access alone is not a blanket code license.
Numerical experiment assets belong in `benchmark/`, with source binding and
provenance; this directory is the research and decision index.

Repository rules remain authoritative: [contributing](../../CONTRIBUTING.md),
[benchmark provenance](../../benchmark/PROVENANCE.md),
[Excel certification](../EXCEL_CERTIFICATION.md) and
[verification depth](../VERIFICATION_DEPTH.md).
