# Normal upper-tail kernel study (#72)

**Status: recorded PASS under Amendment 1, retained; the one-shot holdout
record is enforced (#80); study acceptance stays open under #78 and #79.**
Excel-free. This study
designs and validates a replacement for `PROB_NormalUpperTailPositive` before
any VBA is written. It is not evidence that the kernel is fixed. The fix lands
only in a real-Excel wave, together with source, regression tests, exports
and truthful provenance.

Design: `PREREGISTRATION.md`, frozen before the committed derivation. Its §0
discloses the scratch exploration that chose the design. Amendment 1, written
after the first run, replaces one acceptance condition (A2) and is disclosed
as such. A dated note of 2026-11-03 records how the one-shot holdout rule is
now enforced; it is not an amendment.

## Defect

The current kernel is accurate in absolute terms but not in relative terms:
- In its Hart region (`3 ≤ z < 7.07`), relative error reaches 2.9E-9.
- Its exponential factor rounds `z²` first, which costs up to 5.7E-14 in the
  far tail.

Because of this, the survival contracts of `Normal_Survival`,
`Lognormal_Survival` and `NormalStandard_Survival` are limited to the central
region (`SurvivalTailRel`). It also blocks the #34 Student-t route (finding
F1).

## Candidate

- A degree-(7, 7) rational for `Q(z)·exp(z²/2)` on `[0, 3)`.
- A 60-level Laplace continued fraction on `[3, 38.5)`.
- An exponential factor split so that its main argument is exact.
- Every constant built from 15-significant-digit literal pairs, as the VBE
  requires.

The acceptance bound is `2^-50` relative, the bound of the #34 study.

## Result of the derivation (2026-10-13)

`derive_kernel.py` ran the preregistered fit in 74 s. `validate_kernel.py`
applied the design and wrote `results.json`.

**First run: failed A2 only (#74).** A2 required every coefficient to be
positive. The numerator's `z^7` coefficient is −2.54E-10. A2 exists so that
Horner evaluation cannot cancel, and the term's actual effect is a
cancellation factor of only 1 + 1.5E-7 on `[0, 3]`. The scratch prototype
(§0) had the same sign; its positivity check left out the last numerator
coefficient. Under §6 the study failed, and the holdout was not evaluated.

**Amendment 1.** The maintainer chose an amendment over a new
preregistration. A2 is replaced by A2′, which bounds the cancellation factor
by `1 + 2^-20`. The coefficients were not refitted, and nothing else changed.
The frozen A2 stays recorded as `A2_as_frozen`.

**Second run: PASS.** The committed record reports one evaluation of the
holdout. The validator of that run could neither prevent nor detect a repeat;
[Holdout exposure record](#holdout-exposure-record-80) describes the controls
that now enforce it.

| Condition | Result |
| --- | --- |
| A1 approximant | 1.03E-17 ≤ 2^-55 on 23,001 points |
| A2′ cancellation (Amendment 1) | numerator 1 + 1.5E-7, denominator 1 (all positive), limit 1 + 2^-20 |
| A3 literals | all 17 constants rebuild their Double exactly |
| A4 continued fraction | truncation 2.0E-18 at z = 3 ≤ 2^-56, decreasing to 4.3E-50 at z = 10 |
| A5 kernel on V (23,751 points + seams) | worst Q 5.3E-16, 1 − Q 3.9E-16, density 3.4E-16, all ≤ 2^-50 (8.9E-16) |
| A6 holdout H (15,000 points, evaluated once) | worst Q 5.5E-16, 1 − Q 3.9E-16, density 3.3E-16 |
| A7 #34 route | worst 3.7E-16 at the 36 fitting points; 1.8E-16 at its 21 holdout points |
| A8 no regression | holds at all 3,227 points where today's kernel meets the bound |
| Oracle stability | ≥ 58.0 digits; quadrature leg ≥ 98.3 digits at 151 points |

Worst relative error by band, today's kernel against the candidate:

| band | today | candidate |
| --- | ---: | ---: |
| `0 ≤ z < 3` | 1.5E-14 | 5.3E-16 |
| `3 ≤ z < 7.07` | 2.9E-9 | 4.4E-16 |
| `7.07 ≤ z < 37.5` | 5.7E-14 | 5.0E-16 |

**Predictions for the Excel wave (§7).** All are recorded in `results.json`,
and none is outside its bound:
- P1, 41 direct rows: worst 3.7E-16.
- P2, 27 standardised rows: worst 2.3E-16.
- P3, 11 interval rows: worst 1.0E-16.
- P4, 80 inverse rows: worst 2.9E-15, with no predicted contract breach.
  The `split_boundary` inverse rows fall from 2.8E-11–3.8E-10 today to the
  binary64 floor (≤ 4.1E-16).

The model of today's code reproduces every affected committed Excel
observation within 4 ulps.

**Two defects in the validator itself were corrected before the result above
was recorded.** Neither touches the design or a threshold:
- The quadrature leg integrated `exp(−t²/2)` directly over `[z, ∞)`. That
  agreed to only 4.3 digits at z = 37.5. It now integrates
  `exp(−z s − s²/2)` on a `1/(z + 1)` scale, and agrees to ≥ 98 digits.
- `Normal_IntervalProbability` rows were read as `(lower, upper, mean)`. The
  density-helper exporter calls them as `(lower, upper, 0, sd)`.

**Next.** The kernel is ready to be implemented in VBA with the recorded
literals, in the Excel wave that §9 describes. That wave tests the §7
predictions against real Excel. Before that, #78 (conditioning rules of the
P2 predictions) and #79 (fit convergence and result provenance) must close.

## Holdout exposure record (#80)

The preregistration evaluates the holdout H once (§5). Until #80 that rule
was kept only by convention: every run of `validate_kernel.py` evaluated H
again and rewrote `evaluated_once: true`, and the fixture trusted that
Boolean. Nothing shows that a repeat happened, but nothing could have shown
one.

**I kept the recorded PASS and its evidence exactly.** I added a
`holdout_exposure` block to `results.json` after the last recorded value;
every recorded measurement, including the A6 block of #75, is byte-for-byte
unchanged. I did not evaluate H to make this change, and I did not refit,
change a threshold or change the design. The block records:
- `status` `EVALUATED` with one exposure, and `started_utc`, the time of the
  recorded run, which wrote V immediately before evaluating H;
- the candidate: the SHA-256 of `kernel_coefficients.json`;
- the design: the byte length and SHA-256 of `PREREGISTRATION.md` as
  committed with the record. Later amendments may only append to it;
- the input: H's seeds, counts and ranges, and the SHA-256 of its 15,000
  seeded points. Generating the points evaluates nothing on them;
- `a6_sha256`, the SHA-256 of the canonical A6 block;
- the SHA-256 of the validator and model committed with the record, for
  information. Binding the current code to the results is #79.

The exposure status is one of:

| `holdout_exposure.status` | meaning | `acceptance.A6` |
| --- | --- | --- |
| `NOT_EVALUATED` | H has not been evaluated | absent |
| `EVALUATION_STARTED` | a first evaluation began and recorded no result; H counts as exposed | absent |
| `EVALUATED` | H was evaluated once | present, matching `a6_sha256` |

**Controls.** `holdout_record.py` enforces the record, and
`validate_kernel.py` uses it:
- Every mode reads and checks the existing record before it computes or
  writes anything. A missing, malformed or self-contradicting record is
  refused, never reset. So is a record whose candidate, holdout definition or
  frozen preregistration text no longer matches the files.
- `--first-evaluation` runs only from `NOT_EVALUATED`. It writes V with
  `EVALUATION_STARTED` before it evaluates H, so a crash leaves the exposure
  recorded. A second attempt is refused before anything is computed, and the
  output stays byte-identical.
- `--first-evaluation`, `--regenerate` and `--init` first take an exclusive
  lock: they create `.results.json.lock` with `O_CREAT | O_EXCL` before they
  read the record, and hold it through every write and the holdout
  evaluation until they return or fail. A second run is refused while the
  lock exists, before it reads, computes or evaluates anything, so two
  concurrent first evaluations evaluate H once. `--verify` takes no lock, but
  is refused while one exists.
- Writes are atomic: an exclusively created temporary file, then a rename. A
  temporary file left by a killed run blocks every mode until you inspect it.
  A write is also refused if the output changed while the run was computing;
  the check runs after the temporary file is created.

**If a run is killed while it holds the lock**, the lock file stays behind,
because it is a plain file rather than an OS lock. Every mode is then refused
and names the lock. To recover, make sure no validator is running, check
`holdout_exposure.status` in `results.json`, then delete
`.results.json.lock` by hand. Deleting the lock never reopens the holdout: an
`EVALUATION_STARTED` record still refuses a re-evaluation. Exclusive creation
is atomic on local POSIX and Windows (NTFS) file systems, but may not be on
some network file systems, so run the validator on a local checkout.
- `test_kernel_model.py` pins the A6 block by its SHA-256 and figures, and
  checks the record and its identity, so CI rejects a replaced A6.
- `test_holdout_record.py` exercises first evaluation, second-run refusal,
  interrupted runs, concurrent runs, stale locks and evidence preservation on
  synthetic studies. It never evaluates the real H.

**If a first evaluation is interrupted**, the record stays at
`EVALUATION_STARTED` with status FAIL. No mode evaluates H again from there.
The holdout counts as spent, and only a new, disclosed preregistration with a
new holdout can replace it.

## Running the validator

Run these from this directory. Each one needs an explicit mode:

```sh
python3 validate_kernel.py --verify
python3 validate_kernel.py --regenerate --reason "TEXT"
python3 validate_kernel.py --first-evaluation
python3 validate_kernel.py --init
```

- `--verify` is a read-only replay. It recomputes everything except H (A1–A5,
  A7, A8, the oracles, seams, bands and §7 predictions) and compares it with
  `results.json`. It never evaluates H and never writes. It reports the
  recorded A6 as retained, not as a new measurement, so a replay cannot pass
  for an independent evaluation. It takes about 35 s.
- `--regenerate --reason "TEXT"` recomputes everything except H and rewrites
  `results.json`. It carries `holdout_exposure` and `acceptance.A6` over
  unchanged, and appends `{utc, reason, holdout_evaluated: false}` to
  `regenerations`. The status is derived again from the recomputed
  conditions and the retained A6.
- `--first-evaluation` evaluates H once. For this study it is refused: the
  exposure is `EVALUATED`.
- `--init` is for a new study only. It writes V with `NOT_EVALUATED` when no
  results file exists.

The exit status is 0 on success, 1 on a study failure or replay mismatch, 2 on
a usage error and 3 on a refusal. A refusal leaves the output unchanged, and
says so with its SHA-256.

To change the §7 predictions, for example after amending the conditioning
rules under #78, change `validate_kernel.py`, append the amendment to
`PREREGISTRATION.md`, then run:

```sh
python3 validate_kernel.py --regenerate --reason "#78: <amendment>"
python3 validate_kernel.py --verify
python3 test_kernel_model.py
python3 test_holdout_record.py
```

H is not evaluated at any step, and the recorded A6 is kept.

## Files

| File | Role |
| --- | --- |
| `PREREGISTRATION.md` | frozen design, acceptance conditions, predictions for the Excel wave, Amendment 1 and the dated #80 note |
| `derive_kernel.py` | the preregistered fit; writes `kernel_coefficients.json` |
| `kernel_coefficients.json` | coefficients to 40 digits, their Doubles, and VBA `hi`/`lo` literal pairs |
| `kernel_model.py` | binary64 model of the candidate, of today's kernel, and of every affected public path |
| `validate_kernel.py` | applies §3–§8; verifies, regenerates or writes `results.json` |
| `holdout_record.py` | the holdout exposure record and its checks (#80); standard library only |
| `results.json` | every measurement, the acceptance record, the per-row predictions and the holdout exposure record |
| `test_kernel_model.py` | CI fixture pinning the coefficients, literals, recorded result and recorded A6 |
| `test_holdout_record.py` | CI fixtures of the exposure controls, on synthetic studies |
