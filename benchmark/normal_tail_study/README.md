# Normal upper-tail kernel study (#72)

**Status: recorded PASS under Amendment 1, retained; the one-shot holdout
record is enforced (#80); Amendment 2 restates the P2 conditioning rule
(#78); Amendment 3 defines fit convergence, which the fit meets, and binds
the result to the files that produce it (#79).**
Excel-free. This study
designs and validates a replacement for `PROB_NormalUpperTailPositive` before
any VBA is written. It is not evidence that the kernel is fixed. The fix lands
only in a real-Excel wave, together with source, regression tests, exports
and truthful provenance.

Design: `PREREGISTRATION.md`, frozen before the committed derivation. Its §0
discloses the scratch exploration that chose the design. Amendment 1, written
after the first run, replaces one acceptance condition (A2) and is disclosed
as such. A dated note of 2026-11-03 records how the one-shot holdout rule is
now enforced; it is not an amendment. Amendment 2, of the same date and
written after the holdout had been evaluated, restates the conditioning term
of the P2 bounds; see [Conditioning of the P2 bounds](#conditioning-of-the-p2-bounds-amendment-2-78).
Amendment 3, of the same date and also written after the holdout had been
evaluated, defines the fit convergence criterion and binds the result to its
generator, models, validator, design and inputs; see
[Fit convergence and result binding](#fit-convergence-and-result-binding-amendment-3-79).

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
| Fit convergence (Amendment 3) | CONVERGED: iterate 120 of 120 alternates at 16 points, levelled to 3.34E-3 ≤ 1% (first within 1% at iterate 53) |
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

**Next.** The study's acceptance evidence now covers every §6 failure
condition: A1–A8, the oracles and, under Amendment 3, fit convergence. The
record is bound to the files that produced it. The kernel can be implemented
in VBA with the recorded literals, in the Excel wave that §9 describes. That
wave tests the §7 predictions against real Excel. This study is not evidence
that the VBA kernel is fixed. Exact-source Excel validation remains open
under #72.

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
  information. The current code is bound by the `provenance` block (#79).

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

## Conditioning of the P2 bounds (Amendment 2, #78)

The frozen P2 rule bounds a standardised row's error by `B + κ·δ`. Two of
its expressions fail on covered rows:
- the survival condition number `z·φ(z)/Q(z)` is negative for `z < 0`, so
  the bound falls below `B` (four rows);
- the Lognormal `δ` divides by `|ln x − μ|`, which is zero at `x = 1, μ = 0`
  (two rows).

The validator had used `|z|` and dropped the zero-denominator term since the
first run (#74), without saying so. I appended Amendment 2 to
`PREREGISTRATION.md` to state and derive the rule. I wrote it after the
holdout had been evaluated (2026-10-13) and the predictions recorded, and it
says so. Nothing was refitted, no threshold changed, and the holdout was not
evaluated.

The restated rule is `error ≤ B + φ(z)·η / F(z)`:
- `F` is the survival `Q` or the cumulative `Φ` at the exact `z`, and
  `κ = |z|·φ(z)/F(z)` is its relative condition number `|z·F′/F|`, which is
  nonnegative for both signs;
- `η` bounds the absolute error of the computed `z`: `2^-52·|z|` for Normal,
  and `2^-52·(|z| + |ln x|/σ)` for Lognormal (two roundings and one ulp of
  `Log`).

Where `z ≠ 0`, `φ·η/F` is exactly the frozen `κ·δ`. At the centre, `Log(1)`
is exactly 0, `η = 0` and the bound is `B`. Near it, the term is finite and
tends to `√(2/π)·2^-52·|μ|/σ`. P3 inherits the same term for each tail.

After regenerating `results.json` (`--regenerate`, holdout not evaluated),
every P2 row records `z`, `κ` and `η`. Every predicted value, predicted
error and pass/fail outcome is unchanged. One bound,
`Lognormal_Cumulative(0.45, 0.2, 0.6)`, moved by one ulp because the same
quantity is now evaluated in another order. `test_conditioning.py` checks the
recorded rows against the frozen product. It also checks negative-`z`,
exact-centre (`x = 1, μ = 0`) and near-centre rows, the split at 3, and the
37.5 and 38.5 cut-offs.

## Fit convergence and result binding (Amendment 3, #79)

§6 makes a fit that does not converge a study failure, but §2 defined no
criterion. The derivation kept the best of 120 Lawson iterations and
recorded no decision. Separately, `results.json` bound only the coefficient
file, so a change to the generator, a model, the validator or an input grid
could leave recorded values stale while the fixture stayed green. I appended
Amendment 3 to `PREREGISTRATION.md` to close both gaps. I wrote it after the
fit and the holdout had been observed, and it says so. Nothing was refitted,
no threshold changed, and the holdout was not evaluated.

**Criterion.** The selected iterate's signed relative error at the 500 nodes
must alternate at `7 + 7 + 2 = 16` points and be levelled:
`node_max / L − 1 ≤ 0.01`. Here `L` is the de la Vallée Poussin lower bound on
the best node maximum any (7, 7) rational can reach. Otherwise the fit is
`CAP_EXHAUSTED`, if the node maximum was still falling at the frozen cap, or
`NOT_CONVERGED`. Either one fails the study, and so does a missing record or
status. I chose the 1% tolerance after seeing this fit's ratio, 3.34E-3, and
the amendment discloses that. A tolerance below 3.34E-3 would have failed
this fit.

**Evidence.** I re-ran the frozen fit with the evidence recorded. It
reproduced `kernel_coefficients.json` exactly, so the file is byte-identical,
and it wrote `fit_convergence.json`, which holds the node maximum, alternation
count and levelling ratio of all 120 iterates. Over those iterates:
- the node maximum fell at every iteration;
- every iterate alternated at 16 points;
- the ratio first reached 1% at iterate 53;
- the selected 120th iterate is levelled to 3.34E-3: CONVERGED.

`derive_kernel.py` refuses to write or accept coefficients for any other
status. The validator checks the record against the committed coefficients.
It also recomputes their levelled error and requires the criterion there.

**Binding.** `results.json` has a `provenance` block. It records the SHA-256
of:
- `derive_kernel.py` and `fit_convergence.json`;
- `kernel_model.py`, the #34 route model and its coefficients;
- `validate_kernel.py`, `holdout_record.py` and `result_binding.py`;
- the whole `PREREGISTRATION.md`;
- the three §7 grids, `accuracy_contracts.csv` and the #34 study's
  `results.json`.

The candidate stays bound by `coefficients_sha256`. `test_kernel_model.py`
and `test_result_binding.py` compare every identity with the files in CI, in
under a second. Any difference fails them, and `--verify` reports it as
STALE without recomputing.

A legitimate change to a bound file needs
`validate_kernel.py --regenerate --reason TEXT`. That run records the
identities that changed. It cannot re-bind the candidate or the frozen text,
because the holdout record refuses both. So a change to the main accuracy
grid, for example, now also requires this study to be regenerated.

## Running the validator

Run these from this directory. Each one needs an explicit mode:

```sh
python3 validate_kernel.py --verify
python3 validate_kernel.py --regenerate --reason "TEXT"
python3 validate_kernel.py --first-evaluation
python3 validate_kernel.py --init
```

- `--verify` is a read-only replay. If a bound file changed since the record
  was written, it reports STALE and stops. Otherwise it recomputes everything
  except H (fit convergence, A1–A5, A7, A8, the oracles, seams, bands and §7
  predictions) and compares it with `results.json`. It never evaluates H and never writes. It reports the
  recorded A6 as retained, not as a new measurement, so a replay cannot pass
  for an independent evaluation. It takes about 35 s.
- `--regenerate --reason "TEXT"` recomputes everything except H and rewrites
  `results.json`. It carries `holdout_exposure` and `acceptance.A6` over
  unchanged, and appends `{utc, reason, holdout_evaluated: false,
  identities_changed}` to `regenerations`. The status is derived again from the recomputed
  conditions and the retained A6.
- `--first-evaluation` evaluates H once. For this study it is refused: the
  exposure is `EVALUATED`.
- `--init` is for a new study only. It writes V with `NOT_EVALUATED` when no
  results file exists.

The exit status is 0 on success, 1 on a study failure or replay mismatch, 2 on
a usage error and 3 on a refusal. A refusal leaves the output unchanged, and
says so with its SHA-256.

To change the §7 predictions, as Amendment 2 did under #78, change
`validate_kernel.py`, append the amendment to `PREREGISTRATION.md`, then run:

```sh
python3 validate_kernel.py --regenerate --reason "#<issue>: <amendment>"
python3 validate_kernel.py --verify
python3 test_kernel_model.py
python3 test_holdout_record.py
python3 test_conditioning.py
python3 test_fit_convergence.py
python3 test_result_binding.py
```

H is not evaluated at any step, and the recorded A6 is kept. The same
sequence applies to any change to a bound file (Amendment 3).

To check the fit itself, run `python3 derive_kernel.py --verify`. It fits
again (about 40 to 75 s), compares the result with `kernel_coefficients.json`
and `fit_convergence.json`, and writes nothing.

## Files

| File | Role |
| --- | --- |
| `PREREGISTRATION.md` | frozen design, acceptance conditions, predictions for the Excel wave, Amendment 1, the dated #80 note, Amendment 2 (#78) and Amendment 3 (#79) |
| `derive_kernel.py` | the preregistered fit and its convergence decision; writes `fit_convergence.json`, and `kernel_coefficients.json` only for a converged fit that does not replace a recorded one |
| `kernel_coefficients.json` | coefficients to 40 digits, their Doubles, and VBA `hi`/`lo` literal pairs |
| `fit_convergence.json` | the convergence status, the criterion and the per-iteration evidence of the fit (Amendment 3) |
| `kernel_model.py` | binary64 model of the candidate, of today's kernel, and of every affected public path |
| `validate_kernel.py` | applies §3–§8; verifies, regenerates or writes `results.json` |
| `holdout_record.py` | the holdout exposure record and its checks (#80); standard library only |
| `result_binding.py` | the identities `results.json` binds, and the check that they are current (#79); standard library only |
| `results.json` | every measurement, the acceptance record, the per-row predictions, the provenance block and the holdout exposure record |
| `test_kernel_model.py` | CI fixture pinning the coefficients, literals, recorded result and recorded A6 |
| `test_conditioning.py` | CI fixture of Amendment 2: the recorded P2 bounds, negative-`z`, centre, near-centre and seam rows |
| `test_holdout_record.py` | CI fixtures of the exposure controls, on synthetic studies |
| `test_fit_convergence.py` | CI fixtures of Amendment 3's convergence decision: the committed record, and deliberate capped, nonconvergent and missing-status fits |
| `test_result_binding.py` | CI fixtures of the result binding: the committed identities, and mutations of every bound file |
