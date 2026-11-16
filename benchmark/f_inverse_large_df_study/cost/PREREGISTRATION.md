# #359: a cheaper beta inverse — preregistration

Registered 2026-11-16 (Europe/Rome), **before either file of the holdout is
exported and before any production source changes**. The two grid files are
committed with their design and references only; the observed column of both
is blank. The Python model of the candidate has not been run on any row of
the holdout.

## The question

Since #35 a Beta or F inverse takes 6 to 580 times what it took before
(`../production/PRODUCTION_RESULT.md`). The time is in the pair continued
fraction that every Newton step evaluates to 2^-100. Can most calls do
without it, at the accuracy v1.0.0 has?

## The candidate

`selective_model.py`, candidate S, version 3. `FIT_RESULT.md` records how the
fitting sweep shaped it; every point of that sweep is revealed and none is
evidence here.

`PROB_TryBetaInvRegularized(Probability, ComplementProbability, A, B)` keeps
the orientation, the seed, the bracket and the residual on the smaller tail
that #35 gave it. What changes is the forward evaluation inside the loop:

1. **Phase A.** Newton with the binary64 forward function, its continued
   fraction stopped at 4 * 2^-52 where the source stops at 3E-14. The two
   tails are formed as today: for an iterate below the mean,
   `(sa + 1) / (sa + sb + 2)`, the lower tail is the function at the iterate
   and the upper its complement; above the mean the upper tail is the
   function at `1 - u` with the shapes exchanged and the lower its
   complement. An iterate that settles below the mean is compared with its
   two adjacent Doubles and the one with the smallest residual is returned.
2. **Phase B**, only when the iterate phase A settles on is at or above the
   mean, or phase A has not settled in 60 iterations: the loop of v1.0.0
   continues from that iterate, pair fraction and guarded stagnation
   included.

The stop of the binary64 fraction becomes a parameter whose default is
today's 3E-14, so every other caller of the forward function is unchanged.
The regime the source answers with the routine before #35, a lower-tail
target that rounds to one, is not touched.

The same pull request carries two changes to the same module that #400
records and that alter no result: comments that still call #35 open, and the
local variable `Base` renamed.

## The holdout

`build_grid.py` writes the rows, the same in both files; `--check` rebuilds
them and compares.

| Kind | What it holds |
| --- | --- |
| `direct_beta` | Beta at seven pairs of shapes of at least 0.5 and six probabilities from 1E-10 to 1 - 1E-10 |
| `small_shape` | Beta with a shape below 0.5, which the Python model does not cover, at four probabilities from 1E-7 to 1 - 1E-7 |
| `balanced_f` | F at six pairs of moderate degrees of freedom and the same six probabilities |
| `large_f` | F with one degree of freedom far above the other, six pairs, at six probabilities from 1E-9 to 1 - 1E-9 |
| `branch_seam` | Beta at the probability whose quantile is the mean of (3, 40) and of (0.8, 12), and 1 and 4 units in the last place either side |
| `half_seam` | F at the orientation boundary of (6, 10) and (10, 6), and 1 and 4 units in the last place either side |
| `student_t` | Student-t at six degrees of freedom up to 40 and four probabilities above the median |

No row is a point of the fitting sweep, of the production holdout of #35 or
of the cost benchmark, with an F row read as the Beta problem it solves;
`test_cost.py` checks it. The seam probabilities come from 120-digit
arithmetic, not from the model. References are the quantile to 40 digits
from 120-digit arithmetic.

`M_STATS_PROBDIST_BIC.bas` calls the public function each row names and
records the result as the bit pattern of its Double. The inputs travel as
bit patterns. `Export_BetaInverseCostBaseline` fills the baseline file from
the source of v1.0.0 (study export `bic_baseline`); `Export_BetaInverseCost`
fills the other from the candidate (`bic_holdout`).

## The cost measurement

Every full Excel regression ends with `Test_STATS_PROBDIST_GammaCost_CI`,
which times the public inverses and writes one `COST` line each: Beta at four
pairs of shapes, F at four pairs of degrees of freedom, each at p = 0.1 and
0.9, and Student-t at three degrees of freedom. Each timing repeats the call
until 0.5 s has elapsed.

The lines of the run that fills the baseline file are committed as
`cost_baseline.txt`, those of the run that fills the candidate file as
`cost_candidate.txt`. One run each, on the same self-hosted runner. A run
that fails for a reason unrelated to the candidate is repeated whole; no
timing is repeated to obtain a better figure.

In the model the candidate makes no pair evaluation on twelve of the sixteen
Beta and F lines, and at least one on four: Beta (0.7, 25) and (50, 80) at
p = 0.9, F (50, 1E8) at p = 0.9 and F (1E5, 2) at p = 0.1. The split is fixed
here, from the model; if the source takes phase B on one of the twelve, that
line is judged as one of the twelve.

## Acceptance rules

`analyze.py` implements them. The error is relative to the reference.

| Rule | Statement |
| --- | --- |
| R0 | Every cell of both files is a Double or a worksheet error: none is blank and none is a runtime fault |
| R1 | Where both files hold a number, the candidate's error is at most the baseline's plus 3E-14 |
| R2 | Where the baseline is within 1E-12 of the reference, so is the candidate |
| R3 | A row the baseline answers with a number, the candidate answers with a number |
| R4 | Where the baseline is within 1E-13 of the reference, so is the candidate |
| C0 | Both cost files hold every benchmark line, each from a call that returned a number |
| C1 | On the twelve lines without a pair evaluation, the candidate takes at most 5 times what the routine before #35 took |
| C2 | On the four lines with a pair evaluation, the candidate takes at most half of what v1.0.0 takes in the baseline run |

For C1 the time before #35 is carried to the baseline run: the recorded
time before, times the baseline run's time, over the recorded time after,
line by line, from the cost table of `../production/PRODUCTION_RESULT.md`.
The source before #35 is not run again.

R2 and R4 are conditional on the baseline because Addendum 1 of the
production registration taught it: a bound applied to rows where the
accepted source already fails measures that source, not the change. Such
rows are held by R1 and reported.

The candidate is **accepted** only if every rule holds. Also required, as
for every change of a `.bas` file: the full Excel regression passes, and the
strict accuracy gate passes on the main grid and on the consolidated holdout
with no contract changed.

The 3E-14 of R1 is the size of the binary64 noise on the residual below the
mean. I first wrote 1E-14 in the issue; the sweep showed two rows moving by
up to 2.1E-14 and I widened it before this registration, on 2026-11-16.

## Predictions

1. R0 to R4 hold on every row.
2. Most rows change their Double; about as many move closer to the
   reference as farther.
3. Few `student_t` rows change, or none: in #35 none did. I have not
   checked which of these rows reach the routine.
4. C1 holds on the twelve lines, most of them within 3 times the routine
   before #35.
5. C2 holds on the four lines; the gain there is the smallest, between 2 and
   10 times.
6. `small_shape` is where I am least sure: the model does not cover it, and
   it is the kind most likely to break R1.
7. The main grid and the consolidated holdout change in some observations
   and every contract holds.

## What a failure means

- R1, R2 or R4 on rows below the mean: the binary64 residual is not enough
  there, and the selective rule is rejected as it stands. I do not widen a
  bound after the export.
- C1 only: the candidate is correct and not cheap enough. The result is
  recorded and the issue stays open.
- Anything else is recorded as found. A second candidate needs a second
  registration and a holdout of its own.

## What this does not measure

- The legacy regime, a lower-tail target that rounds to one.
- The cost on other hosts, or on points outside the sixteen lines.
- Shapes above 5E9 or probabilities beyond 1E-10 from 0 or 1.

## Addendum 1 — 2026-11-16, with the candidate source, before its export

Written after the baseline run and before the candidate has been run in
Excel. No rule, bound or row changes. It records where the source differs
from the sentence above or from the model, so that the verdict is read
against what was built.

- **The evaluation at 1/2.** The orientation is still decided by one
  evaluation at 1/2, and that evaluation is now binary64, with the tight
  stop. The model has it so; "keeps the orientation" above did not say it.
  Near the boundary either member is close to 1/2 and either choice is
  sound.
- **A first loop that cannot evaluate.** Where the binary64 fraction does not
  converge, the model raises. The source hands the iterate to the second
  loop, as it does for an iterate that has not settled.
- **The stop.** The constant is 8.9E-16, which admits the same last changes
  as 4 * 2^-52 and does not depend on how a literal is parsed.
- **Above the mean in the first loop.** The model forms the factor from
  `1 - (1 - u)`; the source forms it from the iterate. Such an iterate is
  never returned by the first loop.
- **`Base` and one comment stay.** The registration says the pull request
  renames the local `Base`. It does not: the routine is the generated block
  of `../../gamma_route_study/b64_route_vba.py`, which a frozen study binds
  byte for byte, as it binds the header that still calls #26 pending. Both
  are left as they are and #400 records why. The comments that call #35 open
  are corrected.
