# #331: what every public function returns on bad input

The section "What a function returns on bad input" of
[docs/PUBLIC_API.md](../../docs/PUBLIC_API.md#what-a-function-returns-on-bad-input)
is generated here, from a survey run in Excel.

| File | What it is |
| --- | --- |
| `build_grid.py` | Derives the survey from `docs/PUBLIC_API.txt`: one valid call of every public function, and that call with each argument replaced in turn by each class of input. Also generates the dispatch of the exporter and the battery of the VBA test module |
| `error_behavior_grid.csv` | The survey and what Excel returned for each call |
| `M_STATS_PROBDIST_ERB.bas` | Makes each call and records what came back: a number, a worksheet error, or a raised VBA error (study export `erb_survey`) |
| `check_rules.py` | States the rules, checks them on every call, checks each listed difference between functions, and writes the section of the page |
| `test_survey.py` | The survey, the exporter's dispatch and the battery against the public surface; each rule on a row made to break it |

## Two checks, and what each covers

- **The survey** covers every class of input, the text and the empty cell
  included, but only when it is exported: it is a study export, not part of
  every regression. `check_rules.py --check` runs in the evidence-tool suite
  on the committed export.
- **The battery**, `Test_Core_ApiErrorContract` in the VBA test module, runs
  at every regression. It checks, for every public function, that a valid
  call returns no error and that each argument set to not-a-number,
  +infinity and -infinity returns `#NUM!`. It does not check the text, the
  empty cell or the domain of each parameter: those are in the survey and in
  the family tests.

A new public function without a row in either fails `test_survey.py`.

## What the first survey found

Run 43962906694, before any change to the source:

- a standard deviation that is not a number returned `#VALUE!` in the Normal
  and Lognormal families, in 17 function and argument pairs. Fixed in the
  same pull request; run 43967812085 shows all 17 as `#NUM!` and nothing
  else changed;
- `K_STATS_NormalStandard_InverseCumulativeFast` validates nothing. Kept: it
  is its documented contract;
- six places where functions answer the same input differently. Three are
  kept with a reason and three are not settled (#418).

The first attempt, run 43960861784, exported nothing: it reached the
functions through `Application.Run`, an error raised inside a function
called that way does not reach the caller's handler, and Excel stopped on a
dialog. The exporter calls every function directly since.

## What this does not show

- Any value between the classes: -1 stands for every negative number, 1E308
  for every large one.
- More than one argument wrong at a time.
- What the `Status` argument holds after a failure. The survey records the
  returned value only.
- Any Excel build other than 16.0 build 20710, 64-bit.
