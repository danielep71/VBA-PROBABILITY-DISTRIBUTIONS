# Amendment 1 — result of the second holdout (#13)

Recorded 2026-11-13 (Europe/Rome). `AMENDMENT_1.md` was committed and pushed
at `1c20ef9` with blank baselines; this result was exported from that commit.

## Evidence

- Excel run 43499940834 on `1c20ef9`, runner `ITPF545VX1`, Excel 16.0 build
  20430, 64-bit: 6,018 of 6,018 regression assertions; evidence commit
  `56ca24f`. The main grid and the holdout were exported in the same run and
  no observation changed.
- `gamma_holdout2.csv`: 624 rows, raw SHA-256
  `1a2fb1e66ef1ae680d556895768aee0e5c0e3ddd24ce329929a3ddb2568fadf9`.
- `chisquare_holdout2.csv`: 208 rows, raw SHA-256
  `4c0d99c81a5d58a6f701743b5d2ab78df98e264b7f809e0f7af8aa232fba51ee`.
- An earlier run on the same commit, 43495950425, exported the two study
  files alone. They are byte-identical to the committed ones. It could not be
  recorded by itself: the two new exporter modules are sources that the
  main-grid manifest binds, so the main grid had to be exported with them.

The current path measured here is the cumulative before any precision
dispatch, as the amendment requires.

## Verdict

`python3 analyze_amendment1.py gamma_holdout2.csv chisquare_holdout2.csv`:
the construction check passes for all 832 rows, and every claim holds.

| Claim | Gamma | Chi-square |
| --- | --- | --- |
| C1: rounded, 40 bits or fewer, candidate below current | holds in 9 buckets | holds in 9 buckets |
| C2: exact, current not above candidate | holds in 26 bucket-classes | holds in 13 buckets |
| C3: rounded, above 40 bits, current at most 5E-13 | holds; worst 1.79E-13 at 42 bits | holds; worst 2.08E-13 at 42 bits |
| C4: exact with scale other than one, candidate at most 3E-13 | holds; worst 8.68E-14 | not applicable |

Worst relative error, rounded transform:

| Bits | Gamma current | Gamma candidate | Chi-square current | Chi-square candidate |
| ---: | ---: | ---: | ---: | ---: |
| 51 | 3.98E-16 | 1.27E-13 | 4.56E-16 | 2.01E-14 |
| 47 | 5.61E-15 | 5.53E-14 | 6.54E-15 | 5.62E-14 |
| 43 | 8.92E-14 | 1.06E-13 | 1.04E-13 | 3.64E-14 |
| 42 | 1.79E-13 | 7.28E-14 | 2.08E-13 | 9.36E-14 |
| 39 | 1.43E-12 | 9.29E-14 | 1.67E-12 | 2.84E-14 |
| 35 | 2.29E-11 | 1.22E-13 | 2.67E-11 | 1.30E-13 |
| 31 | 3.66E-10 | 8.91E-14 | 4.27E-10 | 1.29E-13 |
| 27 | 5.85E-09 | 5.44E-14 | 6.83E-09 | 1.03E-13 |
| 23 | 9.36E-08 | 8.96E-14 | 1.09E-07 | 9.82E-14 |
| 19 | 1.50E-06 | 1.54E-13 | 1.75E-06 | 4.69E-14 |
| 15 | 2.40E-05 | 7.48E-14 | 2.80E-05 | 8.94E-14 |
| 11 | 3.83E-04 | 8.86E-14 | 4.47E-04 | 4.82E-14 |
| 7 | 6.10E-03 | 8.57E-14 | 7.11E-03 | 7.40E-14 |

With an exact transform the current path stays between 5.8E-17 and 1.1E-16
at every bucket, and the candidate between 2.1E-14 and 1.1E-13.

## What follows

The rule of `AMENDMENT_1.md` may be wired into the two cumulatives, with its
own Excel wave and the zero-movement check. Two numbers from this holdout
describe what the wired cumulative will carry:

- recovered arguments (rounded, 40 bits or fewer): worst candidate error
  1.54E-13;
- arguments left on the direct path (rounded, above 40 bits): worst current
  error 2.08E-13, against the registered bound of 5E-13.

No cutoff was taken from this holdout: 40 bits is the cutoff registered
before it was exported.
