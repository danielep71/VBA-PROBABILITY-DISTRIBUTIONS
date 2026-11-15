# 10. Numerical limitations

All figures in this chapter come from [`accuracy_guarantees.csv`](../../../benchmark/accuracy_guarantees.csv) and [`numerical_limitations.csv`](../../../benchmark/numerical_limitations.csv) in the repository. Unless stated otherwise, a figure is a relative error: 1E-13 means roughly 13 correct significant digits.

## 10.1 Two tiers: guarantee and measured target

Each tested combination of function and parameter region is an accuracy contract. Every contract carries two numbers.

| Tier | Meaning |
|---|---|
| Guarantee | What the release promises; a miss blocks the release |
| Measured target | What the code achieves today; reported as evidence only |

The guarantee is derived from the measured target by a fixed rule.

| Kind of function | Guarantee |
|---|---|
| Density, PMF, cumulative, survival, moments | larger of target and 1E-13 |
| Inverse functions | larger of target and 1E-12 |
| Beta cumulative and survival, balanced shapes | larger of target and 2E-13 |
| Absolute and residual measures | equal to the target |

The guarantee is never stricter than the measured target. Where the target is looser than the floor, the guarantee equals the target, which is why some guarantees below are looser than 1E-13.

Examples of **measured** targets in the main region, given as evidence and not as guarantees: 5E-15 for the Standard Normal cumulative, 3E-15 for the Gamma cumulative and 2E-15 for the Geometric PMF.

## 10.2 What a guarantee means

- It applies to the tested functions, parameter regions and points. It is not a bound for every possible input.
- Outside the tested regions, a function returns `#NUM!` or may be less accurate.
- It is a release threshold checked against independently computed reference values, not an external certification.
- Whether a release has passed its checks is stated in the README assurance block and the release notes.

## 10.3 Guarantees by family

The table gives the guarantee in the main tested region. Where cumulative and survival differ, the looser figure is shown.

| Family | Density or PMF | Cumulative, survival | Inverse |
|---|---|---|---|
| Standard Normal, Normal, Lognormal | 1E-13 | 1E-13 | 1E-12 |
| Student t | 1E-13 | 1.3E-12 | 3.0E-12 |
| Chi-square | 1E-13 | 2.6E-10 | 4.7E-12 |
| F | 1E-13 | 1.1E-10 | 2E-10 |
| Gamma | 1E-13 | 1E-13 | 1E-12 |
| Beta, balanced shapes | 1E-13 | 2E-13 | 1E-12 |
| Beta, unbalanced shapes | 4E-12 | 2E-10 | 1E-10 |
| Exponential, Weibull, Uniform | 1E-13 | 1E-13 | 1E-12 |
| Binomial | 1E-12 | 5E-8 | exact integer |
| Poisson | 1E-13 | 2E-13 | exact integer |
| Geometric, Discrete Uniform | 1E-13 | 1E-13 | exact integer |
| Negative Binomial | 2E-12 | 5E-11 | exact integer |
| Hypergeometric | 2E-13 | 1E-13 | exact integer |

Notes:

- "Balanced" means the smaller Beta shape is at least one tenth of the larger.
- All moment functions, the z-score, the interval probabilities and the Lognormal parameter conversion: 1E-13.
- `LogPMF`: 1E-13; 2E-13 for Negative Binomial and Hypergeometric.
- `K_STATS_NormalStandard_InverseCumulativeFast`: 5E-9.
- Discrete inverses are held to an absolute error of 1E-9, which identifies the integer.
- Main regions: degrees of freedom below 1E5 for the Student t inverse, below 1E6 for the Chi-square inverse, and up to 1E5 for F. Normal-family survival guarantees are stated for standardized values up to 6 (Standard Normal), 4 (Normal) and 3.5 (Lognormal).
- The contracts record Binomial trials up to 1E7, a Poisson mean up to 1E6, a Negative Binomial success count up to 5000 and a Hypergeometric population up to 1E5. Larger accepted arguments (section 6.5) are outside the measured evidence.

Regions with a different guarantee:

- **Normal-family inverses, tail probability near 1E-12:** 1E-10 for the Standard Normal, 2E-10 for the Normal inverse survival, 1E-9 for the Lognormal inverse survival.
- **Student t, degrees of freedom 1E5 to 1E8:** 1E-13 for cumulative and survival up to an absolute $t$ of 8; 1E-12 for the inverse. The density is held to 5E-12 for degrees of freedom from 1E2 to 1E20.
- **Chi-square, degrees of freedom 1E6 to 1E8:** 2E-12 for cumulative and survival, 1E-12 for the inverse.
- **F:** 2E-11 for cumulative and survival with both degrees of freedom from 1E6 to 1E10; 1E-12 for the inverse with a degree of freedom above 1E5.
- **Gamma cumulative and survival:** 2E-13 for shapes from 1E2 to 1E6, up to 2E-12 for shapes up to 1E10.
- **Beta cumulative and survival:** 1E-12 with both shapes from 1E2 to 1E6, 5E-10 with both shapes up to 1E12.
- **Beta and F density with both shapes or degrees of freedom from 1E2 to 1E20:** 1E-5 and 5E-5 near the mode. That is about five significant digits. Validate such a density independently if you need more.

The large-parameter rows for cumulative and survival functions were measured near the centre of the distribution. The guarantee file lists further contracts at the edge of the Double range.

## 10.4 Recorded limitations

The project keeps a register of numerical limitations. Every entry for v1.0.0 follows. "Mitigated" means the library refuses the affected inputs with `#NUM!`. "Characterized" means the behaviour is measured and documented, not removed. "Resolved" means it has been corrected and is kept as history.

**F with very large degrees of freedom (mitigated).** F cumulative, survival and inverse. Above 1E10 degrees of freedom the calculation loses accuracy without noticing; at 1E10 it was measured at about 5.1E-10. Any degree of freedom above 1E10 returns `#NUM!`. The F density has its own limit of 1E20.

**Student t with very large degrees of freedom (mitigated).** Student t cumulative, survival and inverse. Above 1E8 degrees of freedom the tail calculation loses accuracy while still appearing to converge: the register records 1.2E-9 at 1E8, growing roughly tenfold per factor of ten. Degrees of freedom above 1E8 return `#NUM!`. The density is not restricted. Beyond about 1E15 the Student t equals the Normal to Double precision.

**Chi-square with very large degrees of freedom (mitigated).** Chi-square cumulative, survival and inverse. The calculation is not validated above 1E8 degrees of freedom; 3.1E-11 was measured at 1E8. Larger values return `#NUM!`. The density is covered up to 1E20.

**Student t far tail at large degrees of freedom (characterized).** Student t cumulative, survival and inverse. For degrees of freedom from 1E5 to 1E8 and an absolute $t$ above 8, no contract applies and the error grows with $t$. Measured in Excel: 3.4E-15 just above 8, at most 1.84E-14 up to 14, and 9.7E-14 at 20 and 30, where the tail is below 1E-88.

**Student t just below 1E5 degrees of freedom (characterized).** Student t cumulative and survival. Measured in Excel at the largest value below 1E5: 3.16E-12 at $t = 2$, 5.97E-13 at $t = 4$ and 4.19E-13 at $t = 8$, outside the 1E-13 forward guarantee. How far below 1E5 the effect extends is not measured.

**Student t steps at internal boundaries (characterized).** Student t cumulative and survival. The method changes at 1E5 degrees of freedom and at an absolute $t$ of 8. Across 1E5 degrees of freedom with $t$ fixed, the probability can move in the wrong direction by up to 3.16E-12 (at $t = 2$). Across $t = 8$ the results stay in order. Survival plus cumulative equals 1, and survival at $t$ equals cumulative at $-t$, at these boundaries.

**Beta inverse with two very small shapes (characterized).** When both shapes are roughly below 1E-3, almost all probability sits at 0 and 1. For some probabilities the true quantile is closer to 0 or 1 than any Double can express, and no meaningful result exists. This is a limit of the number format, not a solver failure. Accuracy statements apply only where the quantile can be represented.

**F inverse with two very small degrees of freedom (characterized).** The F distribution is computed through the Beta distribution, so the same effect appears: for some probabilities the true quantile leaves the range of a Double.

**Beta and F inverse for a probability extremely close to 0 or 1 (characterized).** For a probability within $2^{-53}$ (about 1.1E-16) of 0 or 1 whose quantile lies beyond one half, the smaller of $x$ and $1 - x$ carries an error of about $2^{-52} / (1 - x)$. This regime is not measured in Excel across its range.

**Normal upper tail (resolved).** Standard Normal, Normal and Lognormal survival. An earlier method lost relative accuracy beyond standardized values of about 3. It has been replaced; Excel now measures at most 2.56E-16 from -2 to 6. No action is required.

**Internal log-gamma helper (characterized).** An internal helper, not a worksheet function, cannot return an accurate result for arguments below about 3.85E-308, because the result is too small to be stored precisely.

## 10.5 Speed in particular regions

Timing depends on the machine. These figures come from the project's measured Excel run.

- Gamma, Chi-square and Poisson probabilities: about 0.06 to 0.23 ms per call for shapes from 0.0001 to 10,000.
- Slower cases: an upper tail evaluated below `Shape + 1` when the lower tail exceeds 0.9, and results below 1E-290, at 13 to 400 ms per call; inverses for shapes below 0.01 whose quantile is not extremely small, at up to a second. Shapes above 1E15 are also slow.
- Gamma or Chi-square inverse: 0.05 to 2 ms up to shape 100, about 20 ms at shape 10,000.
- Beta or F inverse: 0.6 to 28 ms per call; balanced cases are the slowest.
