# #24 direct-Q preparation and exact #26 seam observation

Status: **diagnosis only; no new VBA candidate accepted or implemented**.
This is retrospective exploration on already revealed points. It is not a
preregistration or a new independent holdout. B7 remains rejected.

## Exact Excel observation

Diagnostic commit `732e745814ba70c9f1680710f7a177773209a63b`, PR merge checkout
`9fdc1542f604cb2d6e74e1aec1d8a9889dfee107`,
[Excel run 42470083523](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/actions/runs/42470083523),
job 128698575955, Excel 16.0 build 20430, 64-bit.
The result remains **1945 total / 1905 PASS / 40 FAIL**. No export took place.
The complete failure excerpt is `exact_seam_run_42470083523.txt`.

The test generator copies the established accuracy export's two-part formatter,
with private test-only names. The assertions, coordinates, production source,
thresholds and expected assertion count are unchanged. The decimal high part is
parsed to binary64 in VBA before subtracting it; therefore this bit comparison
reconstructs `float(hi) + float(lo)`. It does not sum the decimal strings at
arbitrary precision. No existing evaluator/parser is changed here.

At A=.5, scale=1, x_left=predecessor(1.5), x_right=1.5:

| Surface | Left binary64 | Right binary64 | Wrong-direction jump |
| --- | --- | --- | --- |
| CDF | `0x1.d55e5a70068e7p-1` | `0x1.d55e5a70068e4p-1` | -3 CDF ulps |
| SF | `0x1.550d2c7fcb8c8p-4` | `0x1.550d2c7fcb8dep-4` | +22 SF ulps |

All four values match `seam_diagnosis.public` bit for bit. This establishes model
fidelity at this counterexample only, not throughout the domain. It supersedes
the padded Format$ diagnostics as evidence of the jump magnitude.

## Direct-Q identity and rounding diagnosis

For positive a and x, termwise integration of exp(-t) gives

```
S = sum(n>=1) (-x)^n / (n! * (a+n))
logP = a*log(x) - logGamma(1+a) + log1p(a*S)
P = exp(logP)
Q = -expm1(logP)
```

The explored scope is only `0<a<=1, x<a+1`; the upper CF and boundary stay fixed.
Unlike log of a rounded P, the expression can retain a small negative logP when
P rounds to one. All 269 A<=1 rows from the already revealed B7 record are used;
the old `independent` field is retained solely as a historical row identifier.
A>1 is excluded, not silently passed. The stored 90/120-digit references are
reused; this is not a new oracle certification.

| Diagnostic arm | Accuracy failures (P 3E-15 / Q 5E-15) | All adjacent monotonicity failures | New relative to B6 |
| --- | ---: | ---: | ---: |
| Binary64 recurrence, math.fsum, Python log1p/expm1 | 0 | 26 | 18 |
| 100-digit series, rounded argument and library transcendentals | 0 | 19 | 17 |
| 100-digit arithmetic except existing rounded logGamma1p | 0 | 0 | 0 |
| Ideal lower P/Q from high precision | 0 | 0 | 0 |

The binary64 arm is **rejected**, despite meeting pointwise error thresholds and
repairing the original A=.5 junction in the model. For example at A=.34,
X=`0x1.570a3d70a3d6bp+0` -> `0x1.570a3d70a3d6cp+0`, its Q increases.
Python math.fsum/log1p/expm1 make this an optimistic feasibility probe, not a VBA
emulation. The high-precision arms are sensitivity experiments, not production
solutions and not proof that a particular finite-precision implementation will
work. Exact summation alone is insufficient; loss in argument/logP formation
and transcendentals must be addressed together.

Reproduce with `python benchmark/gamma_cf_study/coverage/direct_q_diagnosis.py`.
It writes the source-bound JSON and intentionally exits 1 for the rejected
binary64 arm. Unexpected cap exhaustion, missing references, invalid logP or
failure of the four exact Excel comparisons raises instead of reporting success.

Next: develop a bounded precision strategy for the complete lower P/Q path on
revealed diagnostics, then freeze a fresh disjoint validation set before scoring
an implementable candidate. Do not move the seam, relax contracts or tune against
any purportedly unrevealed B7 points. #241/#242 remain draft and blocked; neither
#24 nor #26 is ready to close.
