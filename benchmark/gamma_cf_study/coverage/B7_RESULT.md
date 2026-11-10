# B7 rejected: isolated compensation is not sufficient

The preregistered differential screen in commit 02303d6 evaluated 89 revealed
fitting points and 200 new points. The production VBA was not changed.

B7 removes the original A=.5 seam failure in the source-derived model, but fails
its frozen decision rules:

- One new survival-contract failure at A=.34,
  X=`0x1.570a3d70a3d6bp+0`: relative error moves from 4.637759129595115E-15
  (B6, passing 5E-15) to 8.262480325560744E-15 (B7, failing).
- Three new monotonicity failures: CDF at A=.021 and CDF/SF at A=.88, each
  between the predecessor of binary64(1+A) and binary64(1+A).
- Independent quadrature does not corroborate two fitting references at the
  selected precision: A=100, X=202 and 808. These are recorded as oracle
  failures, never scored as PASS. All 200 new-point references satisfy the
  two-precision stability and quadrature checks. The numerical candidate
  already fails independently of those two oracle failures.

`screen_b7.py` regenerates `b7_screen.json`, retaining exact binary64 inputs,
B6/B7 outputs, references, per-surface errors, failure identities, runtime and
LF-canonical source/script/protocol hashes. A formatting-only rerun adds these
input bindings; it does not tune the candidate. The 200-point set is now revealed.
An initial development run was stopped before independent evaluation to correct
the model's predicate to A<=1, matching the frozen protocol; the A>1 path is
unchanged in the recorded runs.

## Counterexample decomposition

`seam_diagnosis.py` reproduces the source arithmetic with explicit binary64
stores and the existing extended-intermediate model. At A=.5,
X=predecessor(1.5), the lower P model is 0.91673548333644994, about 3.2949 ulps
above the high-precision reference. Replacing only the prefactor with its
correctly rounded reference leaves this result unchanged. Replacing only the
sum, or compensating the sum, gives 0.91673548333644961 (about .2949 ulp).
At X=1.5 the B6 CF model gives Q=0.08326451666355036 and P=0.91673548333644961.
Thus the model locates the original junction mismatch in lower accumulation.
It does not establish that compensation alone is globally safe: the screen
above falsifies that proposed repair.

Excel run 42469185215 on synthetic merge ce9a483 executed 1945 assertions:
1905 PASS / 40 FAIL, unchanged from the preceding B6 result. Its extra formatted
output confirms the survival jump at displayed precision: left .08326451666355010,
right .08326451666355040. VBA Format$ pads a roughly 15-significant-digit
conversion even when 17 decimal places are requested; the printed CDF values
both become .91673548333645000. Therefore this log does NOT establish bitwise
model fidelity or an exact Excel jump magnitude. Do not treat the additional
zeroes as additional significant digits. Exact-output extraction remains needed
before any production repair.

## Disposition

Reject B7. Keep B6 integration blocked and both #26/#24 open. Do not retune
against these now-revealed points. The unchanged lower-region Q=1-P construction
still amplifies low-order P movement; any next design must address that coupled
arithmetic and use a new preregistered independent set. No source fix, fresh
export, threshold waiver or manifest rebinding is claimed by this study.
