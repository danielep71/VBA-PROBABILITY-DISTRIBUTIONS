# Private scaled-product VBA result

**Primitive port PASS; full release regression FAIL.**

[Excel run 42476278844](https://github.com/danielep71/VBA-PROBABILITY-DISTRIBUTIONS/actions/runs/42476278844)
was manually dispatched on exact candidate
`86a7e47be621b32350fb343aa8afaa212eddddb9`, not a synthetic merge.
Excel 16.0 build 20430, 64-bit, finished 2026-11-08T11:34:44Z.
The execution-backed compile, import and cleanup stages passed.
The regression result is **7,104 total / 7,063 PASS / 41 FAIL**.

All **4,657 new private arithmetic assertions passed**, covering every binary64
input exponent, 216 rational-reference product pairs, direct halfway rounding,
signed zero, invalid inputs and bit-layout controls. The original 500 private
Precision v2 fidelity assertions also remain passing. This establishes the
new primitive's observed VBA behavior on the frozen fixtures; it is not a
full-domain Gamma acceptance claim.

The 41 failure details are identical, in order and numeric text, to the retained
private-v2 run 42473052390. None starts with the new `SP ` assertion prefix.
All pre-existing VBA regression text and production modules are byte-equivalent
(after the repository's CRLF normalization) to the pre-port candidate. The new
implementation is private to the test module. Public P/Q routes are unchanged.

## Evidence retained and verified

`excel_scaled_product_artifact.json` retains the original ZIP member bytes as
base64, including BOM and CRLF, following the prior private-v2 artifact format.

- Artifact ID: `13615671268`.
- Original ZIP SHA-256:
  `6ffb48e054c0221058d46c659591a11ef8a90394df6fac2e5079a46e070ea5b8`.
- Original test log SHA-256:
  `286d6c0f3cf8621cb63e161552bc259fdb774925107b42ab9fdc51b0c9b16b83`.
- Every source digest in the record matches its Git blob at the exact candidate.
- Original member digests, log binding, seven-source inventory, compile/cleanup
  status, counts and all 41 previous failure details are checked by the fifth
  fixture method in `test_scaled_vba_port.py`.
- No grids were exported. No live certification record or observation manifest
  was rebound to this candidate.

Reproduce the primitive and retained-evidence checks:

```bash
python benchmark/gamma_cf_study/coverage/test_scaled_product.py
python benchmark/gamma_cf_study/coverage/test_scaled_vba_port.py
```

Both pass locally (seven and five fixture methods respectively). All 18 strict
static checks passed on the port tree, including 397 repository tests with
four declared skips. The new retained-evidence fixture was subsequently run
against the downloaded artifact and passed. These local results are distinct
from hosted integration checks.

The full local numerical-tooling runner executed all 68 commands: 66 passed;
the coefficient proof and its mutation control initially failed because SymPy
was missing from this workspace. After installing the repository's hash-locked
`benchmark/requirements-ci.txt` (SymPy 1.14.0, mpmath 1.3.0), both failed commands
were rerun and passed. Thus all 68 checks have passing local evidence across
the full run and targeted reruns; no failed check was waived. This does not
make the separate live Accuracy Gate or the complete Excel regression green.

## Remaining work

The next numerical step is general exponent-safe DD operations and complete
lower/upper Gamma error control, followed by validated public dispatch and
caller-seam repair, including A>1. The #13/#14 positive-ratio and inverse
composition work remains required. Freeze a fresh independent probability set
only after that complete candidate is fixed; then run exact-source Excel,
refresh exports and obtain current integration checks/review. None of the
revealed primitive or probability sets can be relabelled a new holdout.

#24/#26 remain open. #241/#242 remain draft and blocked by the 41 retained
public failures. The successful primitive port does not authorize a merge or
release and changes no accuracy threshold.
