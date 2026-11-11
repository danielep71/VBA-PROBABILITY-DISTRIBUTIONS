# DD kernel profiling (#307)

Measure the validated production kernel before optimizing it. No routing or
numerical expression is changed in the repository. The untouched #307 holdout
is neither imported nor evaluated.

`Gamma DD Kernel Profile` runs in Excel on the existing Windows x64 runner.
`prepare.py` generates two temporary instrumented copies of SPECIALFUNCS;
removing their entry/exit statements must recover every original source byte.
`Run-KernelProfile.ps1` imports the production modules into a disposable workbook.
It measures the original code, coarse mathematical scopes, arithmetic scopes,
and finally the restored original code. Each pass uses the eight previously
revealed #298 points for P and Q, plus inverse (shape 2, p=.5) and
(shape 100, p=.9), scale 1. Each of three samples has a separate warmup;
forward samples contain four calls, inverse samples one. Every evaluated result
must match the initial uninstrumented result bit for bit, including warmup
through the returned sample result. Scope stack imbalance fails the run.

QPC wall-clock times include ordinary call overhead and a result-bit check.
The baseline/restored samples measure actual uninstrumented performance.
Instrumented inclusive/exclusive times describe nested scopes, not an additive
list of independent costs. Exclusive time includes uninstrumented descendants
and profiler overhead; it is not pure arithmetic time. Entry/exit overhead is
calibrated with 100,000 empty pairs and reported without speculative subtraction.
Counts are exact for completed scopes. Profiling errors invalidate the pass;
there is no silently ignored unwind or assumed fallback attribution.

`profile.json` records the candidate commit, Excel build, all raw samples and
bits. `instrumentation.json` maps scope IDs and hashes LF-normalized production
source. The generated source is retained in the run artifact. These measurements
are diagnostics, not new regression certification or holdout acceptance.
Production source and historical certification manifests remain unchanged.

Local preparation: `python benchmark/kernel_profile/prepare.py`.
Actual measurement requires Windows, Excel 365 64-bit and pre-enabled VBProject
access: `powershell -File benchmark/kernel_profile/Run-KernelProfile.ps1`.
