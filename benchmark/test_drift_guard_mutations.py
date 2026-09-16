#!/usr/bin/env python3
"""Mutation controls for drift detectors that otherwise only have a green path.

A non-zero child process is not sufficient proof: setup/import errors could make
these tests falsely green. Each mutation therefore requires the detector's
specific diagnostic as well as a failing exit code.
"""
from __future__ import annotations

import json
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

HERE = Path(__file__).resolve().parent
failures: list[str] = []


def run(script: Path, cwd: Path) -> subprocess.CompletedProcess[str]:
    return subprocess.run([sys.executable, script.name], cwd=cwd,
                          capture_output=True, text=True)


def output(proc: subprocess.CompletedProcess[str]) -> str:
    return (proc.stdout or "") + (proc.stderr or "")


# 1. Incomplete-gamma parity: move the production seam. The parity detector
# must reject a production/study dispatch drift instead of merely proving today's
# matching implementations happen to agree.
with tempfile.TemporaryDirectory() as td:
    root = Path(td)
    shutil.copy2(HERE / "test_igamma_parity.py", root / "test_igamma_parity.py")
    shutil.copy2(HERE / "_igamma.py", root / "_igamma.py")
    shutil.copytree(HERE / "chisq_reference_study", root / "chisq_reference_study")
    igamma = root / "_igamma.py"
    text = igamma.read_text(encoding="utf-8")
    needle = "return mp.mpf(a) + 1"
    if needle not in text:
        failures.append("parity mutation seam anchor not found")
    else:
        igamma.write_text(text.replace(needle, "return mp.mpf(a) + 2", 1), encoding="utf-8")
        proc = run(root / "test_igamma_parity.py", root)
        out = output(proc)
        if proc.returncode == 0:
            failures.append("incomplete-gamma parity accepted an intentionally moved production seam")
        elif "seam is a + 1" not in out or "FAIL: incomplete-gamma parity" not in out:
            failures.append("incomplete-gamma parity failed for an unexpected reason:\n" + out[-1200:])

# 2. Student-t coefficients: alter one exact rational in the stored authority.
# Fresh symbolic derivation must reject it.
with tempfile.TemporaryDirectory() as td:
    root = Path(td)
    study = root / "student_t_large_df_study"
    shutil.copytree(HERE / "student_t_large_df_study", study)
    coeff_path = study / "coefficients.json"
    data = json.loads(coeff_path.read_text(encoding="utf-8"))
    terms = data["g"]["1"]["terms"]
    if not terms:
        failures.append("coefficient mutation fixture found no g_1 terms")
    else:
        terms[0][1] = int(terms[0][1]) + 1
        coeff_path.write_text(json.dumps(data, indent=2) + "\n", encoding="utf-8")
        proc = run(study / "test_coefficients.py", study)
        out = output(proc)
        if proc.returncode == 0:
            failures.append("Student-t coefficient fixture accepted an intentionally corrupted rational")
        elif "stored g_1 does not match a fresh derivation" not in out or "FAIL: Student-t large-df coefficients" not in out:
            failures.append("Student-t coefficient fixture failed for an unexpected reason:\n" + out[-1200:])

if failures:
    print("FAIL: drift-guard mutation controls")
    for failure in failures:
        print("  - " + failure)
    raise SystemExit(1)
print("PASS: drift-guard mutation controls (incomplete-gamma seam drift and Student-t coefficient corruption are rejected for the intended detector reasons)")
