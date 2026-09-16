#!/usr/bin/env python3
"""Mutation controls for drift detectors that otherwise only have a green path."""
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
        if proc.returncode == 0:
            failures.append("incomplete-gamma parity accepted an intentionally moved production seam")

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
        if proc.returncode == 0:
            failures.append("Student-t coefficient fixture accepted an intentionally corrupted rational")

if failures:
    print("FAIL: drift-guard mutation controls")
    for failure in failures:
        print("  - " + failure)
    raise SystemExit(1)
print("PASS: drift-guard mutation controls (incomplete-gamma seam drift and Student-t coefficient corruption are rejected)")
