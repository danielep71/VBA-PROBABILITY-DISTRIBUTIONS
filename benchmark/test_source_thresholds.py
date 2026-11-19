#!/usr/bin/env python3
"""Negative controls for check_source_thresholds.py."""
from __future__ import annotations

import subprocess
import sys
import tempfile
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
import check_source_thresholds as C  # noqa: E402

failures: list[str] = []


def expect_match(text: str, label: str) -> None:
    if not any(pattern.search(text) for _, pattern in C.CLAIMS):
        failures.append(label + " was not detected")


def expect_clean(text: str, label: str) -> None:
    if any(pattern.search(text) for _, pattern in C.CLAIMS):
        failures.append(label + " was falsely detected")


expect_match("' CDF <= 5E-15 relative", "canonical threshold claim")
expect_match("' density accurate to 2E-14", "accuracy-to claim")
expect_match("' df above 1E8 are rejected", "cap-above claim")
expect_match("' validated to roughly 1E9", "validated-to claim")
expect_match("' F_Density is unrestricted", "unrestricted-density claim")
expect_clean("Private Const PROB_F_MAX_DF As Double = 1E10", "authoritative constant text")
expect_clean("' scale parameter is unrestricted", "non-envelope unrestricted argument")

# Exercise the real repository too: the positive path must currently be clean.
proc = subprocess.run([sys.executable, "check_source_thresholds.py"], cwd=HERE,
                      capture_output=True, text=True)
if proc.returncode != 0:
    failures.append("live repository check did not pass: " + (proc.stdout + proc.stderr).strip())

# The scan reads the modules where the repository keeps them: a claim in a
# module of src/modules/ is found, a tree with no module there fails closed
# instead of passing on an empty set (#438), and the legacy flat src/ is not
# scanned.
with tempfile.TemporaryDirectory() as tmp:
    root = Path(tmp)
    (root / "src" / "modules").mkdir(parents=True)
    module = root / "src" / "modules" / "M_STATS_PROBDIST_FIXTURE.bas"
    module.write_text("Attribute VB_Name = \"M_STATS_PROBDIST_FIXTURE\"\n' CDF <= 5E-15 relative\n",
                      encoding="ascii")
    code = C.main(["--root", str(root)])
    if code != 1:
        failures.append("a threshold claim in src/modules/ must fail the scan")
    module.write_text("Attribute VB_Name = \"M_STATS_PROBDIST_FIXTURE\"\n' see accuracy_contracts.csv\n",
                      encoding="ascii")
    if C.main(["--root", str(root)]) != 0:
        failures.append("a clean module in src/modules/ must pass the scan")
    module.unlink()
    (root / "src" / "M_STATS_PROBDIST_FLAT.bas").write_text("' CDF <= 5E-15 relative\n", encoding="ascii")
    if C.main(["--root", str(root)]) != 1:
        failures.append("a scan that finds no module under src/modules/ must fail, not pass")

if failures:
    print("FAIL: source-threshold checker fixtures")
    for failure in failures:
        print("  - " + failure)
    raise SystemExit(1)
print("PASS: source-threshold checker fixtures (threshold/accuracy/cap/envelope claims fail; authoritative and unrelated text pass; a claim in src/modules/ fails and an empty scan fails)")
