#!/usr/bin/env python3
"""Negative controls for check_public_api.py."""
from __future__ import annotations

import os
import shutil
import sys
import tempfile
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
import check_public_api as A

FAILS: list[str] = []


def check(condition: bool, message: str) -> None:
    if not condition:
        FAILS.append(message)


def expect_fail(root: Path, manifest: Path, label: str) -> None:
    rc = A.main(["--root", str(root), "--manifest", str(manifest)])
    check(rc != 0, label + " must fail")


def expect_pass(root: Path, manifest: Path, label: str) -> None:
    rc = A.main(["--root", str(root), "--manifest", str(manifest)])
    check(rc == 0, label + " must pass")


def module_text(decl: str, body: str = "    K_STATS_Test = X\n") -> str:
    return (
        'Attribute VB_Name = "M"\n'
        "Option Explicit\n\n"
        + decl
        + "\n"
        + body
        + "End Function\n"
    )


def make_repo() -> tuple[Path, Path]:
    root = Path(tempfile.mkdtemp())
    for rel in A.MODULES:
        path = root / rel
        path.parent.mkdir(parents=True, exist_ok=True)
        if rel.endswith("CORE.bas"):
            path.write_text(
                module_text(
                    "Public Function K_STATS_Test( _\n"
                    "    ByVal X As Double, _\n"
                    "    Optional ByRef Status As String = \"\") _\n"
                    "    As Variant"
                ),
                encoding="utf-8",
            )
        else:
            path.write_text('Attribute VB_Name = "Empty"\nOption Explicit\n', encoding="utf-8")
    manifest = root / "docs" / "PUBLIC_API.txt"
    manifest.parent.mkdir(parents=True, exist_ok=True)
    manifest.write_text(A.render(A.collect(root)), encoding="utf-8")
    return root, manifest


root, manifest = make_repo()
try:
    expect_pass(root, manifest, "baseline")

    core = root / A.MODULES[0]
    baseline = core.read_text(encoding="utf-8")

    # Function-body and formatting changes are not API changes.
    core.write_text(baseline.replace("K_STATS_Test = X", "K_STATS_Test = X + 0#"), encoding="utf-8")
    expect_pass(root, manifest, "implementation-only edit")

    wrapped = baseline.replace(
        "Public Function K_STATS_Test( _\n    ByVal X As Double, _",
        "Public    Function   K_STATS_Test(    _\n       ByVal   X   As   Double,    _",
    )
    core.write_text(wrapped, encoding="utf-8")
    expect_pass(root, manifest, "whitespace/line wrapping edit")

    mutations = {
        "rename": baseline.replace("K_STATS_Test", "K_STATS_Renamed"),
        "parameter type": baseline.replace("X As Double", "X As Long"),
        "ByVal to ByRef": baseline.replace("ByVal X", "ByRef X"),
        "optional removed": baseline.replace("Optional ByRef Status", "ByRef Status"),
        "default changed": baseline.replace('Status As String = ""', 'Status As String = "ok"'),
        "return type": baseline.replace("As Variant\n", "As Double\n", 1),
        "parameter order": baseline.replace(
            "ByVal X As Double, _\n    Optional ByRef Status As String = \"\"",
            "Optional ByRef Status As String = \"\", _\n    ByVal X As Double",
        ),
    }
    for label, text in mutations.items():
        core.write_text(text, encoding="utf-8")
        expect_fail(root, manifest, label)

    # Additive public API is drift until deliberately accepted in the manifest.
    core.write_text(
        baseline + "\nPublic Function K_STATS_Added(ByVal X As Double) As Double\n"
        "    K_STATS_Added = X\nEnd Function\n",
        encoding="utf-8",
    )
    expect_fail(root, manifest, "added K_STATS function")

    # Internal PROB namespace may not leak into the production public surface.
    core.write_text(
        baseline + "\nPublic Function PROB_Leaked(ByVal X As Double) As Double\n"
        "    PROB_Leaked = X\nEnd Function\n",
        encoding="utf-8",
    )
    expect_fail(root, manifest, "Public PROB helper exposure")

    # Private PROB helpers remain allowed and irrelevant to the API manifest.
    core.write_text(
        baseline + "\nPrivate Function PROB_Internal(ByVal X As Double) As Double\n"
        "    PROB_Internal = X\nEnd Function\n",
        encoding="utf-8",
    )
    expect_pass(root, manifest, "Private PROB helper")

    # Public exporters outside src/ are intentionally not part of the product API.
    exporter = root / "benchmark" / "SomeExporter.bas"
    exporter.write_text(
        'Attribute VB_Name = "SomeExporter"\nPublic Sub Export_Something()\nEnd Sub\n',
        encoding="utf-8",
    )
    expect_pass(root, manifest, "benchmark Public Sub outside production API")
finally:
    shutil.rmtree(root, ignore_errors=True)

if FAILS:
    print("FAIL: public API checker fixtures")
    for failure in FAILS:
        print("  - " + failure)
    raise SystemExit(1)
print("PASS: public API checker fixtures (name, type, ByVal/ByRef, Optional/default, "
      "parameter order, return type, additive API and Public PROB exposure fail; "
      "body/format/private-helper/benchmark-exporter edits pass)")
