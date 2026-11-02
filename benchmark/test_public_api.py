#!/usr/bin/env python3
"""Negative controls for check_public_api.py."""
from __future__ import annotations

import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
import check_public_api as A

FAILS: list[str] = []
FIXTURE_MODULES = (
    "src/M_STATS_PROBDIST_CORE.bas",
    "src/M_STATS_PROBDIST_SPECIALFUNCS.bas",
    "src/M_STATS_PROBDIST_NORMALFAMILY.bas",
    "src/M_STATS_PROBDIST_TFAMILY.bas",
    "src/M_STATS_PROBDIST_CONTINUOUS.bas",
    "src/M_STATS_PROBDIST_DISCRETE.bas",
)


def check(condition: bool, message: str) -> None:
    if not condition:
        FAILS.append(message)


def expect_fail(root: Path, manifest: Path, label: str) -> None:
    rc = A.main(["--root", str(root), "--manifest", str(manifest)])
    check(rc != 0, label + " must fail")


def expect_pass(root: Path, manifest: Path, label: str) -> None:
    rc = A.main(["--root", str(root), "--manifest", str(manifest)])
    check(rc == 0, label + " must pass")


def git(root: Path, *args: str) -> None:
    subprocess.run(["git", "-C", str(root), *args], check=True, capture_output=True)


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
    git(root, "init", "-q")
    for rel in FIXTURE_MODULES:
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
    # The production module list comes from the Git index, so stage the fixtures.
    git(root, "add", "--", "src")
    manifest = root / "docs" / "PUBLIC_API.txt"
    manifest.parent.mkdir(parents=True, exist_ok=True)
    manifest.write_text(A.render(A.collect(root)), encoding="utf-8")
    return root, manifest


root, manifest = make_repo()
try:
    expect_pass(root, manifest, "baseline")

    core = root / FIXTURE_MODULES[0]
    second = root / FIXTURE_MODULES[1]
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

    # Moving an unchanged function between production modules is API drift too.
    core.write_text('Attribute VB_Name = "M"\nOption Explicit\n', encoding="utf-8")
    second.write_text(baseline, encoding="utf-8")
    expect_fail(root, manifest, "module movement")

    # PROB_* linkage is project-internal and deliberately excluded from the product manifest.
    core.write_text(
        baseline + "\nPublic Function PROB_ProjectHelper(ByVal X As Double) As Double\n"
        "    PROB_ProjectHelper = X\nEnd Function\n",
        encoding="utf-8",
    )
    second.write_text('Attribute VB_Name = "Empty"\nOption Explicit\n', encoding="utf-8")
    expect_pass(root, manifest, "project-scoped Public PROB helper")

    core.write_text(
        baseline + "\nPrivate Function PROB_Internal(ByVal X As Double) As Double\n"
        "    PROB_Internal = X\nEnd Function\n",
        encoding="utf-8",
    )
    expect_pass(root, manifest, "Private PROB helper")

    # Public exporters outside src/ are intentionally not part of the product API.
    exporter = root / "benchmark" / "SomeExporter.bas"
    exporter.parent.mkdir(parents=True, exist_ok=True)
    exporter.write_text(
        'Attribute VB_Name = "SomeExporter"\nPublic Sub Export_Something()\nEnd Sub\n',
        encoding="utf-8",
    )
    expect_pass(root, manifest, "benchmark Public Sub outside production API")

    # The production module list is derived from the tracked src/ files.
    extra_api = (
        'Attribute VB_Name = "Extra"\nOption Explicit\n\n'
        "Public Function K_STATS_Extra(ByVal X As Double) As Double\n"
        "    K_STATS_Extra = X\nEnd Function\n"
    )
    untracked = root / "src" / "M_STATS_PROBDIST_EXTRA.bas"
    untracked.write_text(extra_api, encoding="utf-8")
    check(
        "src/M_STATS_PROBDIST_EXTRA.bas" not in A.production_modules(root),
        "present-but-untracked src module must not be picked up",
    )
    expect_pass(root, manifest, "present-but-untracked src module")
    git(root, "add", "--", str(untracked.relative_to(root)))
    check(
        "src/M_STATS_PROBDIST_EXTRA.bas" in A.production_modules(root),
        "newly tracked src module must be picked up",
    )
    expect_fail(root, manifest, "newly tracked src module with K_STATS API")
    git(root, "rm", "-q", "--cached", "--", str(untracked.relative_to(root)))
    untracked.unlink()
    expect_pass(root, manifest, "untracked module removed again")

    nested = root / "src" / "sub" / "M_STATS_PROBDIST_NESTED.bas"
    nested.parent.mkdir(parents=True, exist_ok=True)
    nested.write_text(extra_api.replace("Extra", "Nested"), encoding="utf-8")
    git(root, "add", "--", str(nested.relative_to(root)))
    expect_fail(root, manifest, "tracked nested src module with K_STATS API")
    git(root, "rm", "-q", "--cached", "--", str(nested.relative_to(root)))
    shutil.rmtree(nested.parent)

    # A tracked production module that is missing on disk fails closed.
    third = root / FIXTURE_MODULES[2]
    third_text = third.read_text(encoding="utf-8")
    third.unlink()
    expect_fail(root, manifest, "tracked production module missing on disk")
    third.write_text(third_text, encoding="utf-8")
    expect_pass(root, manifest, "tracked production module restored")

    # Without Git the tracked set is unknown: fail closed, never glob.
    no_git = Path(tempfile.mkdtemp())
    try:
        shutil.copytree(root / "src", no_git / "src")
        shutil.copytree(root / "docs", no_git / "docs")
        expect_fail(no_git, no_git / "docs" / "PUBLIC_API.txt", "root that is not a Git work tree")
        nested_root = root / "nested"
        shutil.copytree(no_git, nested_root)
        expect_fail(nested_root, nested_root / "docs" / "PUBLIC_API.txt",
                    "root below the Git top level")
        shutil.rmtree(nested_root)
    finally:
        shutil.rmtree(no_git, ignore_errors=True)
    expect_pass(root, manifest, "baseline after module-list fixtures")
finally:
    shutil.rmtree(root, ignore_errors=True)

if FAILS:
    print("FAIL: public API checker fixtures")
    for failure in FAILS:
        print("  - " + failure)
    raise SystemExit(1)
print("PASS: public API checker fixtures (name, type, ByVal/ByRef, Optional/default, "
      "parameter order, return type, additive API, module movement, a newly tracked or "
      "missing src module and a non-Git root fail; body/format/PROB-helper/"
      "benchmark-exporter edits and untracked src files pass)")
