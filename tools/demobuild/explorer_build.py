#!/usr/bin/env python3
"""Build the Explorer package from the user's workbook, without Excel.

Run from the repository: python tools/demobuild/explorer_build.py
Output (ignored by Git) in test-results/demo/:

  Probability-Distributions-Demo.xlsm   the baseline, with its VBA project
                                        rebuilt from the sources of this commit
  explorer-plan.json                    what Evolve-Explorer.ps1 applies in Excel
  manifest.json                         inputs and digests

The baseline, examples/demo/baseline/, is the workbook the user formatted by
hand: every part except xl/vbaProject.bin is kept byte for byte, so sheets,
styles, charts, the Ribbon and the navigation bar stay as the user left them.
The VBA project holds the library modules of src/, the demo module, the
Explorer module and the Function Wizard registration, plus the workbook's event
code. Evolve-Explorer.ps1 then writes the formulas, names and validation in
desktop Excel. Excel compiles the project when the workbook opens; nothing here
runs it.
"""
from __future__ import annotations

import hashlib
import io
import json
import sys
import zipfile
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT / "tools" / "devbuild"))
sys.path.insert(0, str(Path(__file__).resolve().parent))
from build_workbook import git, module_text, project_id  # noqa: E402
from vba_project import DOCUMENT, Module, build_vba_project, document_module  # noqa: E402

import build_demo  # noqa: E402
import explorer_plan  # noqa: E402

DEMO = ROOT / "examples" / "demo"
BASELINE = DEMO / "baseline" / "Explorer-baseline.xlsm"
WORKBOOK_NAME = "Probability-Distributions-Demo.xlsm"

EXPLORER_EVENTS = (
    "\r\nPrivate Sub Workbook_Open()\r\n"
    "    M_STATS_PROBDIST_REGISTER.PROB_Register_PublicUDFs\r\n"
    "    M_STATS_PROBDIST_EXPLORER.ExplorerOpened\r\n"
    "End Sub\r\n"
    "\r\nPrivate Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)\r\n"
    '    If Sh.CodeName = "shExplorer" Then M_STATS_PROBDIST_EXPLORER.ExplorerInputsChanged Target\r\n'
    "End Sub\r\n"
)


def vba_modules() -> tuple[list[Module], list[Path]]:
    sheets = build_demo.sheet_map()
    sources = sorted((ROOT / "src").glob("*.bas")) + [
        DEMO / "M_STATS_PROBDIST_DEMO.bas", DEMO / "M_STATS_PROBDIST_EXPLORER.bas",
        DEMO / "M_STATS_PROBDIST_REGISTER.bas"]
    base = build_demo.workbook_module()
    modules = [Module(base.name, base.source + EXPLORER_EVENTS, DOCUMENT)]
    modules += [document_module(sheet.code, False) for sheet in sheets]
    modules += [Module(path.stem, module_text(path.read_bytes())) for path in sources]
    return modules, sources


def build() -> tuple[bytes, dict[str, Any], dict[str, str]]:
    commit = git(ROOT, "rev-parse", "HEAD")
    modules, sources = vba_modules()
    with zipfile.ZipFile(BASELINE) as package:
        infos = package.infolist()
        parts = {info.filename: package.read(info.filename) for info in infos}
    if "xl/vbaProject.bin" not in parts:
        raise ValueError("the baseline has no VBA project")
    parts["xl/vbaProject.bin"] = build_vba_project(modules, project_id(commit))
    out = io.BytesIO()
    with zipfile.ZipFile(out, "w") as archive:
        for info in infos:
            copy = zipfile.ZipInfo(info.filename, info.date_time)
            copy.compress_type = zipfile.ZIP_DEFLATED
            copy.create_system = 0
            copy.external_attr = 0
            archive.writestr(copy, parts[info.filename])
    plan = explorer_plan.plan()
    inputs = sources + [BASELINE, DEMO / "explorer-families.json", Path(__file__),
                        Path(explorer_plan.__file__), Path(build_demo.__file__),
                        ROOT / "tools/devbuild/vba_project.py", ROOT / "tools/devbuild/build_workbook.py",
                        ROOT / "VERSION"]
    hashes = {str(path.relative_to(ROOT)): hashlib.sha256(path.read_bytes()).hexdigest() for path in inputs}
    return out.getvalue(), plan, hashes


def main() -> int:
    output = ROOT / "test-results" / "demo"
    output.mkdir(parents=True, exist_ok=True)
    data, plan, hashes = build()
    (output / WORKBOOK_NAME).write_bytes(data)
    (output / "explorer-plan.json").write_text(json.dumps(plan, indent=1), encoding="utf-8")
    manifest = {"commit": git(ROOT, "rev-parse", "HEAD"), "status": "UNVERIFIED IN EXCEL",
                "dirty": bool(git(ROOT, "status", "--porcelain")),
                "workbook_sha256": hashlib.sha256(data).hexdigest(), "inputs": hashes}
    (output / "manifest.json").write_text(json.dumps(manifest, indent=2) + "\n", encoding="utf-8")
    print(output)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
