#!/usr/bin/env python3
"""Build the Explorer package from the template, without Excel.

Run from the repository: python tools/demobuild/explorer_build.py
Output (ignored by Git) in test-results/demo/:

  Probability-Distributions-Demo.xlsm   the template with a VBA project built
                                        from the sources of this commit
  explorer-plan.json                    what Evolve-Explorer.ps1 applies in Excel
  probabilities-plan.json               what Evolve-Probabilities.ps1 applies in Excel
  manifest.json                         inputs and digests

The template, src/workbook/PROBDIST_Template.xlsx, is the Explorer workbook
without its code (explorer_template.py): every part is kept byte for byte,
except the Ribbon, written from src/workbook/customUI14.xml, so sheets,
styles, charts and the navigation bar stay as the user left them. The VBA
project holds the library modules of src/modules/, the modules of src/workbook/ (the
demo module, the Explorer module and the Function Wizard registration) and
the workbook's event code, src/workbook/ThisWorkbook.cls; the sheet modules
hold no code. Evolve-Explorer.ps1 then writes the formulas, names and
validation in desktop Excel. Excel compiles the project when the workbook
opens; nothing here runs it.
"""
from __future__ import annotations

import hashlib
import json
import sys
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT / "tools" / "devbuild"))
sys.path.insert(0, str(Path(__file__).resolve().parent))
from build_workbook import git, module_text, project_id  # noqa: E402
from vba_project import DOCUMENT, Module, build_vba_project, document_module  # noqa: E402

import build_demo  # noqa: E402
import explorer_plan  # noqa: E402
import explorer_template  # noqa: E402
import probabilities_plan  # noqa: E402

DEMO = ROOT / "examples" / "demo"
WORKBOOK = ROOT / "src" / "workbook"
TEMPLATE = explorer_template.TEMPLATE
RIBBON = explorer_template.RIBBON
RIBBON_PART = explorer_template.RIBBON_PART
THIS_WORKBOOK = WORKBOOK / "ThisWorkbook.cls"
WORKBOOK_NAME = "Probability-Distributions-Demo.xlsm"


def document_source(path: Path) -> str:
    """A document module's code as the VBE stores it: the exported .cls without
    the VERSION 1.0 CLASS block that comes before its attributes."""
    text = module_text(path.read_bytes())
    start = text.find('Attribute VB_Name = "')
    if start < 0:
        raise ValueError(f"{path}: no Attribute VB_Name line")
    return text[start:]


def vba_modules() -> tuple[list[Module], list[Path]]:
    sheets = build_demo.sheet_map()
    sources = sorted((ROOT / "src" / "modules").glob("*.bas")) + sorted(WORKBOOK.glob("*.bas"))
    modules = [Module("ThisWorkbook", document_source(THIS_WORKBOOK), DOCUMENT)]
    modules += [document_module(sheet.code, False) for sheet in sheets]
    modules += [Module(path.stem, module_text(path.read_bytes())) for path in sources]
    return modules, sources + [THIS_WORKBOOK]


def build() -> tuple[bytes, dict[str, Any], dict[str, str]]:
    commit = git(ROOT, "rev-parse", "HEAD")
    modules, sources = vba_modules()
    infos, parts = explorer_template.read_package(TEMPLATE)
    parts[RIBBON_PART] = RIBBON.read_bytes()
    parts = explorer_template.enable_macros(parts, build_vba_project(modules, project_id(commit)))
    data = explorer_template.write_package(infos, parts)
    plan = explorer_plan.plan()
    inputs = sources + [TEMPLATE, RIBBON, DEMO / "explorer-families.json", Path(__file__),
                        Path(explorer_plan.__file__), Path(probabilities_plan.__file__), Path(explorer_template.__file__),
                        Path(build_demo.__file__),
                        ROOT / "tools/devbuild/vba_project.py", ROOT / "tools/devbuild/build_workbook.py",
                        ROOT / "VERSION"]
    hashes = {str(path.relative_to(ROOT)): hashlib.sha256(path.read_bytes()).hexdigest() for path in inputs}
    return data, plan, hashes


def main() -> int:
    output = ROOT / "test-results" / "demo"
    output.mkdir(parents=True, exist_ok=True)
    data, plan, hashes = build()
    (output / WORKBOOK_NAME).write_bytes(data)
    (output / "explorer-plan.json").write_text(json.dumps(plan, indent=1), encoding="utf-8")
    (output / "probabilities-plan.json").write_text(json.dumps(probabilities_plan.plan(), indent=1), encoding="utf-8")
    manifest = {"commit": git(ROOT, "rev-parse", "HEAD"), "status": "UNVERIFIED IN EXCEL",
                "dirty": bool(git(ROOT, "status", "--porcelain")),
                "workbook_sha256": hashlib.sha256(data).hexdigest(), "inputs": hashes}
    (output / "manifest.json").write_text(json.dumps(manifest, indent=2) + "\n", encoding="utf-8")
    print(output)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
