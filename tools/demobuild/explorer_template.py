#!/usr/bin/env python3
"""The macro-free template of the Explorer workbook, and its way back.

  python tools/demobuild/explorer_template.py WORKBOOK.xlsm

writes src/workbook/PROBDIST_Template.xlsx from a workbook that Excel has
saved, normally the Probability-Distributions-Demo.xlsm that the Distribution
Explorer workflow verified: every part is kept as Excel wrote it except the
VBA project, which is dropped with its relationship and content type, and the
workbook part, whose content type becomes the one of an .xlsx, and the
Ribbon, written from src/workbook/customUI14.xml. The sensitivity label
that Office stores in docMetadata/LabelInfo.xml (it names the tenant of the
machine that saved the workbook) is removed with its relationship and content
type, and so are the hidden names that carry xlm="1" (Excel's marker for an
Excel 4.0 macro name; Excel writes them for the parameters of LET and LAMBDA
and rebuilds them when it loads the formulas), because Excel refuses to open
an .xlsx that holds them: "a macro-free file that contains macro-enabled
content". The sheets, formulas, names, validation and charts stay, so the
template is the workbook without its code.

explorer_build.py goes the other way: enable_macros() puts a VBA project,
built from the sources of the commit, back into the template's parts. Update
the template, and its SHA-256 in .github/binary-allowlist.json, when the
layout of the sheets changes; a change to the VBA source needs no new
template.
"""
from __future__ import annotations

import io
import re
import sys
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
TEMPLATE = ROOT / "src" / "workbook" / "PROBDIST_Template.xlsx"
RIBBON = ROOT / "src" / "workbook" / "customUI14.xml"
RIBBON_PART = "customUI/customUI14.xml"

VBA_PART = "xl/vbaProject.bin"
CONTENT_TYPES = "[Content_Types].xml"
WORKBOOK_RELS = "xl/_rels/workbook.xml.rels"
MACRO_MAIN = "application/vnd.ms-excel.sheet.macroEnabled.main+xml"
PLAIN_MAIN = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"
VBA_TYPE = "application/vnd.ms-office.vbaProject"
VBA_RELATIONSHIP = "http://schemas.microsoft.com/office/2006/relationships/vbaProject"
# The relationship Id that enable_macros gives the project; strip_macros
# removes whatever Id Excel gave it.
VBA_ID = "rIdVbaProject"

WORKBOOK_OVERRIDE = re.compile(r'<Override PartName="/xl/workbook\.xml" ContentType="([^"]+)"/>')
VBA_OVERRIDE = re.compile(r'<Override PartName="/xl/vbaProject\.bin" ContentType="[^"]+"/>')
MACRO_NAME = re.compile(r'<definedName [^>]*xlm="1"[^>]*>[^<]*</definedName>')
WORKBOOK_PART = "xl/workbook.xml"
LABEL_PART = "docMetadata/LabelInfo.xml"
LABEL_OVERRIDE = re.compile(r'<Override PartName="/docMetadata/LabelInfo\.xml" ContentType="[^"]+"/>')
LABEL_REL = re.compile(r'<Relationship Id="[^"]+" Type="[^"]+" Target="docMetadata/LabelInfo\.xml"/>')
PACKAGE_RELS = "_rels/.rels"
VBA_REL = re.compile(r'<Relationship Id="[^"]+" Type="' + re.escape(VBA_RELATIONSHIP)
                     + r'" Target="vbaProject\.bin"/>')


class TemplateError(ValueError):
    pass


def _replace_one(pattern: re.Pattern[str], text: str, new: str, what: str) -> str:
    found = pattern.findall(text)
    if len(found) != 1:
        raise TemplateError(f"expected one {what}, found {len(found)}")
    return pattern.sub(new, text)


def strip_macros(parts: dict[str, bytes]) -> dict[str, bytes]:
    """The parts of a macro-enabled workbook as those of an .xlsx."""
    if VBA_PART not in parts:
        raise TemplateError(f"no {VBA_PART}: not a macro-enabled workbook")
    out = {name: data for name, data in parts.items() if name != VBA_PART}
    types = parts[CONTENT_TYPES].decode("utf-8")
    types = _replace_one(VBA_OVERRIDE, types, "", "VBA project content type")
    types = _replace_one(WORKBOOK_OVERRIDE, types,
                         f'<Override PartName="/xl/workbook.xml" ContentType="{PLAIN_MAIN}"/>',
                         "workbook content type")
    out[WORKBOOK_PART] = MACRO_NAME.sub("", parts[WORKBOOK_PART].decode("utf-8")).encode("utf-8")
    if LABEL_PART in out:
        del out[LABEL_PART]
        types = LABEL_OVERRIDE.sub("", types)
        package = out[PACKAGE_RELS].decode("utf-8")
        out[PACKAGE_RELS] = LABEL_REL.sub("", package).encode("utf-8")
    out[CONTENT_TYPES] = types.encode("utf-8")
    rels = parts[WORKBOOK_RELS].decode("utf-8")
    out[WORKBOOK_RELS] = _replace_one(VBA_REL, rels, "", "VBA project relationship").encode("utf-8")
    return out


def enable_macros(parts: dict[str, bytes], project: bytes) -> dict[str, bytes]:
    """The parts of the macro-free template with `project` as their VBA project."""
    if VBA_PART in parts:
        raise TemplateError(f"the template already holds {VBA_PART}")
    types = parts[CONTENT_TYPES].decode("utf-8")
    if WORKBOOK_OVERRIDE.findall(types) != [PLAIN_MAIN]:
        raise TemplateError("the template's workbook part is not that of an .xlsx")
    types = WORKBOOK_OVERRIDE.sub(
        f'<Override PartName="/xl/workbook.xml" ContentType="{MACRO_MAIN}"/>'
        f'<Override PartName="/{VBA_PART}" ContentType="{VBA_TYPE}"/>', types)
    rels = parts[WORKBOOK_RELS].decode("utf-8")
    if f'Id="{VBA_ID}"' in rels or "</Relationships>" not in rels:
        raise TemplateError(f"cannot add the relationship {VBA_ID} to {WORKBOOK_RELS}")
    rels = rels.replace("</Relationships>", f'<Relationship Id="{VBA_ID}" Type="{VBA_RELATIONSHIP}" '
                        'Target="vbaProject.bin"/></Relationships>')
    out = dict(parts)
    out[CONTENT_TYPES] = types.encode("utf-8")
    out[WORKBOOK_RELS] = rels.encode("utf-8")
    out[VBA_PART] = project
    return out


def read_package(path: Path) -> tuple[list[zipfile.ZipInfo], dict[str, bytes]]:
    with zipfile.ZipFile(path) as package:
        infos = package.infolist()
        return infos, {info.filename: package.read(info.filename) for info in infos}


def write_package(infos: list[zipfile.ZipInfo], parts: dict[str, bytes]) -> bytes:
    """A package of `parts`: those of `infos` in their order and with their
    times, then any new part, all deflated."""
    order = [info.filename for info in infos if info.filename in parts]
    order += sorted(name for name in parts if name not in order)
    times = {info.filename: info.date_time for info in infos}
    out = io.BytesIO()
    with zipfile.ZipFile(out, "w") as archive:
        for name in order:
            entry = zipfile.ZipInfo(name, times.get(name, (1980, 1, 1, 0, 0, 0)))
            entry.compress_type = zipfile.ZIP_DEFLATED
            entry.create_system = 0
            entry.external_attr = 0
            archive.writestr(entry, parts[name])
    return out.getvalue()


def main(argv: list[str]) -> int:
    if len(argv) != 1:
        print(__doc__, file=sys.stderr)
        return 2
    infos, parts = read_package(Path(argv[0]))
    if RIBBON_PART not in parts:
        raise TemplateError(f"no {RIBBON_PART}: not the Explorer workbook")
    parts[RIBBON_PART] = RIBBON.read_bytes()
    TEMPLATE.write_bytes(write_package(infos, strip_macros(parts)))
    print(TEMPLATE)
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
