#!/usr/bin/env python3
"""Build the Excel workbook under test from the source of one commit.

Usage: python3 tools/devbuild/build_workbook.py [OUTPUT] [--manifest PATH] [--root DIR]
Default: test-results/workbook-build/probdist-regression.xlsm and, beside it,
probdist-regression.manifest.json. Git ignores test-results/.

The workbook's VBA project holds exactly the modules that
.github/excel-evidence-policy.json lists in regression_sources (the library
modules under src/ and the regression harness under tests/), in that order,
plus the empty document modules of the workbook and its one sheet. A listed
module that is missing, a .bas under src/ or tests/ that is not listed, and a
module whose VB_Name differs from its file name stop the build. vba_project.py
writes the project from the [MS-OVBA] and [MS-CFB] specifications with
compressed source and no p-code cache, so Excel compiles it from source on
open, and the build needs no Trust access to the VBA project object model.
No tracked file is changed: every module is embedded as it is, with CRLF line
endings.

The commit is stamped on the Build sheet, under the workbook-level defined
names build_commit and build_commit_time. A tracked change anywhere in the
working tree, or an untracked listed module, stamps the build as not
evidence-grade. The manifest records the commit, the SHA-256 of every module
(of its Git blob and of the text embedded in the workbook) and of the
workbook; tools/devbuild/verify_workbook.py checks it.

For a given commit and working tree the output is byte-identical: the stamp
carries the commit's committer time, not the clock; the project ID derives
from the commit; the package parts are stored, not deflated, in a fixed order
with a fixed timestamp, so the bytes do not depend on the zlib build. The
build runs no Excel and compiles nothing; the first Excel run proves that the
project compiles. Standard library only.
"""

from __future__ import annotations

import argparse
import hashlib
import io
import json
import re
import subprocess
import sys
import uuid
import zipfile
from collections.abc import Sequence
from dataclasses import dataclass
from pathlib import Path
from typing import Any
from xml.sax.saxutils import escape, quoteattr

sys.path.insert(0, str(Path(__file__).resolve().parent))
from vba_project import (  # noqa: E402
    VBA_ENCODING,
    Module,
    build_vba_project,
    document_module,
)

ROOT = Path(__file__).resolve().parents[2]
TOOL = "tools/devbuild/build_workbook.py"
POLICY = ".github/excel-evidence-policy.json"
SOURCE_DIRECTORIES = ("src", "tests")
DEFAULT_OUTPUT = "test-results/workbook-build/probdist-regression.xlsm"
MANIFEST_SUFFIX = ".manifest.json"
FORMAT = "probdist-workbook-build-1"

WORKBOOK_CODE_NAME = "ThisWorkbook"
SHEET_TAB = "Build"
SHEET_CODE_NAME = "shBuild"
# Defined name -> (cell on the Build sheet, label written to its left).
STAMP_COMMIT = "build_commit"
STAMP_TIME = "build_commit_time"
STAMP_CELLS = {STAMP_COMMIT: ("B3", "Commit"), STAMP_TIME: ("B4", "Commit time")}
SOURCES_CELL = ("B5", "Sources")
TITLE = (f"Built from source by {TOOL}; Excel compiles the VBA project on open. "
         "Never commit this workbook.")
NOT_EVIDENCE = " (working tree with changes: NOT evidence-grade)"
ZIP_TIME = (1980, 1, 1, 0, 0, 0)
VB_NAME = re.compile(r'^Attribute VB_Name = "([^"]+)"$')

MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
RELATIONSHIPS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PACKAGE_RELATIONSHIPS = "http://schemas.openxmlformats.org/package/2006/relationships"
VBA_TYPE = "http://schemas.microsoft.com/office/2006/relationships/vbaProject"
XML_HEADER = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'


class BuildError(Exception):
    """The policy or the source cannot be built."""


@dataclass(frozen=True)
class Source:
    path: str          # repository-relative POSIX path, as the policy lists it
    name: str          # the module name: the file stem, equal to its VB_Name
    text: str          # the module text embedded in the workbook (CRLF)


@dataclass(frozen=True)
class Build:
    commit: str
    time: str
    clean: bool
    sources: tuple[Source, ...]
    workbook: Path
    manifest: Path


def sha256(data: bytes) -> str:
    return "sha256:" + hashlib.sha256(data).hexdigest()


def git(root: Path, *args: str) -> str:
    return git_bytes(root, *args).decode("utf-8").strip()


def git_bytes(root: Path, *args: str) -> bytes:
    completed = subprocess.run(["git", "-C", str(root), *args], capture_output=True,
                               check=False)
    if completed.returncode:
        message = completed.stderr.decode("utf-8", "replace").strip()
        raise BuildError(f"git {' '.join(args)}: {message}")
    return completed.stdout


def regression_sources(policy: dict[str, Any]) -> list[str]:
    """The policy's regression_sources, checked: unique .bas files directly under
    src/ or tests/."""
    listed = policy.get("regression_sources")
    if not isinstance(listed, list) or not listed \
            or not all(isinstance(item, str) for item in listed):
        raise BuildError(f"{POLICY}: regression_sources must be a non-empty list of paths")
    for item in listed:
        parts = item.split("/")
        if len(parts) != 2 or parts[0] not in SOURCE_DIRECTORIES or not item.endswith(".bas"):
            raise BuildError(f"{POLICY}: {item} is not a .bas file directly under "
                             f"{' or '.join(d + '/' for d in SOURCE_DIRECTORIES)}")
    stems = [Path(item).stem.casefold() for item in listed]
    if len(set(stems)) != len(stems):
        raise BuildError(f"{POLICY}: regression_sources lists a module twice")
    return list(listed)


def load_policy(root: Path) -> list[str]:
    try:
        policy = json.loads((root / POLICY).read_text(encoding="utf-8"))
    except (OSError, ValueError) as error:
        raise BuildError(f"{POLICY}: {error}") from error
    if not isinstance(policy, dict):
        raise BuildError(f"{POLICY}: not a JSON object")
    return regression_sources(policy)


def module_text(data: bytes) -> str:
    """A module's source as the VBE stores it: CP1252 text with CRLF line endings."""
    try:
        text = data.decode(VBA_ENCODING)
    except UnicodeDecodeError as error:
        raise ValueError(f"not {VBA_ENCODING} text: {error}") from error
    text = text.replace("\r\n", "\n").replace("\n", "\r\n")
    return text if text.endswith("\r\n") else text + "\r\n"


def read_sources(root: Path, listed: Sequence[str]) -> list[Source]:
    """Every listed module, in policy order; a missing, unlisted or misnamed module
    stops the build."""
    missing = [item for item in listed if not (root / item).is_file()]
    if missing:
        raise BuildError(f"{', '.join(missing)}: listed in {POLICY} regression_sources but "
                         "not found")
    present = sorted(f"{directory}/{path.name}" for directory in SOURCE_DIRECTORIES
                     for path in (root / directory).glob("*.bas"))
    extra = [item for item in present if item not in listed]
    if extra:
        raise BuildError(f"{', '.join(extra)}: not listed in {POLICY} regression_sources; "
                         "list it there, or move it out of src/ and tests/")
    sources = []
    for item in listed:
        name = Path(item).stem
        try:
            text = module_text((root / item).read_bytes())
        except ValueError as error:
            raise BuildError(f"{item}: {error}") from error
        match = VB_NAME.match(text.split("\r\n", 1)[0])
        if not match or match.group(1) != name:
            raise BuildError(f'{item}: the first line must be Attribute VB_Name = "{name}"')
        sources.append(Source(item, name, text))
    return sources


def is_clean(root: Path, listed: Sequence[str]) -> bool:
    """No tracked change anywhere, and every listed module tracked."""
    if git(root, "status", "--porcelain", "--untracked-files=no"):
        return False
    tracked = set(git(root, "ls-files", "--", *listed).splitlines())
    return all(item in tracked for item in listed)


def column_number(letters: str) -> int:
    number = 0
    for letter in letters:
        number = number * 26 + ord(letter) - 64
    return number


def column_letters(number: int) -> str:
    letters = ""
    while number:
        number, remainder = divmod(number - 1, 26)
        letters = chr(65 + remainder) + letters
    return letters


def split_cell(cell: str) -> tuple[int, int]:
    match = re.match(r"^([A-Z]{1,3})([1-9][0-9]*)$", cell)
    assert match is not None, cell
    return int(match.group(2)), column_number(match.group(1))


def stamp_values(commit: str, time: str, clean: bool) -> dict[str, str]:
    return {STAMP_COMMIT: commit + ("" if clean else NOT_EVIDENCE), STAMP_TIME: time}


def sheet_cells(stamp: dict[str, str], sources: Sequence[Source]) -> dict[tuple[int, int], str]:
    """The text cells of the Build sheet, by (row, column)."""
    cells: dict[tuple[int, int], str] = {(1, 1): TITLE}
    for name, (cell, label) in STAMP_CELLS.items():
        row, column = split_cell(cell)
        cells[(row, column - 1)] = label
        cells[(row, column)] = stamp[name]
    row, column = split_cell(SOURCES_CELL[0])
    cells[(row, column - 1)] = SOURCES_CELL[1]
    cells[(row, column)] = ", ".join(source.path for source in sources)
    return cells


def worksheet_xml(cells: dict[tuple[int, int], str]) -> str:
    rows: dict[int, list[tuple[int, str]]] = {}
    for (row, column), text in cells.items():
        rows.setdefault(row, []).append((column, text))
    body = "".join(
        f'<row r="{row}">' + "".join(
            f'<c r="{column_letters(column)}{row}" t="inlineStr"><is>'
            f'<t xml:space="preserve">{escape(text)}</t></is></c>'
            for column, text in sorted(rows[row])) + "</row>"
        for row in sorted(rows))
    return (f'{XML_HEADER}<worksheet xmlns="{MAIN}" xmlns:r="{RELATIONSHIPS}">'
            f"<sheetPr codeName={quoteattr(SHEET_CODE_NAME)}/>"
            f"<sheetData>{body}</sheetData></worksheet>")


def workbook_xml() -> str:
    names = ""
    for name, (cell, _) in sorted(STAMP_CELLS.items()):
        row, column = split_cell(cell)
        reference = f"{SHEET_TAB}!${column_letters(column)}${row}"
        names += f"<definedName name={quoteattr(name)}>{escape(reference)}</definedName>"
    return (f'{XML_HEADER}<workbook xmlns="{MAIN}" xmlns:r="{RELATIONSHIPS}">'
            f"<workbookPr codeName={quoteattr(WORKBOOK_CODE_NAME)}/>"
            f"<bookViews><workbookView/></bookViews><sheets>"
            f'<sheet name={quoteattr(SHEET_TAB)} sheetId="1" r:id="rIdSheet1"/></sheets>'
            f"<definedNames>{names}</definedNames></workbook>")


STYLES = (f'{XML_HEADER}<styleSheet xmlns="{MAIN}">'
          '<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>'
          '<fills count="2"><fill><patternFill patternType="none"/></fill>'
          '<fill><patternFill patternType="gray125"/></fill></fills>'
          '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
          '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>'
          '</cellStyleXfs><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" '
          'borderId="0" xfId="0"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" '
          'xfId="0" builtinId="0"/></cellStyles></styleSheet>')


# The package parts, in the order they are written.
PARTS = ("[Content_Types].xml", "_rels/.rels", "xl/workbook.xml", "xl/_rels/workbook.xml.rels",
         "xl/styles.xml", "xl/worksheets/sheet1.xml", "xl/vbaProject.bin")


def package(cells: dict[tuple[int, int], str], vba_bin: bytes) -> bytes:
    """The .xlsm package: the same bytes for the same input, on any platform."""
    parts = {
        "[Content_Types].xml": (
            f'{XML_HEADER}<Types xmlns="http://schemas.openxmlformats.org/package/2006/'
            'content-types"><Default Extension="rels" ContentType="application/'
            'vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" '
            'ContentType="application/xml"/><Default Extension="bin" '
            'ContentType="application/vnd.ms-office.vbaProject"/><Override '
            'PartName="/xl/workbook.xml" ContentType="application/'
            'vnd.ms-excel.sheet.macroEnabled.main+xml"/><Override PartName="/xl/styles.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.'
            'styles+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="'
            'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
            "</Types>").encode("utf-8"),
        "_rels/.rels": (
            f'{XML_HEADER}<Relationships xmlns="{PACKAGE_RELATIONSHIPS}"><Relationship '
            f'Id="rId1" Type="{RELATIONSHIPS}/officeDocument" Target="xl/workbook.xml"/>'
            "</Relationships>").encode("utf-8"),
        "xl/workbook.xml": workbook_xml().encode("utf-8"),
        "xl/_rels/workbook.xml.rels": (
            f'{XML_HEADER}<Relationships xmlns="{PACKAGE_RELATIONSHIPS}">'
            f'<Relationship Id="rIdSheet1" Type="{RELATIONSHIPS}/worksheet" '
            'Target="worksheets/sheet1.xml"/>'
            f'<Relationship Id="rIdStyles" Type="{RELATIONSHIPS}/styles" Target="styles.xml"/>'
            f'<Relationship Id="rIdVBA" Type="{VBA_TYPE}" Target="vbaProject.bin"/>'
            "</Relationships>").encode("utf-8"),
        "xl/styles.xml": STYLES.encode("utf-8"),
        "xl/worksheets/sheet1.xml": worksheet_xml(cells).encode("utf-8"),
        "xl/vbaProject.bin": vba_bin,
    }
    assert tuple(parts) == PARTS
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w") as archive:
        for name, data in parts.items():
            info = zipfile.ZipInfo(name, date_time=ZIP_TIME)
            info.compress_type = zipfile.ZIP_STORED
            # zipfile records the host system by default; fix it so that a build on
            # Windows and one on Linux give the same bytes.
            info.create_system = 0
            info.external_attr = 0
            archive.writestr(info, data)
    return buffer.getvalue()


def project_id(commit: str) -> str:
    return "{" + str(uuid.uuid5(uuid.NAMESPACE_URL, f"probdist-workbook:{commit}")).upper() + "}"


def vba_modules(sources: Sequence[Source]) -> list[Module]:
    return [document_module(WORKBOOK_CODE_NAME, True), document_module(SHEET_CODE_NAME, False),
            *(Module(source.name, source.text) for source in sources)]


def git_blob_digest(root: Path, commit: str, path: str) -> str | None:
    try:
        return sha256(git_bytes(root, "cat-file", "blob", f"{commit}:{path}"))
    except BuildError:
        return None


def manifest(root: Path, build: Build, workbook: bytes) -> dict[str, Any]:
    return {
        "format": FORMAT,
        "tool": TOOL,
        "policy": POLICY,
        "commit": build.commit,
        "commit_time": build.time,
        "clean": build.clean,
        "stamp": stamp_values(build.commit, build.time, build.clean),
        "vba_project": {"compile_from_source": True, "p_code": False,
                        "document_modules": [WORKBOOK_CODE_NAME, SHEET_CODE_NAME]},
        "modules": [{"name": source.name, "path": source.path,
                     "git_blob_sha256": git_blob_digest(root, build.commit, source.path),
                     "module_sha256": sha256(source.text.encode(VBA_ENCODING))}
                    for source in build.sources],
        "workbook": {"file": build.workbook.name, "sha256": sha256(workbook)},
    }


def default_manifest(target: Path) -> Path:
    return target.with_name(target.stem + MANIFEST_SUFFIX)


def render(sources: Sequence[Source], commit: str, time: str, clean: bool) -> bytes:
    """The workbook bytes: a function of the modules and the stamp only."""
    vba_bin = build_vba_project(vba_modules(sources), project_id(commit))
    return package(sheet_cells(stamp_values(commit, time, clean), sources), vba_bin)


def sources_at(root: Path, commit: str, listed: Sequence[str]) -> list[Source]:
    """The listed modules as the commit stores them, with CRLF line endings."""
    sources = []
    for item in listed:
        blob = git_bytes(root, "cat-file", "blob", f"{commit}:{item}")
        try:
            sources.append(Source(item, Path(item).stem, module_text(blob)))
        except ValueError as error:
            raise BuildError(f"{item} at {commit[:12]}: {error}") from error
    return sources


def build(root: Path, target: Path, manifest_path: Path | None = None) -> Build:
    """Write the workbook to ``target`` and its manifest beside it (or to
    ``manifest_path``); describe what they hold."""
    listed = load_policy(root)
    sources = read_sources(root, listed)
    commit = git(root, "rev-parse", "HEAD")
    result = Build(commit, git(root, "show", "-s", "--format=%cI", "HEAD"),
                   is_clean(root, listed), tuple(sources), target,
                   manifest_path or default_manifest(target))
    data = render(sources, result.commit, result.time, result.clean)
    target.parent.mkdir(parents=True, exist_ok=True)
    target.write_bytes(data)
    result.manifest.parent.mkdir(parents=True, exist_ok=True)
    result.manifest.write_text(json.dumps(manifest(root, result, data), indent=2) + "\n",
                               encoding="utf-8", newline="\n")
    return result


def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=(__doc__ or "").splitlines()[0])
    parser.add_argument("output", nargs="?", type=Path,
                        help=f"the .xlsm to write (default: {DEFAULT_OUTPUT})")
    parser.add_argument("--manifest", type=Path,
                        help=f"the manifest to write (default: OUTPUT with {MANIFEST_SUFFIX})")
    parser.add_argument("--root", type=Path, default=ROOT)
    args = parser.parse_args(argv)
    root = args.root.resolve()
    target = args.output or root / DEFAULT_OUTPUT
    try:
        result = build(root, target, args.manifest)
    except (BuildError, ValueError, OSError) as error:
        print(f"FAIL: {error}")
        return 1
    grade = "" if result.clean else NOT_EVIDENCE
    print(f"built {result.workbook}: commit {result.commit}{grade}")
    print(f"modules: {', '.join(source.name for source in result.sources)}")
    print(f"manifest: {result.manifest}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
