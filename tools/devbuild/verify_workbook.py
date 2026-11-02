#!/usr/bin/env python3
"""Check that a built workbook is bound to its commit and its sources.

Usage: python3 tools/devbuild/verify_workbook.py MANIFEST [--workbook PATH]
       [--commit REV] [--require-verified] [--root DIR]

MANIFEST is the manifest that tools/devbuild/build_workbook.py wrote beside
the workbook; --workbook defaults to the file it names, in the same folder.
The binding is:

- VERIFIED when every check below passes and the build was clean;
- FAIL when any check finds a mismatch;
- UNVERIFIED when there is nothing to check against (no manifest, no
  workbook, or a commit this repository does not have) or the build came
  from a working tree with changes, which is never evidence-grade.

The checks: the workbook's SHA-256 equals the manifest's; its build stamp
(defined name build_commit) names the manifest's commit; its VBA project
compiles from source (_VBA_PROJECT version 0xFFFF, no p-code) and holds the
manifest's modules and document modules, no more and no fewer, each module
stream decompressing to text with the recorded SHA-256; the modules are
exactly the regression_sources of .github/excel-evidence-policy.json at the
commit, in its order; and each module's Git blob at the commit has the
recorded SHA-256, and so, for a clean build, has its embedded text. With
--commit, the manifest's commit must be that commit.

The exit code is 0 for VERIFIED, 1 for FAIL, and for UNVERIFIED 0, or 1 with
--require-verified. The check runs no Excel: VERIFIED binds the workbook to
its sources; it does not show that the project compiles or that a test
passes. Standard library only.
"""

from __future__ import annotations

import argparse
import io
import json
import posixpath
import re
import sys
import zipfile
from collections.abc import Sequence
from dataclasses import dataclass, field
from pathlib import Path, PurePosixPath
from typing import Any
from xml.etree import ElementTree

sys.path.insert(0, str(Path(__file__).resolve().parent))
import build_workbook  # noqa: E402
from vba_project import VBA_ENCODING, VBA_PROJECT_STREAM, CompoundFile, decompress  # noqa: E402

ROOT = build_workbook.ROOT
VERIFIED, FAIL, UNVERIFIED = "VERIFIED", "FAIL", "UNVERIFIED"
SPREADSHEET = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
RELATIONSHIP = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
PACKAGE = "{http://schemas.openxmlformats.org/package/2006/relationships}"
REFERENCE = re.compile(r"^(?:'((?:[^']|'')+)'|([^'!]+))!\$?([A-Z]{1,3})\$?([0-9]+)$")
PROJECT_MODULE = re.compile(r"^(Module|Document|Class|BaseClass|Package)=([^/\r\n]+)")


@dataclass
class Verdict:
    status: str = UNVERIFIED
    findings: list[str] = field(default_factory=list)
    notes: list[str] = field(default_factory=list)


@dataclass(frozen=True)
class Recorded:
    """What a manifest records."""

    commit: str
    clean: bool
    workbook_file: str
    workbook_sha256: str
    modules: tuple[dict[str, str | None], ...]
    document_modules: tuple[str, ...]


def load_manifest(path: Path) -> Recorded:
    data: dict[str, Any] = json.loads(path.read_text(encoding="utf-8"))
    if data.get("format") != build_workbook.FORMAT:
        raise ValueError(f"format is {data.get('format')!r}, expected {build_workbook.FORMAT}")
    if not isinstance(data["clean"], bool) or not re.fullmatch(r"[0-9a-f]{40}", data["commit"]):
        raise ValueError("commit or clean is malformed")
    modules = tuple({"name": str(item["name"]), "path": str(item["path"]),
                     "git_blob_sha256": item["git_blob_sha256"],
                     "module_sha256": str(item["module_sha256"])} for item in data["modules"])
    return Recorded(data["commit"], data["clean"], str(data["workbook"]["file"]),
                    str(data["workbook"]["sha256"]), modules,
                    tuple(str(name) for name in data["vba_project"]["document_modules"]))


def cell_text(package: zipfile.ZipFile, part: str, cell: str) -> str | None:
    """The text of one cell of a worksheet part: inline, shared or formula string."""
    root = ElementTree.fromstring(package.read(part))
    for element in root.iter(f"{SPREADSHEET}c"):
        if element.get("r") != cell:
            continue
        kind = element.get("t")
        if kind == "inlineStr":
            return "".join(node.text or "" for node in element.iter(f"{SPREADSHEET}t"))
        value = element.find(f"{SPREADSHEET}v")
        if value is None or value.text is None:
            return None
        if kind == "s":
            shared = ElementTree.fromstring(package.read("xl/sharedStrings.xml"))
            item = list(shared.iter(f"{SPREADSHEET}si"))[int(value.text)]
            return "".join(node.text or "" for node in item.iter(f"{SPREADSHEET}t"))
        return value.text
    return None


def read_stamp(package: zipfile.ZipFile, name: str) -> str | None:
    """The value of the workbook-level defined name ``name``, read from the parts."""
    book = ElementTree.fromstring(package.read("xl/workbook.xml"))
    target = next((item.text or "" for item in book.iter(f"{SPREADSHEET}definedName")
                   if item.get("name", "").casefold() == name.casefold()
                   and item.get("localSheetId") is None), None)
    match = REFERENCE.match(target or "")
    if not match:
        return None
    tab = (match.group(1) or "").replace("''", "'") or match.group(2)
    sheet = next((item for item in book.iter(f"{SPREADSHEET}sheet")
                  if item.get("name") == tab), None)
    if sheet is None:
        return None
    relations = ElementTree.fromstring(package.read("xl/_rels/workbook.xml.rels"))
    relation = next((item for item in relations.iter(f"{PACKAGE}Relationship")
                     if item.get("Id") == sheet.get(RELATIONSHIP)), None)
    if relation is None:
        return None
    location = relation.get("Target", "")
    part = location.lstrip("/") if location.startswith("/") else \
        posixpath.normpath(str(PurePosixPath("xl") / location))
    return cell_text(package, part, f"{match.group(3)}{match.group(4)}")


def workbook_findings(data: bytes, recorded: Recorded) -> list[str]:
    """What the workbook's stamp and VBA project contradict in the manifest."""
    findings: list[str] = []
    expected_stamp = build_workbook.stamp_values(recorded.commit, "", recorded.clean)
    try:
        with zipfile.ZipFile(io.BytesIO(data)) as package:
            stamp = read_stamp(package, build_workbook.STAMP_COMMIT)
            vba_bin = package.read("xl/vbaProject.bin")
    except (KeyError, ValueError, IndexError, zipfile.BadZipFile,
            ElementTree.ParseError) as error:
        return [f"workbook: not readable: {error!r}"]
    if stamp != expected_stamp[build_workbook.STAMP_COMMIT]:
        findings.append(f"workbook: build stamp {build_workbook.STAMP_COMMIT} is {stamp!r}, "
                        f"expected {expected_stamp[build_workbook.STAMP_COMMIT]!r}")
    try:
        project = CompoundFile(vba_bin)
        project_text = project.read("PROJECT").decode(VBA_ENCODING)
        streams = set(project.children(project.children(0)["VBA"]))
        version = project.read("VBA/_VBA_PROJECT")
    except (KeyError, ValueError, UnicodeDecodeError) as error:
        return [*findings, f"workbook: VBA project not readable: {error}"]
    if version != VBA_PROJECT_STREAM:
        findings.append("workbook: _VBA_PROJECT is not version 0xFFFF without a performance "
                        "cache, so Excel would not compile the project from source")
    declared = [match.group(1, 2) for match in map(PROJECT_MODULE.match,
                                                   project_text.split("\r\n")) if match]
    expected = [("Document", name) for name in recorded.document_modules] + \
        [("Module", str(item["name"])) for item in recorded.modules]
    if declared != expected:
        findings.append(f"workbook: the VBA project declares {declared}, the manifest "
                        f"records {expected}")
    names = [name for _, name in expected]
    unexpected = sorted(streams - {*names, "dir", "_VBA_PROJECT"})
    if unexpected:
        findings.append(f"workbook: the VBA project holds streams the manifest does not "
                        f"record: {', '.join(unexpected)}")
    for item in recorded.modules:
        name = str(item["name"])
        try:
            text = decompress(project.read(f"VBA/{name}"))
        except (KeyError, ValueError) as error:
            findings.append(f"workbook: module {name} not readable: {error}")
            continue
        if build_workbook.sha256(text) != item["module_sha256"]:
            findings.append(f"workbook: module {name} has sha256 "
                            f"{build_workbook.sha256(text)}, the manifest records "
                            f"{item['module_sha256']}")
    return findings


def source_findings(root: Path, recorded: Recorded, verdict: Verdict) -> list[str]:
    """What the commit's policy and Git blobs contradict in the manifest."""
    findings: list[str] = []
    try:
        policy = json.loads(build_workbook.git_bytes(
            root, "show", f"{recorded.commit}:{build_workbook.POLICY}").decode("utf-8"))
        listed = build_workbook.regression_sources(policy)
    except build_workbook.BuildError as error:
        return [f"{build_workbook.POLICY} at {recorded.commit[:12]}: {error}"]
    except ValueError as error:
        return [f"{build_workbook.POLICY} at {recorded.commit[:12]}: not JSON: {error}"]
    paths = [str(item["path"]) for item in recorded.modules]
    for path in listed:
        if path not in paths:
            findings.append(f"{path}: in regression_sources at {recorded.commit[:12]} but "
                            "missing from the workbook")
    for path in paths:
        if path not in listed:
            findings.append(f"{path}: in the workbook but not in regression_sources at "
                            f"{recorded.commit[:12]}")
    if not findings and paths != listed:
        findings.append(f"the workbook's module order {paths} differs from regression_sources "
                        f"{listed}")
    for item in recorded.modules:
        path, name = str(item["path"]), str(item["name"])
        if Path(path).stem != name:
            findings.append(f"{path}: recorded as module {name}")
        try:
            blob = build_workbook.git_bytes(root, "cat-file", "blob", f"{recorded.commit}:{path}")
        except build_workbook.BuildError:
            if path in listed:
                findings.append(f"{path}: not in commit {recorded.commit[:12]}")
            continue
        if build_workbook.sha256(blob) != item["git_blob_sha256"]:
            findings.append(f"{path}: Git blob sha256 {build_workbook.sha256(blob)} at "
                            f"{recorded.commit[:12]}, the manifest records "
                            f"{item['git_blob_sha256']}")
        try:
            embedded = build_workbook.module_text(blob).encode(VBA_ENCODING)
        except ValueError as error:
            findings.append(f"{path} at {recorded.commit[:12]}: {error}")
            continue
        if build_workbook.sha256(embedded) == item["module_sha256"]:
            continue
        if recorded.clean:
            findings.append(f"{path}: the workbook embeds text other than the commit's "
                            f"(sha256 {item['module_sha256']}, the commit gives "
                            f"{build_workbook.sha256(embedded)})")
        else:
            verdict.notes.append(f"{path}: built from a working-tree change, not from "
                                 f"{recorded.commit[:12]}")
    return findings


def verify(root: Path, manifest: Path, workbook: Path | None = None,
           commit: str | None = None) -> Verdict:
    verdict = Verdict()
    if not manifest.is_file():
        verdict.notes.append(f"no manifest {manifest}: the workbook is not bound to a commit")
        return verdict
    try:
        recorded = load_manifest(manifest)
    except (OSError, ValueError, KeyError, TypeError) as error:
        verdict.status = FAIL
        verdict.findings.append(f"{manifest.name}: not a valid manifest: {error!r}")
        return verdict
    workbook = workbook or manifest.parent / recorded.workbook_file
    if not workbook.is_file():
        verdict.notes.append(f"no workbook {workbook}: the manifest is not checked against one")
        return verdict
    findings: list[str] = []
    if commit is not None:
        try:
            wanted = build_workbook.git(root, "rev-parse", "--verify", f"{commit}^{{commit}}")
        except build_workbook.BuildError:
            wanted = commit
        if wanted != recorded.commit:
            findings.append(f"the manifest records commit {recorded.commit}, expected {wanted}")
    data = workbook.read_bytes()
    if build_workbook.sha256(data) != recorded.workbook_sha256:
        findings.append(f"workbook {workbook.name}: sha256 {build_workbook.sha256(data)}, the "
                        f"manifest records {recorded.workbook_sha256}")
    findings += workbook_findings(data, recorded)
    try:
        build_workbook.git(root, "cat-file", "-e", f"{recorded.commit}^{{commit}}")
        known = True
    except build_workbook.BuildError:
        known = False
        verdict.notes.append(f"commit {recorded.commit} is not in {root}: the sources are "
                             "not checked")
    if known:
        findings += source_findings(root, recorded, verdict)
    verdict.findings += findings
    if findings:
        verdict.status = FAIL
    elif not known:
        verdict.status = UNVERIFIED
    elif not recorded.clean:
        verdict.status = UNVERIFIED
        verdict.notes.append("built from a working tree with changes: not evidence-grade")
    else:
        verdict.status = VERIFIED
    return verdict


def report(manifest: Path, verdict: Verdict) -> str:
    lines = [f"Manifest: {manifest}"]
    lines += [f"FAIL: {finding}" for finding in verdict.findings]
    lines += [f"note: {note}" for note in verdict.notes]
    lines.append(f"Binding: {verdict.status} (no Excel is run: this binds the workbook to its "
                 "sources; it proves no compilation and no test result)")
    return "\n".join(lines)


def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=(__doc__ or "").splitlines()[0])
    parser.add_argument("manifest", type=Path, help="the manifest written by build_workbook.py")
    parser.add_argument("--workbook", type=Path,
                        help="the workbook (default: the file the manifest names, beside it)")
    parser.add_argument("--commit", help="fail unless the manifest records this commit")
    parser.add_argument("--require-verified", action="store_true",
                        help="exit 1 unless the binding is VERIFIED")
    parser.add_argument("--root", type=Path, default=ROOT)
    args = parser.parse_args(argv)
    verdict = verify(args.root.resolve(), args.manifest, args.workbook, args.commit)
    print(report(args.manifest, verdict))
    if verdict.status == FAIL:
        return 1
    return 1 if args.require_verified and verdict.status != VERIFIED else 0


if __name__ == "__main__":
    sys.exit(main())
