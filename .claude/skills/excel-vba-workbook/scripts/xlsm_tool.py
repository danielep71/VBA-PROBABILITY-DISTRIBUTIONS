#!/usr/bin/env python3
"""Build a macro-enabled Excel workbook (.xlsm) from VBA source, without Excel.

Three commands:

  inject   put VBA modules (and optionally a Ribbon) into an existing .xlsx or
           .xlsm, keeping every other part of it byte for byte;
  build    write a plain workbook from a small JSON spec, then inject into it;
  inspect  list what a workbook holds, check it for the mistakes that make
           Excel refuse or repair a file, and optionally extract its VBA.

Run any command with -h for its options. Standard library only; vba_project.py
must sit beside this file.

The VBA project is written as compressed source with no compiled cache, so
Excel compiles it the first time the workbook opens. Nothing here runs Excel:
a successful build proves the package is well formed, not that the VBA
compiles or behaves. `inspect --check` reports what can be known without Excel.
"""
from __future__ import annotations

import argparse
import glob
import hashlib
import io
import json
import posixpath
import re
import struct
import sys
import uuid
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET
from xml.sax.saxutils import escape, quoteattr

sys.dont_write_bytecode = True      # keep the skill folder as it was installed
sys.path.insert(0, str(Path(__file__).resolve().parent))
from vba_project import (  # noqa: E402
    DOCUMENT,
    MAX_NAME,
    VBA_ENCODING,
    CompoundFile,
    Module,
    build_vba_project,
    decompress,
    document_module,
)

MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
CT = "http://schemas.openxmlformats.org/package/2006/content-types"
MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
VBA_REL = "http://schemas.microsoft.com/office/2006/relationships/vbaProject"
MACRO_CT = "application/vnd.ms-excel.sheet.macroEnabled.main+xml"
# RibbonX namespace -> (part name, relationship type)
RIBBON = {
    "http://schemas.microsoft.com/office/2009/07/customui":
        ("customUI/customUI14.xml", "http://schemas.microsoft.com/office/2007/relationships/ui/extensibility"),
    "http://schemas.microsoft.com/office/2006/01/customui":
        ("customUI/customUI.xml", "http://schemas.microsoft.com/office/2006/relationships/ui/extensibility"),
}
CALLBACK_ATTRIBUTES = ("onLoad", "onAction", "onChange", "loadImage", "getEnabled", "getVisible", "getLabel",
                       "getPressed", "getText", "getImage", "getScreentip", "getSupertip", "getDescription",
                       "getKeytip", "getSize", "getShowLabel", "getShowImage", "getItemCount", "getItemID",
                       "getItemLabel", "getItemImage", "getItemScreentip", "getItemSupertip",
                       "getSelectedItemID", "getSelectedItemIndex", "getContent", "getTitle",
                       "getItemHeight", "getItemWidth")
FONT_APPLIED, FORMAT_APPLIED = ' applyFont="1"', ' applyNumberFormat="1"'
HEADER = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
ZIP_TIME = (1980, 1, 1, 0, 0, 0)
IDENTIFIER = re.compile(r"[A-Za-z][A-Za-z0-9_]*\Z")
VB_NAME = re.compile(r'^Attribute VB_Name = "([^"]+)"\s*$', re.M)
CLASS_HEADER = re.compile(r"\AVERSION [\d.]+ CLASS\r?\nBEGIN\r?\n(?:.*\r?\n)*?END\r?\n")


class Refused(Exception):
    """A request this tool does not carry out; the message says why and what to do."""


# ----------------------------------------------------------------- VBA text ----
def crlf(text: str) -> str:
    text = text.replace("\r\n", "\n").replace("\r", "\n").replace("\n", "\r\n")
    return text if text.endswith("\r\n") else text + "\r\n"


def read_source(path: Path) -> str:
    data = path.read_bytes()
    if data.startswith(b"\xef\xbb\xbf"):
        data = data[3:]
    try:
        text = data.decode("utf-8")
    except UnicodeDecodeError:
        text = data.decode(VBA_ENCODING)
    try:
        text.encode(VBA_ENCODING)
    except UnicodeEncodeError as error:
        bad = text[error.start]
        line = text.count("\n", 0, error.start) + 1
        raise Refused(f"{path}: line {line} has {bad!r}, which Windows-1252 cannot hold. The VBA editor "
                      "stores modules in that code page: replace the character (in a string literal, "
                      "use ChrW).") from None
    return text


def check_name(name: str, what: str) -> None:
    if not IDENTIFIER.match(name) or len(name) > MAX_NAME:
        raise Refused(f"{what} {name!r}: a VBA name starts with a letter, holds letters, digits and "
                      f"underscores only, and has {MAX_NAME} characters at most")


def standard_module(path: Path) -> Module:
    text = read_source(path)
    if CLASS_HEADER.match(text) or re.search(r"^Attribute VB_(Base|PredeclaredId|Exposed)\b", text, re.M):
        raise Refused(f"{path}: this is a class or document module. Class modules and UserForms are not "
                      "supported. For the code of ThisWorkbook or of a sheet, pass the file with "
                      "--workbook-code or --sheet-code instead.")
    found = VB_NAME.search(text)
    name = found.group(1) if found else path.stem
    check_name(name, f"{path}: module name")
    if not found:
        text = f'Attribute VB_Name = "{name}"\n' + text
    elif not text.lstrip().startswith("Attribute VB_Name"):
        raise Refused(f"{path}: Attribute VB_Name must be the first line of a .bas export")
    return Module(name, crlf(text))


def document_code(path: Path) -> str:
    """The code of a document module, without any export header."""
    text = read_source(path).replace("\r\n", "\n")
    text = CLASS_HEADER.sub("", text.replace("\n", "\r\n")).replace("\r\n", "\n")
    lines = [line for line in text.split("\n") if not line.startswith("Attribute VB_")]
    while lines and not lines[0].strip():
        lines.pop(0)
    if lines and lines[0].strip().lower() == "option explicit":
        lines.pop(0)
    return "\n".join(lines)


def with_code(name: str, workbook: bool, code: str | None) -> Module:
    empty = document_module(name, workbook)
    if not code or not code.strip():
        return empty
    return Module(name, empty.source + "\r\n" + crlf(code.strip("\n")), DOCUMENT)


# ------------------------------------------------------------------ package ----
class Package:
    def __init__(self, data: bytes):
        with zipfile.ZipFile(io.BytesIO(data)) as archive:
            self.order = archive.namelist()
            self.parts = {name: archive.read(name) for name in self.order}
        if "[Content_Types].xml" not in self.parts or "xl/workbook.xml" not in self.parts:
            raise Refused("not an Excel workbook package: [Content_Types].xml or xl/workbook.xml is missing "
                          "(an .xls or .xlsb file must be saved as .xlsx first)")

    def text(self, name: str) -> str:
        return self.parts[name].decode("utf-8")

    def put(self, name: str, data: str | bytes) -> None:
        if name not in self.parts:
            self.order.append(name)
        self.parts[name] = data.encode("utf-8") if isinstance(data, str) else data

    def data(self) -> bytes:
        names = ["[Content_Types].xml"] + [name for name in self.order if name != "[Content_Types].xml"]
        out = io.BytesIO()
        with zipfile.ZipFile(out, "w") as archive:
            for name in names:
                info = zipfile.ZipInfo(name, ZIP_TIME)
                info.compress_type = zipfile.ZIP_DEFLATED
                info.create_system = 0
                info.external_attr = 0
                archive.writestr(info, self.parts[name])
        return out.getvalue()

    def sheets(self) -> list[tuple[str, str, str]]:
        """[(tab name, part name, kind)] in tab order; kind is the relationship's last word."""
        targets = {}
        for element in ET.fromstring(self.parts["xl/_rels/workbook.xml.rels"]):
            target = element.get("Target", "")
            part = target.lstrip("/") if target.startswith("/") else posixpath.normpath("xl/" + target)
            targets[element.get("Id")] = (part, element.get("Type", "").rsplit("/", 1)[-1])
        found = []
        for sheet in ET.fromstring(self.parts["xl/workbook.xml"]).iter("{%s}sheet" % MAIN):
            part, kind = targets[sheet.get("{%s}id" % REL)]
            found.append((sheet.get("name", ""), part, kind))
        return found


def opening_tag(xml: str, tag: str) -> re.Match[str] | None:
    return re.search(r"<(?:\w+:)?%s\b[^>]*?/?>" % tag, xml)


def set_code_name(xml: str, root: str, properties: str, name: str) -> str:
    """Give the sheetPr / workbookPr element the codeName, creating the element when absent."""
    found = opening_tag(xml, properties)
    if found:
        tag = found.group(0)
        if re.search(r'\bcodeName="', tag):
            new = re.sub(r'\bcodeName="[^"]*"', f'codeName="{name}"', tag)
        else:
            new = re.sub(r"(<(?:\w+:)?%s)\b" % properties, r'\1 codeName="%s"' % name, tag, count=1)
        return xml[:found.start()] + new + xml[found.end():]
    element = f'<{properties} codeName="{name}"/>'
    if properties == "workbookPr":
        # CT_Workbook order: fileVersion, fileSharing, workbookPr, ...
        for earlier in ("fileSharing", "fileVersion"):
            before = opening_tag(xml, earlier)
            if before:
                end = before.end() if before.group(0).endswith("/>") else xml.index(">", xml.index(
                    "</", before.end())) + 1
                return xml[:end] + element + xml[end:]
    start = opening_tag(xml, root)
    if start is None:
        raise Refused(f"no <{root}> element found")
    return xml[:start.end()] + element + xml[start.end():]


def current_code_name(xml: str, properties: str) -> str | None:
    found = opening_tag(xml, properties)
    name = re.search(r'\bcodeName="([^"]*)"', found.group(0)) if found else None
    return name.group(1) if name else None


def add_override(types: str, part: str, content_type: str) -> str:
    types = re.sub(r'<Override PartName="/%s"[^>]*/>' % re.escape(part), "", types)
    return types.replace("</Types>", f'<Override PartName="/{part}" ContentType="{content_type}"/></Types>')


def add_default(types: str, extension: str, content_type: str) -> str:
    if re.search(r'<Default Extension="%s"' % extension, types, re.I):
        return types
    return re.sub(r"(<Types\b[^>]*>)", r'\1<Default Extension="%s" ContentType="%s"/>' % (
        extension, content_type), types, count=1)


def add_relationship(rels: str, rel_type: str, target: str, wanted: str) -> str:
    rels = re.sub(r'<Relationship\b[^>]*Type="%s"[^>]*/>' % re.escape(rel_type), "", rels)
    used = set(re.findall(r'\bId="([^"]+)"', rels))
    rid, number = wanted, 1
    while rid in used:
        number += 1
        rid = f"{wanted}{number}"
    return rels.replace("</Relationships>",
                        f'<Relationship Id="{rid}" Type="{rel_type}" Target="{target}"/></Relationships>')


def ribbon_part(xml: bytes) -> tuple[str, str]:
    try:
        root = ET.fromstring(xml)
    except ET.ParseError as error:
        raise Refused(f"the Ribbon file is not well-formed XML: {error}") from None
    namespace = root.tag[1:].split("}")[0] if root.tag.startswith("{") else ""
    if namespace not in RIBBON or not root.tag.endswith("}customUI"):
        raise Refused("the Ribbon file's root must be <customUI> in the namespace "
                      "http://schemas.microsoft.com/office/2009/07/customui (Office 2010 and later)")
    return RIBBON[namespace]


def ribbon_callbacks(xml: bytes) -> set[str]:
    return {element.get(attribute) for element in ET.fromstring(xml).iter()
            for attribute in CALLBACK_ATTRIBUTES if element.get(attribute)} - {None}  # type: ignore[arg-type]


def inject(data: bytes, modules: list[Module], workbook_code: str | None, sheet_code: dict[str, str],
           code_names: dict[str, str], ribbon: bytes | None, project_id: str | None,
           keep_vba: bool = False) -> tuple[bytes, list[str]]:
    """The workbook with the VBA project (and Ribbon) put in; and notes for the user."""
    package = Package(data)
    notes: list[str] = []
    if "xl/vbaProject.bin" in package.parts and not keep_vba:
        notes.append("the workbook already had a VBA project: it is replaced, not merged")
    sheets = package.sheets()
    other = [name for name, _, kind in sheets if kind != "worksheet"]
    if other:
        raise Refused(f"the workbook has chart or macro sheets ({', '.join(other)}); only worksheets are "
                      "supported. Move the charts onto worksheets.")
    unknown = (set(code_names) | set(sheet_code)) - {name for name, _, _ in sheets}
    if unknown:
        raise Refused(f"no sheet named {sorted(unknown)}; the sheets are {[name for name, _, _ in sheets]}")

    workbook = set_code_name(package.text("xl/workbook.xml"), "workbook", "workbookPr", "ThisWorkbook")
    package.put("xl/workbook.xml", workbook)
    documents = [with_code("ThisWorkbook", True, workbook_code)]
    taken = {"thisworkbook"} | {module.name.lower() for module in modules}
    for number, (name, part, _) in enumerate(sheets, 1):
        xml = package.text(part)
        code = code_names.get(name) or current_code_name(xml, "sheetPr") or f"Sheet{number}"
        check_name(code, f"sheet {name!r}: CodeName")
        if code.lower() in taken:
            raise Refused(f"the name {code} is used twice (sheet CodeNames, module names and ThisWorkbook "
                          "share one namespace)")
        taken.add(code.lower())
        package.put(part, set_code_name(xml, "worksheet", "sheetPr", code))
        documents.append(with_code(code, False, sheet_code.get(name)))

    identifier = project_id or "{%s}" % str(uuid.UUID(hashlib.sha256(
        b"".join(module.source.encode(VBA_ENCODING) for module in documents + modules)).hexdigest()[:32])).upper()
    package.put("xl/vbaProject.bin", build_vba_project(documents + modules, identifier))

    types = package.text("[Content_Types].xml")
    types = add_default(types, "bin", "application/vnd.ms-office.vbaProject")
    types = add_default(types, "xml", "application/xml")
    types = add_override(types, "xl/workbook.xml", MACRO_CT)
    package.put("[Content_Types].xml", types)
    package.put("xl/_rels/workbook.xml.rels", add_relationship(
        package.text("xl/_rels/workbook.xml.rels"), VBA_REL, "vbaProject.bin", "rIdVbaProject"))

    if ribbon is not None:
        part, rel_type = ribbon_part(ribbon)
        for old, _ in RIBBON.values():
            package.parts.pop(old, None)
            package.order = [name for name in package.order if name != old]
        package.put(part, ribbon)
        root = package.text("_rels/.rels")
        for _, old_type in RIBBON.values():
            root = re.sub(r'<Relationship\b[^>]*Type="%s"[^>]*/>' % re.escape(old_type), "", root)
        package.put("_rels/.rels", add_relationship(root, rel_type, part, "rIdRibbon"))
    notes.append("VBA not compiled: Excel compiles it when the workbook opens. Open it in Excel and run "
                 "Debug > Compile VBAProject before relying on it.")
    return package.data(), notes


# ---------------------------------------------------------- plain workbook ----
def column_number(letters: str) -> int:
    number = 0
    for letter in letters:
        number = number * 26 + ord(letter) - 64
    return number


def sheet_xml(cells: dict[str, object], widths: dict[str, float], style_of: dict[str, int]) -> str:
    rows: dict[int, list[tuple[int, str]]] = {}
    for reference, value in cells.items():
        found = re.fullmatch(r"([A-Z]{1,3})([1-9]\d*)", reference)
        if not found:
            raise Refused(f"cell reference {reference!r}: use the form A1")
        style = f' s="{style_of[reference]}"' if style_of.get(reference) else ""
        if isinstance(value, bool):
            body = f'<c r="{reference}"{style} t="b"><v>{int(value)}</v></c>'
        elif isinstance(value, (int, float)):
            body = f'<c r="{reference}"{style}><v>{value!r}</v></c>'
        elif isinstance(value, str) and value.startswith("=") and len(value) > 1:
            body = f'<c r="{reference}"{style}><f>{escape(value[1:])}</f></c>'
        else:
            body = (f'<c r="{reference}"{style} t="inlineStr"><is><t xml:space="preserve">'
                    f'{escape(str(value))}</t></is></c>')
        rows.setdefault(int(found.group(2)), []).append((column_number(found.group(1)), body))
    data = "".join(f'<row r="{row}">{"".join(body for _, body in sorted(items))}</row>'
                   for row, items in sorted(rows.items()))
    cols = ""
    if widths:
        cols = "<cols>" + "".join(
            f'<col min="{column_number(letters)}" max="{column_number(letters)}" width="{width}" customWidth="1"/>'
            for letters, width in sorted(widths.items(), key=lambda item: column_number(item[0]))) + "</cols>"
    return f'{HEADER}<worksheet xmlns="{MAIN}" xmlns:r="{REL}">{cols}<sheetData>{data}</sheetData></worksheet>'


def plain_workbook(sheets: list[dict[str, object]], names: dict[str, str]) -> bytes:
    """A minimal .xlsx: one part per sheet, two cell styles (normal and bold)."""
    if not sheets:
        raise Refused("a workbook needs one sheet at least")
    seen: set[str] = set()
    for sheet in sheets:
        name = str(sheet.get("name", ""))
        if not name or len(name) > 31 or re.search(r"[\[\]:*?/\\]", name) or name.lower() in seen:
            raise Refused(f"sheet name {name!r}: 1 to 31 characters, none of [ ] : * ? / \\, and unique")
        seen.add(name.lower())
    # One cell style per (bold, number format) pair in use; style 0 is the default.
    styles: list[tuple[bool, str]] = [(False, "")]
    style_maps: list[dict[str, int]] = []
    for sheet in sheets:
        bold = set(sheet.get("bold", []))  # type: ignore[call-overload]
        formats = dict(sheet.get("formats", {}))  # type: ignore[call-overload]
        mapping: dict[str, int] = {}
        for reference in sorted(bold | set(formats)):
            key = (reference in bold, str(formats.get(reference, "")))
            if key not in styles:
                styles.append(key)
            mapping[reference] = styles.index(key)
        style_maps.append(mapping)
    codes = sorted({code for _, code in styles if code})
    number_formats = "".join(f'<numFmt numFmtId="{164 + index}" formatCode={quoteattr(code)}/>'
                             for index, code in enumerate(codes))
    cell_formats = "".join(
        f'<xf numFmtId="{164 + codes.index(code) if code else 0}" fontId="{1 if bold else 0}" fillId="0" '
        f'borderId="0" xfId="0"{FONT_APPLIED if bold else ""}{FORMAT_APPLIED if code else ""}/>'
        for bold, code in styles)
    parts = {
        "[Content_Types].xml": f'{HEADER}<Types xmlns="{CT}">'
                               '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
                               '<Default Extension="xml" ContentType="application/xml"/>'
                               '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
                               '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
                               + "".join(f'<Override PartName="/xl/worksheets/sheet{n}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
                                         for n in range(1, len(sheets) + 1)) + "</Types>",
        "_rels/.rels": f'{HEADER}<Relationships xmlns="{PKG_REL}"><Relationship Id="rId1" '
                       f'Type="{REL}/officeDocument" Target="xl/workbook.xml"/></Relationships>',
        "xl/workbook.xml": f'{HEADER}<workbook xmlns="{MAIN}" xmlns:r="{REL}"><workbookPr/><sheets>'
                           + "".join(f'<sheet name={quoteattr(str(sheet["name"]))} sheetId="{n}" r:id="rId{n}"/>'
                                     for n, sheet in enumerate(sheets, 1)) + "</sheets>"
                           + ("<definedNames>" + "".join(
                               f"<definedName name={quoteattr(name)}>{escape(target)}</definedName>"
                               for name, target in sorted(names.items())) + "</definedNames>" if names else "")
                           + "</workbook>",
        "xl/_rels/workbook.xml.rels": f'{HEADER}<Relationships xmlns="{PKG_REL}">'
                                      + "".join(f'<Relationship Id="rId{n}" Type="{REL}/worksheet" '
                                                f'Target="worksheets/sheet{n}.xml"/>' for n in range(1, len(sheets) + 1))
                                      + f'<Relationship Id="rIdStyles" Type="{REL}/styles" Target="styles.xml"/>'
                                        "</Relationships>",
        "xl/styles.xml": f'{HEADER}<styleSheet xmlns="{MAIN}">'
                         + (f'<numFmts count="{len(codes)}">{number_formats}</numFmts>' if codes else "")
                         + '<fonts count="2"><font><sz val="11"/>'
                         '<name val="Calibri"/></font><font><b/><sz val="11"/><name val="Calibri"/></font></fonts>'
                         '<fills count="2"><fill><patternFill patternType="none"/></fill><fill>'
                         '<patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/>'
                         '<right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1">'
                         '<xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
                         + f'<cellXfs count="{len(styles)}">{cell_formats}</cellXfs>'
                         '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
                         "</styleSheet>",
    }
    for number, sheet in enumerate(sheets, 1):
        parts[f"xl/worksheets/sheet{number}.xml"] = sheet_xml(
            dict(sheet.get("cells", {})), dict(sheet.get("widths", {})), style_maps[number - 1])  # type: ignore[call-overload]
    out = io.BytesIO()
    with zipfile.ZipFile(out, "w") as archive:
        for name, text in parts.items():
            info = zipfile.ZipInfo(name, ZIP_TIME)
            info.compress_type = zipfile.ZIP_DEFLATED
            info.create_system = 0
            info.external_attr = 0
            archive.writestr(info, text.encode("utf-8"))
    return out.getvalue()


# ------------------------------------------------------------------ inspect ----
def dir_modules(directory: bytes) -> list[tuple[str, str, int, str]]:
    """[(module name, stream name, source offset, kind)] from an uncompressed dir stream."""
    position, found = 0, []
    name = stream = ""
    offset, kind, in_modules = 0, "standard", False
    while position + 6 <= len(directory):
        record, size = struct.unpack_from("<HI", directory, position)
        if record == 0x0009:                       # PROJECTVERSION: fixed 6 bytes after a 4 in the size field
            position += 12
            continue
        payload = directory[position + 6:position + 6 + size]
        position += 6 + size
        if record == 0x000F:
            in_modules = True
        elif not in_modules:
            continue
        elif record == 0x0019:
            name = stream = payload.decode(VBA_ENCODING, "replace")
            offset, kind = 0, "standard"
        elif record == 0x0047:
            name = payload.decode("utf-16-le", "replace")
        elif record == 0x001A:
            stream = payload.decode(VBA_ENCODING, "replace")
        elif record == 0x0032:
            stream = payload.decode("utf-16-le", "replace")
        elif record == 0x0031:
            offset = struct.unpack("<I", payload)[0]
        elif record == 0x0022:
            kind = "document or class"
        elif record == 0x002B:
            found.append((name, stream, offset, kind))
        elif record == 0x0010:
            break
    return found


def vba_sources(blob: bytes) -> list[tuple[str, str, str]]:
    """[(module name, kind, source text)] of a vbaProject.bin."""
    project = CompoundFile(blob)
    found = []
    for name, stream, offset, kind in dir_modules(decompress(project.read("VBA/dir"))):
        source = decompress(project.read("VBA/" + stream)[offset:]).decode(VBA_ENCODING, "replace")
        found.append((name, kind, source))
    return found


def lint(source: str) -> list[str]:
    """What can be said about a module's text without a compiler: far from a compile."""
    code = [re.sub(r'"[^"]*"', '""', line).split("'")[0].strip() for line in source.replace("\r\n", "\n").split("\n")]
    code = [line for line in code if line and not line.startswith("Attribute ")]
    findings = []
    if code and not any(line.lower() == "option explicit" for line in code):
        findings.append("no Option Explicit: a misspelt variable becomes a new empty one")
    opened = re.compile(r"^(?:(?:Public|Private|Friend)\s+)?(?:Static\s+)?(Sub|Function|Property)\b", re.I)
    current = None
    for line in code:
        start = opened.match(line)
        end = re.match(r"^End\s+(Sub|Function|Property)\b", line, re.I)
        if start and not re.match(r"^(?:\w+\s+)*Declare\b", line, re.I):
            if current:
                findings.append(f"a {start.group(1)} starts before the previous {current} ends")
            current = start.group(1).capitalize()
        elif end:
            if not current or current.lower() != end.group(1).lower():
                findings.append(f"End {end.group(1)} without a matching start")
            current = None
    if current:
        findings.append(f"the last {current} has no End {current}")
    return findings


def undeclared_prefixes(xml: bytes) -> list[str]:
    """Prefixes named by mc:Ignorable / mc:MustUnderstand that the element does not have in scope."""
    scopes: list[dict[str, str]] = [{}]
    pending: dict[str, str] = {}
    bad: list[str] = []
    for event, node in ET.iterparse(io.BytesIO(xml), events=("start-ns", "start", "end")):
        if event == "start-ns":
            pending[node[0]] = node[1]
        elif event == "start":
            scope = dict(scopes[-1])
            scope.update(pending)
            pending = {}
            scopes.append(scope)
            for attribute in ("Ignorable", "MustUnderstand"):
                bad += [prefix for prefix in node.get("{%s}%s" % (MC, attribute), "").split()
                        if prefix not in scope]
        else:
            scopes.pop()
    return bad


def inspect(data: bytes, extract: Path | None) -> tuple[list[str], list[str]]:
    """(report lines, problems). Problems are things Excel would refuse, repair or report."""
    package = Package(data)
    report, problems = [], []
    for name in package.order:
        if name.endswith((".xml", ".rels", ".vml")):
            try:
                ET.fromstring(package.parts[name])
            except ET.ParseError as error:
                problems.append(f"{name}: not well-formed XML ({error})")
                continue
            bad = undeclared_prefixes(package.parts[name]) if name.endswith(".xml") else []
            if bad:
                problems.append(f"{name}: mc:Ignorable names undeclared prefixes {sorted(set(bad))}; Excel "
                                "reports the file as corrupt")
    types = package.text("[Content_Types].xml")
    macro = MACRO_CT in types
    has_vba = "xl/vbaProject.bin" in package.parts
    report.append(f"macro-enabled content type: {'yes' if macro else 'no'}; VBA project: "
                  f"{'yes' if has_vba else 'no'}")
    if has_vba != macro:
        problems.append("the workbook content type and the presence of xl/vbaProject.bin disagree; Excel "
                        "refuses the file or drops the macros")
    code_names = {"ThisWorkbook": "workbook"}
    workbook_code = current_code_name(package.text("xl/workbook.xml"), "workbookPr")
    for name, part, kind in package.sheets():
        code = current_code_name(package.text(part), "sheetPr") if kind == "worksheet" else None
        report.append(f"sheet {name!r}: {part}, CodeName {code or '(none)'}")
        if code:
            code_names[code] = name
    modules: list[tuple[str, str, str]] = []
    if has_vba:
        try:
            modules = vba_sources(package.parts["xl/vbaProject.bin"])
        except (KeyError, ValueError, struct.error) as error:
            problems.append(f"xl/vbaProject.bin could not be read by this tool: {error}")
        for name, kind, source in modules:
            lines = source.count("\n")
            report.append(f"module {name}: {kind}, {lines} lines")
            problems += [f"module {name}: {finding}" for finding in lint(source)]
        documents = {name for name, kind, _ in modules if kind != "standard"}
        if modules and workbook_code != "ThisWorkbook" and "ThisWorkbook" in documents:
            problems.append(f"workbookPr codeName is {workbook_code!r}; the project's workbook module is "
                            "ThisWorkbook, so Excel will not bind its events")
        for code, sheet in code_names.items():
            if modules and code not in documents:
                problems.append(f"no document module named {code} for {sheet!r}: Excel creates a new empty "
                                "one and may rename it")
    ribbons = [part for part, _ in RIBBON.values() if part in package.parts]
    for part in ribbons:
        callbacks = ribbon_callbacks(package.parts[part])
        report.append(f"Ribbon {part}: {len(callbacks)} callback name(s)")
        if part not in package.text("_rels/.rels"):
            problems.append(f"{part} has no relationship in _rels/.rels: Excel ignores the Ribbon")
        procedures = {name.lower() for _, kind, source in modules if kind == "standard" for name in re.findall(
            r"^[ \t]*(?:Public[ \t]+)?(?:Sub|Function)[ \t]+([A-Za-z_]\w*)", source, re.M | re.I)}
        missing = sorted(name for name in callbacks if name.split(".")[-1].lower() not in procedures)
        if missing:
            problems.append(f"Ribbon callbacks without a procedure in a standard module: {', '.join(missing)}")
    if extract is not None:
        extract.mkdir(parents=True, exist_ok=True)
        for name, kind, source in modules:
            suffix = ".bas" if kind == "standard" else ".cls"
            (extract / (name + suffix)).write_bytes(source.encode(VBA_ENCODING, "replace"))
        report.append(f"extracted {len(modules)} module(s) to {extract}")
    return report, problems


# ---------------------------------------------------------------------- CLI ----
def expand(patterns: list[str]) -> list[Path]:
    found: list[Path] = []
    for pattern in patterns:
        matches = sorted(glob.glob(pattern, recursive=True))
        if not matches:
            raise Refused(f"no file matches {pattern}")
        found += [Path(match) for match in matches]
    return found


def pairs(items: list[str], what: str) -> dict[str, str]:
    found = {}
    for item in items:
        if "=" not in item:
            raise Refused(f"{what} {item!r}: write it as SheetName=value")
        key, value = item.split("=", 1)
        found[key] = value
    return found


def finish(data: bytes, output: Path, notes: list[str]) -> int:
    if output.suffix.lower() not in (".xlsm", ".xltm", ".xlam"):
        raise Refused(f"{output}: a workbook with VBA needs the extension .xlsm; Excel refuses macros in "
                      "a file named .xlsx")
    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_bytes(data)
    report, problems = inspect(data, None)
    print(f"wrote {output} ({len(data)} bytes, sha256 {hashlib.sha256(data).hexdigest()[:16]}...)")
    for line in report:
        print("  " + line)
    for note in notes:
        print("NOTE: " + note)
    for problem in problems:
        print("PROBLEM: " + problem)
    return 1 if problems else 0


def command_inject(arguments: argparse.Namespace) -> int:
    modules = [standard_module(path) for path in expand(arguments.modules)]
    workbook_code = document_code(Path(arguments.workbook_code)) if arguments.workbook_code else None
    sheet_code = {sheet: document_code(Path(path)) for sheet, path in pairs(arguments.sheet_code, "--sheet-code").items()}
    ribbon = Path(arguments.ribbon).read_bytes() if arguments.ribbon else None
    data, notes = inject(Path(arguments.workbook).read_bytes(), modules, workbook_code, sheet_code,
                         pairs(arguments.code_name, "--code-name"), ribbon, arguments.project_id)
    return finish(data, Path(arguments.output), notes)


def command_build(arguments: argparse.Namespace) -> int:
    spec_path = Path(arguments.spec)
    spec = json.loads(spec_path.read_text(encoding="utf-8"))
    base = spec_path.parent

    def resolve(value: str) -> str:
        return str(base / value)

    sheets = spec.get("sheets") or [{"name": "Sheet1"}]
    modules = [standard_module(path) for path in expand([resolve(item) for item in spec.get("modules", [])])]
    workbook_code = document_code(Path(resolve(spec["workbook_code"]))) if spec.get("workbook_code") else None
    sheet_code = {str(sheet["name"]): document_code(Path(resolve(sheet["code"]))) for sheet in sheets
                  if sheet.get("code")}
    code_names = {str(sheet["name"]): str(sheet["code_name"]) for sheet in sheets if sheet.get("code_name")}
    ribbon = Path(resolve(spec["ribbon"])).read_bytes() if spec.get("ribbon") else None
    plain = plain_workbook(sheets, dict(spec.get("defined_names", {})))
    data, notes = inject(plain, modules, workbook_code, sheet_code, code_names, ribbon, spec.get("project_id"))
    return finish(data, Path(arguments.output), notes)


def command_inspect(arguments: argparse.Namespace) -> int:
    report, problems = inspect(Path(arguments.workbook).read_bytes(),
                               Path(arguments.extract) if arguments.extract else None)
    for line in report:
        print(line)
    for problem in problems:
        print("PROBLEM: " + problem)
    if arguments.check:
        print("check: " + ("FAIL" if problems else "PASS (package only; nothing was compiled or run)"))
        return 1 if problems else 0
    return 0


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    commands = parser.add_subparsers(dest="command", required=True)

    one = commands.add_parser("inject", help="put VBA into an existing .xlsx or .xlsm")
    one.add_argument("workbook", help="the workbook to start from; it is not changed")
    one.add_argument("-o", "--output", required=True, help="the .xlsm to write")
    one.add_argument("-m", "--modules", nargs="+", default=[], metavar="BAS",
                     help="standard modules (.bas files or glob patterns), in project order")
    one.add_argument("--workbook-code", metavar="FILE", help="code of ThisWorkbook (events)")
    one.add_argument("--sheet-code", action="append", default=[], metavar="SHEET=FILE",
                     help="code of one sheet's module; repeat for more sheets")
    one.add_argument("--code-name", action="append", default=[], metavar="SHEET=NAME",
                     help="CodeName to give a sheet; repeat. Default: its current one, or SheetN")
    one.add_argument("--ribbon", metavar="XML", help="RibbonX file (customUI)")
    one.add_argument("--project-id", metavar="GUID", help="project GUID in braces; default: derived from the code")
    one.set_defaults(run=command_inject)

    two = commands.add_parser("build", help="write a plain workbook from a JSON spec, with its VBA")
    two.add_argument("spec", help="JSON file; see references/spec.md")
    two.add_argument("-o", "--output", required=True, help="the .xlsm to write")
    two.set_defaults(run=command_build)

    three = commands.add_parser("inspect", help="list, check and extract")
    three.add_argument("workbook")
    three.add_argument("--check", action="store_true", help="exit 1 when a problem is found")
    three.add_argument("--extract", metavar="DIR", help="write every module's source into DIR")
    three.set_defaults(run=command_inspect)

    arguments = parser.parse_args(argv)
    try:
        return int(arguments.run(arguments))
    except Refused as error:
        print(f"refused: {error}", file=sys.stderr)
        return 2
    except (OSError, ValueError, KeyError, json.JSONDecodeError, zipfile.BadZipFile) as error:
        print(f"error: {error}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
