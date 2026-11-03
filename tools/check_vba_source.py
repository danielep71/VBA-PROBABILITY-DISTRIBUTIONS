#!/usr/bin/env python3
"""Check the placement, names, encoding, options, storage and layout of VBA source.

For every tracked .bas, .cls and .frm file: it sits in src/ or tests/ (flat)
or under benchmark/; its file name starts with M_STATS_PROBDIST_, has at most
31 characters and equals its one Attribute VB_Name, unique ignoring case;
ASCII only; Option Explicit; Option Private Module in the modules of
OPTION_PRIVATE_REQUIRED and never in a module that declares a Public K_STATS_
worksheet function; no Option Base, tab or trailing white space; string
literals closed; LF storage in Git with CRLF checkout (.gitattributes).
tools/README.md, "VBA source checks", gives the reasons.
"""

from __future__ import annotations

import argparse
import re
import subprocess
import sys
from collections.abc import Sequence
from pathlib import Path

from vba_lexer import (
    Finding,
    read_component,
    report,
    statements,
    tracked_components,
)

ROOT = Path(__file__).resolve().parents[1]
# Where components live: the production and test modules flat in src/ and
# tests/, the exporters and probes of each study anywhere under benchmark/.
FLAT_HOMES = ("src/", "tests/")
NESTED_HOMES = ("benchmark/",)
NAME_PREFIX = "M_STATS_PROBDIST_"
NAME = re.compile(rf"{NAME_PREFIX}[A-Za-z0-9_]+")
MAX_NAME = 31  # the VBE rejects a longer component name
# The internal numerical layers. Option Private Module keeps their Public
# PROB_* routines inside the VBA project and out of the worksheet Function
# Wizard (README, "Layer 1" and "Layer 2"; docs/PUBLIC_API.md).
OPTION_PRIVATE_REQUIRED = frozenset({
    "src/M_STATS_PROBDIST_CORE.bas",
    "src/M_STATS_PROBDIST_SPECIALFUNCS.bas",
})
# A module that declares a Public K_STATS_ procedure exposes worksheet
# functions; Option Private Module would hide them, so it is forbidden there.
WORKSHEET_FUNCTION = re.compile(
    r"Public\s+(?:Static\s+)?(?:Function|Sub|Property\s+(?:Get|Let|Set))\s+K_STATS_\w*", re.I)

VB_NAME = re.compile(r'^Attribute VB_Name = "([^"]*)"[ \t]*$', re.M)
# Option statements, matched against comment-free statements.
OPTION_EXPLICIT = re.compile(r"Option\s+Explicit", re.I)
OPTION_PRIVATE = re.compile(r"Option\s+Private\s+Module", re.I)
# Explicit array bounds, never Option Base.
OPTION_BASE = re.compile(r"Option\s+Base\b", re.I)
REM_STATEMENT = re.compile(r"Rem(?=[ \t]|$)", re.I)
LINE_NUMBER = re.compile(r"[ \t]*\d+(?=[ \t])")


def placement(path: str) -> list[Finding]:
    """A component sits in a home and is named M_STATS_PROBDIST_<Subject>."""
    findings: list[Finding] = []
    flat = next((home for home in FLAT_HOMES if path.startswith(home)), None)
    if (flat is None or "/" in path[len(flat):]) and not path.startswith(NESTED_HOMES):
        findings.append(Finding(path, None, "VBA component outside src/ and tests/ (no "
                                            "subdirectories) and benchmark/"))
    name = Path(path).stem
    if not NAME.fullmatch(name):
        findings.append(Finding(path, None, f"component name must start with {NAME_PREFIX} "
                                            "and use letters, digits and underscores only"))
    if len(name) > MAX_NAME:
        findings.append(Finding(path, None, f"component name longer than {MAX_NAME} characters"))
    return findings


def non_ascii(path: str, data: bytes) -> list[Finding]:
    if data.startswith(b"\xef\xbb\xbf"):
        return [Finding(path, 1, "byte-order mark; VBA source is ASCII only")]
    return [Finding(path, number, f"non-ASCII byte 0x{byte:02X}; VBA source is ASCII only, "
                                  "build the character with ChrW")
            for number, line in enumerate(data.split(b"\n"), 1)
            for byte in sorted(set(line)) if byte > 0x7F][:20]


def unclosed_literals(text: str) -> list[int]:
    """Lines whose code leaves a string literal open.

    A ' outside a literal and outside [...] (which quotes a sheet name, as in
    ['Sheet 1'!A1]), or Rem at the start of a statement, also after a line
    number, begins a comment; a comment ending in " _" continues on the next
    line.
    """
    open_lines: list[int] = []
    in_comment = False
    for number, line in enumerate(text.split("\n"), 1):
        if in_comment:
            in_comment = line.rstrip().endswith(" _")
            continue
        quoted, statement_start, brackets = False, True, 0
        label = LINE_NUMBER.match(line)
        for index in range(label.end() if label else 0, len(line)):
            char = line[index]
            if quoted:
                quoted = char != '"'
            elif char == '"':
                quoted = True
            elif (char == "'" and not brackets) or \
                    (statement_start and REM_STATEMENT.match(line, index)):
                in_comment = line.rstrip().endswith(" _")
                break
            elif char == "[":
                brackets += 1
            elif char == "]" and brackets:
                brackets -= 1
            statement_start = (char == ":" and not quoted) or \
                (statement_start and char.isspace())
        if quoted:
            open_lines.append(number)
    return open_lines


def layout(path: str, text: str) -> list[Finding]:
    """No tab, no trailing white space."""
    findings = []
    for number, line in enumerate(text.split("\n"), 1):
        if "\t" in line:
            findings.append(Finding(path, number, "tab character; indent with spaces"))
        if line != line.rstrip(" \t"):
            findings.append(Finding(path, number, "trailing white space"))
    return findings


def options(path: str, text: str, required: frozenset[str]) -> list[Finding]:
    """Option Explicit everywhere; Option Private Module where the design needs
    it and never beside worksheet functions; no Option Base."""
    # Comments, continued ones included, are not statements.
    items = statements(text)
    findings: list[Finding] = []
    if not any(OPTION_EXPLICIT.fullmatch(statement) for _, statement in items):
        findings.append(Finding(path, None, "missing Option Explicit"))
    private = [line for line, statement in items if OPTION_PRIVATE.fullmatch(statement)]
    if path in required and not private:
        findings.append(Finding(path, None, "missing Option Private Module; this internal "
                                            "module keeps its Public routines project-scoped"))
    worksheet = next((line for line, statement in items if WORKSHEET_FUNCTION.match(statement)),
                     None)
    if private and worksheet is not None:
        findings.append(Finding(path, private[0], "Option Private Module in a module that "
                                                  "declares Public K_STATS_ worksheet functions "
                                                  f"(line {worksheet}); it would hide them"))
    findings.extend(Finding(path, line, "Option Base; declare explicit array bounds")
                    for line, statement in items if OPTION_BASE.match(statement))
    return findings


def content(root: Path, path: str, names: dict[str, str],
            required: frozenset[str]) -> list[Finding]:
    data = (root / path).read_bytes()
    text = read_component(root, path)
    findings = non_ascii(path, data)
    vb_names = VB_NAME.findall(text)
    stem = Path(path).stem
    if vb_names != [stem]:
        findings.append(Finding(path, None, f'needs exactly one Attribute VB_Name = "{stem}"'))
    elif stem.casefold() in names:
        findings.append(Finding(path, None, f"component name also used by {names[stem.casefold()]}"
                                            " (VBA names are case-insensitive)"))
    else:
        names[stem.casefold()] = path
    if Path(path).suffix.lower() == ".bas" and not text.startswith('Attribute VB_Name = "'):
        findings.append(Finding(path, 1, "a .bas export starts with its Attribute VB_Name line"))
    findings.extend(options(path, text, required))
    findings.extend(Finding(path, number, "string literal not closed")
                    for number in unclosed_literals(text))
    findings.extend(layout(path, text))
    return findings


def storage(root: Path, paths: list[str]) -> list[Finding]:
    """Exported VBA is stored with LF and checked out with CRLF (.gitattributes)."""
    if not paths:
        return []
    completed = subprocess.run(["git", "-C", str(root), "ls-files", "--eol", "-z", "--", *paths],
                               capture_output=True, check=False)
    if completed.returncode:
        raise RuntimeError(completed.stderr.decode("utf-8", errors="replace").strip())
    findings = []
    for record in completed.stdout.split(b"\0"):
        if not record:
            continue
        info, _, path = record.decode("utf-8", errors="surrogateescape").partition("\t")
        fields = info.split()
        index_eol, attributes = fields[0].removeprefix("i/"), " ".join(fields[2:])
        if index_eol not in {"lf", "none", ""}:
            findings.append(Finding(path, None, f"stored with {index_eol} in Git; store with LF"))
        if "eol=crlf" not in attributes:
            findings.append(Finding(path, None, ".gitattributes must check it out with CRLF"))
    return findings


def check(root: Path, required: frozenset[str] = OPTION_PRIVATE_REQUIRED
          ) -> tuple[list[Finding], int]:
    components = tracked_components(root)
    findings: list[Finding] = []
    # A renamed or removed internal module must not silently drop its rule.
    findings.extend(Finding(path, None, "listed in OPTION_PRIVATE_REQUIRED but not a tracked "
                                        "VBA component")
                    for path in sorted(required - set(components)))
    names: dict[str, str] = {}
    for path in components:
        findings.extend(placement(path))
        findings.extend(content(root, path, names, required))
    findings.extend(storage(root, components))
    return findings, len(components)


def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=(__doc__ or "").splitlines()[0])
    parser.add_argument("--root", type=Path, default=ROOT)
    args = parser.parse_args(argv)
    findings, components = check(args.root.resolve())
    return report("vba-source", findings, f"{components} VBA component(s)")


if __name__ == "__main__":
    sys.exit(main())
