#!/usr/bin/env python3
"""Check explicit VBA visibility and public-name collisions between standard modules.

Every procedure and module-level declaration states its visibility: Public
or Private (Global counts as Public), and Friend only in a class module. No
public name (procedure, Const, Enum, Type, variable or Declare) is defined
twice in a standard module or in two standard modules of one collision group,
which Excel reports as an ambiguous name. Each compilation environment of
check_vba_conditionals.py is read separately, so that mutually exclusive #If
branches are never combined.

benchmark/check_public_api.py freezes the K_STATS_ worksheet API; this check
keeps no inventory of its own.
"""

from __future__ import annotations

import argparse
import re
import sys
from collections.abc import Sequence
from dataclasses import dataclass, field
from pathlib import Path

from check_vba_conditionals import reachable_sources
from vba_lexer import Finding, read_component, report, statements, tracked_components

ROOT = Path(__file__).resolve().parents[1]
# Standard modules that can be compiled into one VBA project. I keep a single
# group: the Excel regression imports src/ and tests/ into one workbook
# (.github/excel-evidence-policy.json, regression_sources), and every exporter
# and probe under benchmark/ is imported into a workbook that already holds
# the production modules, where the main-grid evidence binds production, test,
# exporter and study modules together (benchmark/PROVENANCE.md). One group is
# the superset of every import set, so no combination can be ambiguous.
COLLISION_GROUPS: dict[str, tuple[str, ...]] = {
    "one VBA project": ("src/", "tests/", "benchmark/"),
}
PROPERTY_KINDS = frozenset({"property get", "property let", "property set"})

PROC = re.compile(
    r"^\s*(?:(?P<vis>Public|Private|Friend|Global)\s+)?(?:Static\s+)?"
    r"(?P<kind>Sub|Function|Property\s+(?:Get|Let|Set))\s+(?P<name>[A-Za-z_]\w*)\b", re.I)
END_PROC = re.compile(r"^\s*End\s+(?:Sub|Function|Property)\b", re.I)
BLOCK = re.compile(r"^\s*(?:(?P<vis>Public|Private|Global)\s+)?(?P<kind>Enum|Type)\s+"
                   r"(?P<name>[A-Za-z_]\w*)\b", re.I)
END_BLOCK = {"enum": re.compile(r"^\s*End\s+Enum\b", re.I),
             "type": re.compile(r"^\s*End\s+Type\b", re.I)}
VISIBLE = re.compile(r"^\s*(?P<vis>Public|Private|Global|Friend)\s+(?P<rest>.*)$", re.I)
DECLARE = re.compile(r"^Declare\s+(?:PtrSafe\s+)?(?P<kind>Function|Sub)\s+"
                     r"(?P<name>[A-Za-z_]\w*)\b", re.I)
EVENT = re.compile(r"^Event\s+(?P<name>[A-Za-z_]\w*)\b", re.I)
CONST = re.compile(r"^Const\s+(?P<list>.*)$", re.I)
VARIABLE = re.compile(r"^(?:WithEvents\s+)?(?P<list>[A-Za-z_]\w*\b.*)$", re.I)
# Module-level statements that take a visibility keyword but have none.
IMPLICIT = re.compile(r"^\s*(?P<kind>Dim|Const|Enum|Type|Event|"
                      r"Declare\s+(?:PtrSafe\s+)?(?:Function|Sub))\b", re.I)
IDENTIFIER = re.compile(r"\s*(?:WithEvents\s+)?([A-Za-z_]\w*)", re.I)


@dataclass(frozen=True)
class Declaration:
    component: str
    kind: str
    name: str
    path: str
    line: int

    @property
    def key(self) -> str:
        return f"{self.component}.{self.name} ({self.kind})"


def top_level_parts(text: str) -> list[str]:
    """``text`` split at commas outside parentheses and string literals."""
    parts: list[str] = []
    depth = start = 0
    in_string = False
    index = 0
    while index < len(text):
        char = text[index]
        if char == '"':
            if in_string and text[index:index + 2] == '""':
                index += 2
                continue
            in_string = not in_string
        elif not in_string and char == "(":
            depth += 1
        elif not in_string and char == ")" and depth:
            depth -= 1
        elif not in_string and char == "," and not depth:
            parts.append(text[start:index])
            start = index + 1
        index += 1
    parts.append(text[start:])
    return parts


def names(declarators: str) -> list[str]:
    """The identifier of each declarator of a Const or variable list."""
    return [match.group(1) for part in top_level_parts(declarators)
            if (match := IDENTIFIER.match(part))]


def module_declaration(rest: str) -> list[tuple[str, str]]:
    """``(kind, name)`` pairs declared by a module-level statement after its
    visibility keyword."""
    if match := DECLARE.match(rest):
        return [(f"Declare {match.group('kind').capitalize()}", match.group("name"))]
    if match := EVENT.match(rest):
        return [("Event", match.group("name"))]
    if match := CONST.match(rest):
        return [("Const", name) for name in names(match.group("list"))]
    if match := VARIABLE.match(rest):
        return [("Variable", name) for name in names(match.group("list"))]
    return []


@dataclass
class Scan:
    """Public declarations and findings of one component in one environment."""

    path: str
    declarations: list[Declaration] = field(default_factory=list)
    findings: list[Finding] = field(default_factory=list)

    @property
    def standard(self) -> bool:
        return Path(self.path).suffix.lower() == ".bas"

    def declare(self, kind: str, name: str, line: int) -> None:
        self.declarations.append(Declaration(Path(self.path).stem, kind, name, self.path, line))

    def reject(self, line: int, message: str) -> None:
        self.findings.append(Finding(self.path, line, message))

    def visibility(self, line: int, keyword: str | None, what: str) -> bool:
        """Report a missing or misplaced keyword; True when the item is public."""
        folded = (keyword or "").casefold()
        if not folded:
            self.reject(line, f"implicit visibility: {what}; state Public or Private")
        elif folded == "friend" and self.standard:
            self.reject(line, f"Friend in a standard module: {what}; Friend is valid only in "
                              "a class module")
        return folded in {"public", "global"}


def skip_block(scan: Scan, match: re.Match[str], items: list[tuple[int, str]],
               index: int) -> int:
    """Record an Enum or Type; return the index after its End statement."""
    kind, name = match.group("kind").capitalize(), match.group("name")
    line = items[index][0]
    if scan.visibility(line, match.group("vis"), f"{kind} {name}"):
        scan.declare(kind, name, line)
    end = index + 1
    while end < len(items) and not END_BLOCK[kind.casefold()].match(items[end][1]):
        end += 1
    if end == len(items):
        scan.reject(line, f"{kind} {name} is not closed")
        return index + 1
    return end + 1


def scan_statement(scan: Scan, line: int, code: str) -> None:
    """A module-level statement other than a procedure, an Enum or a Type."""
    if implicit := IMPLICIT.match(code):
        kind = " ".join(implicit.group("kind").split()[:1])
        scan.visibility(line, None, f"module-level {kind} statement")
        return
    visible = VISIBLE.match(code)
    if visible is None:
        return
    declared = module_declaration(visible.group("rest"))
    if not declared:
        return
    what = ", ".join(f"{kind} {name}" for kind, name in declared)
    if scan.visibility(line, visible.group("vis"), what):
        for kind, name in declared:
            scan.declare(kind, name, line)


def scan_source(path: str, text: str) -> Scan:
    scan = Scan(path)
    items = statements(text)
    inside = False
    index = 0
    while index < len(items):
        line, code = items[index]
        index += 1
        if not code.strip() or code.lstrip().startswith(("Attribute ", "Option ")):
            continue
        if procedure := PROC.match(code):
            what = f"{' '.join(procedure.group('kind').split())} {procedure.group('name')}"
            if scan.visibility(line, procedure.group("vis"), what):
                kind = " ".join(part.capitalize() for part in procedure.group("kind").split())
                scan.declare(kind, procedure.group("name"), line)
            inside = True
        elif END_PROC.match(code):
            inside = False
        elif not inside and (block := BLOCK.match(code)):
            index = skip_block(scan, block, items, index - 1)
        elif not inside:
            scan_statement(scan, line, code)
    return scan


def scan_component(path: str, text: str) -> tuple[dict[Declaration, list[str]], list[Finding]]:
    """Public declarations with the environments that reach them, and the findings."""
    sources, findings = reachable_sources(path, text)
    declarations: dict[Declaration, list[str]] = {}
    for environment, source in sources.items():
        scan = scan_source(path, source)
        findings.extend(item for item in scan.findings if item not in findings)
        for declaration in scan.declarations:
            declarations.setdefault(declaration, []).append(environment)
    return declarations, findings


def property_pair(first: Declaration, second: Declaration) -> bool:
    """A Property Get with its Let or Set in the same component shares one name."""
    return (first.component.casefold() == second.component.casefold()
            and first.kind.casefold() in PROPERTY_KINDS
            and second.kind.casefold() in PROPERTY_KINDS
            and first.kind.casefold() != second.kind.casefold())


def collisions(declarations: dict[Declaration, list[str]], group: str) -> list[Finding]:
    """One public name declared twice by standard modules in a shared environment."""
    findings: list[Finding] = []
    seen: list[Declaration] = []
    for declaration, environments in declarations.items():
        if Path(declaration.path).suffix.lower() != ".bas":
            continue
        for previous in seen:
            shared = sorted(set(declarations[previous]) & set(environments))
            if not shared or previous.name.casefold() != declaration.name.casefold() \
                    or property_pair(previous, declaration):
                continue
            where = "twice in its module" if previous.path == declaration.path else \
                f"also in {previous.path} ({group})"
            findings.append(Finding(declaration.path, declaration.line,
                                    f"public name {declaration.name!r} declared {where}: "
                                    f"{previous.key} at line {previous.line}, an ambiguous "
                                    f"name in {', '.join(shared)}"))
        seen.append(declaration)
    return findings


def check(root: Path) -> tuple[list[Finding], int, int]:
    components = tracked_components(root)
    declarations: dict[Declaration, list[str]] = {}
    findings: list[Finding] = []
    for path in components:
        found, errors = scan_component(path, read_component(root, path))
        declarations.update(found)
        findings.extend(errors)
    for group, prefixes in COLLISION_GROUPS.items():
        members = {declaration: environments for declaration, environments
                   in declarations.items() if declaration.path.startswith(prefixes)}
        findings.extend(collisions(members, group))
    return findings, len(components), len(declarations)


def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=(__doc__ or "").splitlines()[0])
    parser.add_argument("--root", type=Path, default=ROOT)
    args = parser.parse_args(argv)
    findings, components, declarations = check(args.root.resolve())
    return report("vba-visibility", findings, f"{components} VBA component(s), "
                  f"{declarations} public declaration(s), "
                  f"collision group(s): {', '.join(COLLISION_GROUPS)}")


if __name__ == "__main__":
    sys.exit(main())
