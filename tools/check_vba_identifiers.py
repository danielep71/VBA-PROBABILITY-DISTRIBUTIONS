#!/usr/bin/env python3
"""Check VBA identifiers: declared before assignment, called names that exist,
and no VBA keyword used as a name.

Three rules, each applied in every compilation environment of
check_vba_conditionals.py, so that a declaration in an inactive #If branch
never satisfies an assignment in an active one:

- every name a procedure assigns, with ``=``, ``Let``, ``Set``, ``LSet``,
  ``RSet`` or ``For`` (``For Each`` included), is declared: a parameter, a
  ``Dim``, ``Static``, ``Const`` or ``ReDim`` of the procedure, a module-level
  name, a control of a form, the procedure's own name, a property of the
  module, a public name of a standard module, or a property the host
  provides (HOST_TARGETS, such as Range and Cells anywhere, a UserForm's
  Caption in a form, a Worksheet's Visible in a document module). ``ReDim`` declares, as VBA
  does. A member access (``x.y =``, ``.y =`` in a ``With`` block) is not read;
- every statement-level call, ``Name args`` or ``Call Name(args)``, names a
  procedure or ``Declare`` of the same module, a public procedure or
  ``Declare`` of a standard module, a name of BUILTIN_CALLS (VBA functions
  and Excel methods used as statements), or a method the host provides in
  that kind of module (HOST_CALLS: a UserForm's Hide in a form, a
  Worksheet's Protect in a document module). A qualified call (``Module.Name``,
  ``obj.Method``) and a call inside an expression are not resolved;
- no procedure, parameter, variable, constant, Enum, Enum member, Type,
  Declare or Event is named like a VBA keyword (KEYWORDS). A bracketed Enum
  member such as ``[_First]`` may use any name, as VBA allows.

A statement the analysis cannot read, such as a parameter list it cannot
split, or a single-line If without Then, is a finding, not a silent pass.
tools/README.md, "VBA source checks", lists what the rules cannot see.

A name of ACCEPTED_KEYWORD_NAMES is not reported: it is in a released module
that Excel compiled, and a rename changes certified source, so it waits for the
next Excel wave. Each entry names its file and its reason.

Ported from tools/check_vba_identifiers.py of danielep71/K-TEMPLATE at
9fef90ddd20b8f8c43b39b964a63ed186079b7a4 (MPL-2.0, as this repository). The form,
class and document-module rules are kept unchanged although this repository
tracks standard modules only: one implementation, not a fork to maintain.
"""

from __future__ import annotations

import argparse
import re
import sys
from collections.abc import Iterator, Sequence
from dataclasses import dataclass, field
from pathlib import Path

from check_vba_conditionals import reachable_sources
from vba_lexer import (
    Finding,
    date_delimiter,
    logical_units,
    read_component,
    report,
    split_statements,
    tracked_components,
)

ROOT = Path(__file__).resolve().parents[1]
NAME = r"[A-Za-z][A-Za-z0-9_]*"
# A type suffix, as in Total$ or Values!(1); a ! before a name is a member
# access instead, as in Me!txtAmount or Items!Keys.
SUFFIX = r"(?:[%&!#@$](?![A-Za-z_]))?"
MEMBER_MARK = re.compile(r"\.|![A-Za-z_]")
VISIBILITY = r"(?:(?P<vis>Public|Private|Friend|Global)\s+)?"
PROC = re.compile(rf"^\s*{VISIBILITY}(?:Static\s+)?(?P<kind>Sub|Function|Property\s+(?:Get|Let|Set))"
                  rf"\s+(?P<name>{NAME}){SUFFIX}\s*(?P<rest>.*)$", re.I)
END_PROC = re.compile(r"^\s*End\s+(?:Sub|Function|Property)\b", re.I)
DECLARE = re.compile(rf"^\s*{VISIBILITY}Declare\s+(?:PtrSafe\s+)?(?:Sub|Function)\s+"
                     rf"(?P<name>{NAME})", re.I)
EVENT = re.compile(rf"^\s*{VISIBILITY}Event\s+(?P<name>{NAME})", re.I)
BLOCK = re.compile(rf"^\s*{VISIBILITY}(?P<kind>Enum|Type)\s+(?P<name>{NAME})\s*$", re.I)
END_BLOCK = re.compile(r"^\s*End\s+(?P<kind>Enum|Type)\b", re.I)
MODULE_DECLARATION = re.compile(r"^\s*(?P<vis>Public|Private|Global|Dim)\s+(?P<const>Const\s+)?"
                                r"(?P<list>.*)$", re.I)
MODULE_CONST = re.compile(r"^\s*Const\s+(?P<list>.*)$", re.I)
LOCAL_DECLARATION = re.compile(r"^\s*(?:Dim|Static|Const)\s+(?P<list>.*)$", re.I)
REDIM = re.compile(r"^\s*ReDim\s+(?:Preserve\s+)?(?P<list>.*)$", re.I)
# What may follow a declared name: bounds, a type and a value, nothing else.
DECLARED_REST = r"\s*(?:\(.*\))?\s*(?:As\s+\S.*?)?\s*(?:=.*)?"
DECLARED_ITEM = re.compile(rf"\s*(?:WithEvents\s+)?(?P<name>{NAME}){SUFFIX}{DECLARED_REST}", re.I)
PARAMETER = re.compile(rf"\s*(?P<optional>Optional\s+)?(?:(?:ByVal|ByRef)\s+)?"
                       rf"(?P<paramarray>ParamArray\s+)?(?P<name>{NAME}){SUFFIX}{DECLARED_REST}",
                       re.I)
ENUM_MEMBER = re.compile(rf"^\s*(?:\[(?P<bracketed>[^\]]+)\]|(?P<name>{NAME}))\s*(?:=.*)?$")
# The controls of a form, declared in its header: Begin <ProgID> <Name>.
FORM_CONTROL = re.compile(rf"^\s*Begin\s+\S+\s+(?P<name>{NAME})\s*$", re.I)
LABEL = re.compile(rf"^\s*(?:(?P<name>{NAME})|\d+):(?!=)")
LINE_NUMBER = re.compile(r"^\s*\d+\s+")
LEAD = re.compile(rf"^\s*(?:(?P<keyword>Let|Set|LSet|RSet)\s+)?(?P<name>{NAME}){SUFFIX}", re.I)
FOR = re.compile(rf"^\s*For\s+(?:Each\s+)?(?P<name>{NAME}){SUFFIX}\b", re.I)
CALL = re.compile(r"^\s*Call\s+", re.I)
SINGLE_LINE_IF = re.compile(r"^\s*If\b(?P<condition>.*?)\bThen\b(?P<tail>.*)$", re.I)
ELSE = re.compile(r"\bElse\b", re.I)

# Words a statement may start with that are not a call or an assignment.
STATEMENT_WORDS = frozenset("""
    appactivate attribute beep case chdir chdrive close const date debug declare
    deletesetting dim do doevents else elseif end enum erase error event exit
    filecopy for friend function get global gosub goto if implements input kill let
    line load lock loop lset mkdir name next on open option print private property
    public put raiseevent randomize redim rem reset resume return rmdir rset
    savesetting seek select set setattr static stop sub time type unload unlock
    wend while width with write
""".split())
# Names a statement may call without a declaration in the project: VBA
# functions used as statements and the methods of Excel's Global object,
# which VBA resolves unqualified, in every module; a UserForm's own methods in a form;
# a Worksheet's or Workbook's in a document module (one with VB_Base).
# Extend the list of the module kind when code calls another one.
BUILTIN_CALLS = frozenset("""
    calculate ddeexecute ddepoke ddeterminate doevents executeexcel4macro inputbox msgbox run
    sendkeys shell
""".split())
HOST_CALLS = {
    "form": frozenset("""
        hide move printform redoaction repaint scroll setdefaulttaborder show undoaction
        whatsthismode
    """.split()),
    "document": frozenset("""
        activate calculate copy delete move paste printout printpreview protect refreshall
        save saveas showalldata unprotect
    """.split()),
}
# Built-in statement targets: Mid$(Text, 1, 1) = "x" replaces characters.
BUILTIN_TARGETS = frozenset({"mid", "mid$", "midb", "midb$"})
# Properties the host provides without a declaration, assignable unqualified:
# the writable members of Excel's Global object in every module, a UserForm's own properties in a form,
# and a Worksheet's or Workbook's in a document module (one with VB_Base).
# Extend the list of the module kind when code assigns another one.
HOST_TARGETS = {
    "any": frozenset("activecell activeprinter cells columns range rows selection".split()),
    "form": frozenset("""
        backcolor bordercolor borderstyle caption cycle drawbuffer enabled font forecolor
        height helpcontextid keepscrollbarsvisible left mousepointer picture
        picturealignment picturesizemode picturetiling righttoleft scrollbars scrollheight
        scrollleft scrolltop scrollwidth specialeffect startupposition tag top visible
        whatsthisbutton whatsthishelp width zoom
    """.split()),
    "document": frozenset("""
        displaypagebreaks displayrighttoleft enableautofilter enablecalculation
        enableformatconditionscalculation enableoutlining enableselection saved scrollarea
        standardwidth visible
    """.split()),
}
# "scale" is this repository's addition: a parameter of that name stopped the
# project from compiling in Excel (pull request 242), which is why the
# distribution parameters are named ScaleParam.
KEYWORDS = frozenset("""
    addressof alias and any as attribute base boolean byref byte byval call case
    close const currency date decimal declare dim do double each else elseif empty
    end enum eqv erase error event exit false for friend function get global gosub
    goto if imp implements in input integer is let lib like line long longlong
    longptr loop lset me mod name new next not nothing null object on open option
    optional or paramarray preserve print private property public raiseevent redim
    rem resume return rset scale select set single static step stop string sub then
    time to true type typeof until variant wend while with withevents xor
""".split())
# (path, name in lower case) -> why the name stays for now.
ACCEPTED_KEYWORD_NAMES = {
    ("src/M_STATS_PROBDIST_SPECIALFUNCS.bas", "base"):
        "a local of the released v1.0.0 source, compiled and certified in Excel; Base is "
        "reserved only after Option. Renamed in the next Excel wave (issue 400)",
}


@dataclass(frozen=True)
class Parameter:
    name: str
    optional: bool
    paramarray: bool


@dataclass
class Procedure:
    name: str
    kind: str
    line: int
    public: bool = True
    # None when the parameter list cannot be read.
    parameters: list[Parameter] | None = field(default_factory=list)
    names: set[str] = field(default_factory=set)
    assignments: list[tuple[int, str]] = field(default_factory=list)
    calls: list[tuple[int, str]] = field(default_factory=list)
    # Every statement of the body, for checks that read more than names.
    body: list[tuple[int, str]] = field(default_factory=list)


@dataclass
class Module:
    path: str
    standard: bool
    kind: str = "any"
    names: set[str] = field(default_factory=set)
    public: set[str] = field(default_factory=set)
    callables: set[str] = field(default_factory=set)
    public_callables: set[str] = field(default_factory=set)
    properties: set[str] = field(default_factory=set)
    procedures: list[Procedure] = field(default_factory=list)
    declared: list[tuple[int, str, str]] = field(default_factory=list)
    findings: list[Finding] = field(default_factory=list)

    def declare(self, line: int, name: str, what: str, public: bool) -> None:
        self.names.add(name.casefold())
        if public:
            self.public.add(name.casefold())
        self.declared.append((line, name, what))


def top_level_split(text: str, separator: str = ",") -> list[str]:
    """Split at ``separator`` outside parentheses and date literals; strings
    are already blanked. The comma of ``#January 1, 2020#`` separates nothing
    (this repository's correction; the template splits there)."""
    parts: list[str] = []
    depth = start = 0
    in_date = False
    for index, char in enumerate(text):
        if char == "#" and date_delimiter(text, index, in_date):
            in_date = not in_date
        elif in_date:
            continue
        elif char == "(":
            depth += 1
        elif char == ")" and depth:
            depth -= 1
        elif char == separator and not depth:
            parts.append(text[start:index])
            start = index + 1
    parts.append(text[start:])
    return parts


def closing_parenthesis(text: str, start: int) -> int | None:
    """Index of the parenthesis closing the one at ``start``, or None."""
    depth = 0
    for index in range(start, len(text)):
        if text[index] == "(":
            depth += 1
        elif text[index] == ")":
            depth -= 1
            if depth == 0:
                return index
    return None


def single_line_if(statement: str) -> bool:
    """Whether ``statement`` is an If with a statement after its Then."""
    match = SINGLE_LINE_IF.match(statement)
    return match is not None and bool(match.group("tail").strip())


def statements(source: str) -> Iterator[tuple[int, str]]:
    """``(line, statement)`` of each statement, strings blanked, labels removed."""
    for line, _, kind, code in logical_units(source.split("\n"), blank_strings=True):
        if kind == "directive":
            continue
        label = LABEL.match(code)
        if label and (label.group("name") or "").casefold() not in STATEMENT_WORDS:
            code = code[label.end():]
        code = LINE_NUMBER.sub("", code, count=1)
        parts = split_statements(code)
        # In a single-line If, a colon separates the statements of a branch
        # and Else starts the other branch: split there too, so that no
        # branch is read as part of the statement before it.
        if any(single_line_if(part) for part in parts):
            parts = [piece for part in parts for piece in ELSE.split(part)]
        for statement in parts:
            if statement.strip():
                yield line, statement


def declared_names(text: str) -> list[str] | None:
    """The names a declaration list declares, or None when one cannot be read."""
    names = []
    for item in top_level_split(text):
        match = DECLARED_ITEM.fullmatch(item)
        if not match:
            return None
        names.append(match.group("name"))
    return names


def parameters(module: Module, line: int, procedure: Procedure, rest: str) -> None:
    """Read the parameter list that opens ``rest``, or report that it cannot."""
    if not rest.startswith("("):
        return
    end = closing_parenthesis(rest, 0)
    inner = rest[1:end] if end is not None else None
    if inner is None:
        procedure.parameters = None
        module.findings.append(Finding(module.path, line, f"cannot read the parameters of "
                                                          f"{procedure.name}"))
        return
    if not inner.strip():
        return
    for item in top_level_split(inner):
        match = PARAMETER.fullmatch(item)
        if not match:
            procedure.parameters = None
            module.findings.append(Finding(module.path, line, f"cannot read the parameter "
                                                              f"{item.strip()!r} of "
                                                              f"{procedure.name}"))
            continue
        procedure.names.add(match.group("name").casefold())
        if procedure.parameters is not None:
            procedure.parameters.append(Parameter(match.group("name"),
                                                  bool(match.group("optional")),
                                                  bool(match.group("paramarray"))))
        module.declared.append((line, match.group("name"), "parameter"))


def module_level(module: Module, line: int, statement: str, block: list[str]) -> None:
    """Record one module-level statement; ``block`` holds an open Enum or Type."""
    if block:
        end = END_BLOCK.match(statement)
        if end:
            block.clear()
            return
        if block[0] == "enum":
            member = ENUM_MEMBER.match(statement)
            if member is None:
                module.findings.append(Finding(module.path, line, "cannot read the Enum member "
                                                                  f"{statement.strip()!r}"))
            elif member.group("name"):
                module.declare(line, member.group("name"), "Enum member", block[1] == "public")
        return
    opened = BLOCK.match(statement)
    if opened:
        public = (opened.group("vis") or "Public").casefold() in {"public", "global"}
        module.declare(line, opened.group("name"), opened.group("kind"), public)
        block[:] = [opened.group("kind").casefold(), "public" if public else "private"]
        return
    for pattern, what in ((DECLARE, "Declare"), (EVENT, "Event")):
        match = pattern.match(statement)
        if match:
            public = (match.group("vis") or "Public").casefold() in {"public", "global"}
            module.declare(line, match.group("name"), what, public)
            if what == "Declare":
                module.callables.add(match.group("name").casefold())
                if public:
                    module.public_callables.add(match.group("name").casefold())
            return
    match = MODULE_DECLARATION.match(statement) or MODULE_CONST.match(statement)
    if not match:
        return  # Option, Attribute, Implements and the header of a class or form.
    groups = match.groupdict()
    visibility = (groups.get("vis") or "Private").casefold()
    names = declared_names(groups["list"])
    if names is None:
        module.findings.append(Finding(module.path, line, "cannot read the declaration "
                                                          f"{statement.strip()!r}"))
        return
    what = "Const" if groups.get("const") or statement.lstrip().lower().startswith("const") \
        else "variable"
    for name in names:
        module.declare(line, name, what, visibility in {"public", "global"})


def member_target(item: str) -> bool:
    """Whether a ReDim target is a member, such as State.Values(1 To 3) or
    .Values(1 To 3) in a With block: resized, not declared."""
    head = item.lstrip()
    bounds = head.find("(")
    return MEMBER_MARK.search(head[:bounds] if bounds >= 0 else head) is not None


def local_declaration(module: Module, procedure: Procedure, line: int,
                      statement: str) -> bool:
    """Record a Dim, Static, Const or ReDim; whether ``statement`` was one."""
    redim = REDIM.match(statement)
    match = redim or LOCAL_DECLARATION.match(statement)
    if not match:
        return False
    items = match.group("list")
    if redim:
        items = ",".join(item for item in top_level_split(items) if not member_target(item))
        if not items.strip():
            return True
    names = declared_names(items)
    if names is None:
        module.findings.append(Finding(module.path, line, "cannot read the declaration "
                                                          f"{statement.strip()!r}"))
        return True
    for name in names:
        procedure.names.add(name.casefold())
        module.declared.append((line, name, "variable"))
    return True


def procedure_statement(module: Module, procedure: Procedure, line: int, statement: str) -> None:
    """Record the declarations, assignments and calls of one procedure statement."""
    if local_declaration(module, procedure, line, statement):
        return
    single = SINGLE_LINE_IF.match(statement)
    if single:
        for part in ELSE.split(single.group("tail")):
            if part.strip():
                procedure_statement(module, procedure, line, part)
        return
    if statement.split(None, 1)[0].casefold() == "if":
        module.findings.append(Finding(module.path, line, "cannot read the If statement "
                                                          f"{statement.strip()!r}: no Then"))
        return
    loop = FOR.match(statement)
    if loop:
        procedure.assignments.append((line, loop.group("name")))
        return
    call = CALL.match(statement)
    if call:
        target_or_call(module, procedure, line, statement[call.end():], explicit_call=True)
        return
    target_or_call(module, procedure, line, statement)


def target_or_call(module: Module, procedure: Procedure, line: int, statement: str, *,
                   explicit_call: bool = False) -> None:
    """Classify a statement that starts with a name: an assignment or a call.

    The same reading serves ``Call Name(args)``, with ``explicit_call``: past the
    name and any index in parentheses, a ``.`` or a ``!`` before a name is a
    member access, as in ``Call Items(1).Refresh`` or ``Call Me!txtAmount.SetFocus``,
    which no rule resolves."""
    lead = LEAD.match(statement)
    if not lead or statement.lstrip().startswith("."):
        return  # A member of a With block, or nothing a rule reads.
    keyword = None if explicit_call else lead.group("keyword")
    name = lead.group("name")
    if keyword is None and name.casefold() in STATEMENT_WORDS:
        return
    rest = statement[lead.end():].lstrip()
    after = rest
    if rest.startswith("("):
        end = closing_parenthesis(rest, 0)
        if end is None:
            module.findings.append(Finding(module.path, line, "cannot read the statement "
                                                              f"{statement.strip()!r}"))
            return
        after = rest[end + 1:].lstrip()
    if after.startswith(".") or MEMBER_MARK.match(after):
        return  # A member access, such as x.y = 1, x(1).y = 1 or Me!txtAmount = 1.
    if explicit_call:
        procedure.calls.append((line, name))
    elif after.startswith("=") or keyword is not None:
        procedure.assignments.append((line, name))
    else:
        procedure.calls.append((line, name))


def scan(path: str, source: str) -> Module:
    """The declarations, assignments and calls of one component's source."""
    module = Module(path, path.lower().endswith(".bas"))
    header = path.lower().endswith(".frm")
    if header:
        module.kind = "form"
    current: Procedure | None = None
    block: list[str] = []
    for line, statement in statements(source):
        if header:
            if statement.lower().startswith("attribute "):
                header = False
            else:
                control = FORM_CONTROL.match(statement)
                if control:
                    module.names.add(control.group("name").casefold())
                continue
        if statement.lower().startswith("attribute vb_base ") and path.lower().endswith(".cls"):
            module.kind = "document"
        opener = PROC.match(statement) if not block else None
        if opener:
            kind = re.sub(r"\s+", " ", opener.group("kind")).casefold()
            public = (opener.group("vis") or "Public").casefold() in {"public", "global"}
            current = Procedure(opener.group("name"), kind, line, public)
            module.procedures.append(current)
            module.declared.append((line, current.name, "procedure"))
            module.callables.add(current.name.casefold())
            if public:
                module.public_callables.add(current.name.casefold())
            if kind.startswith("property"):
                module.properties.add(current.name.casefold())
            parameters(module, line, current, opener.group("rest"))
            continue
        if END_PROC.match(statement):
            current = None
        elif current is not None:
            current.body.append((line, statement))
            procedure_statement(module, current, line, statement)
        else:
            module_level(module, line, statement, block)
    return module


def environment_findings(modules: list[Module]) -> list[Finding]:
    """The findings of one environment, across every component."""
    public_names: set[str] = set()
    public_callables: set[str] = set()
    for module in modules:
        if module.standard:
            public_names |= module.public | module.public_callables
            public_callables |= module.public_callables
    findings: list[Finding] = []
    for module in modules:
        findings.extend(module.findings)
        findings.extend(Finding(module.path, line, f"{what} {name} is named like the VBA keyword "
                                                   f"{name.capitalize()}")
                        for line, name, what in module.declared
                        if name.casefold() in KEYWORDS
                        and (module.path, name.casefold()) not in ACCEPTED_KEYWORD_NAMES)
        for procedure in module.procedures:
            assignable = (procedure.names | module.names | module.properties | public_names
                          | {procedure.name.casefold()} | BUILTIN_TARGETS
                          | HOST_TARGETS["any"] | HOST_TARGETS.get(module.kind, frozenset()))
            findings.extend(Finding(module.path, line, f"{name} is assigned in {procedure.name} "
                                                       "but not declared")
                            for line, name in procedure.assignments
                            if name.casefold() not in assignable)
            callable_names = (module.callables | public_callables | BUILTIN_CALLS
                              | HOST_CALLS.get(module.kind, frozenset()))
            findings.extend(Finding(module.path, line, f"{name} is called in {procedure.name} "
                                                       "but no procedure of that name is "
                                                       "reachable")
                            for line, name in procedure.calls
                            if name.casefold() not in callable_names)
    return findings


def check(root: Path) -> tuple[list[Finding], int, int]:
    components = tracked_components(root)
    by_environment: dict[str, list[Module]] = {}
    directive_findings: list[Finding] = []
    procedures = 0
    for path in components:
        sources, findings = reachable_sources(path, read_component(root, path))
        directive_findings.extend(findings)
        for environment, source in sources.items():
            by_environment.setdefault(environment, []).append(scan(path, source))
    environments: dict[Finding, list[str]] = {}
    for environment, modules in by_environment.items():
        procedures = max(procedures, sum(len(module.procedures) for module in modules))
        for finding in environment_findings(modules):
            environments.setdefault(finding, []).append(environment)
    findings = directive_findings + [
        Finding(item.path, item.line, f"{item.message} [{', '.join(names)}]")
        for item, names in environments.items()]
    return findings, len(components), procedures


def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=(__doc__ or "").splitlines()[0])
    parser.add_argument("--root", type=Path, default=ROOT)
    args = parser.parse_args(argv)
    findings, components, procedures = check(args.root.resolve())
    return report("vba-identifiers", findings,
                  f"{components} VBA component(s), {procedures} procedure(s)")


if __name__ == "__main__":
    sys.exit(main())
