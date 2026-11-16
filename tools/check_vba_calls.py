#!/usr/bin/env python3
"""Check VBA calls that resolve to a procedure of the project: their arguments
and whether a qualified call may reach its target.

Built on the procedure tables of check_vba_identifiers.py and applied in each
compilation environment of check_vba_conditionals.py. A call resolves when its
name, unshadowed by a parameter, variable, constant, control or other declared
name, is a procedure of the calling module, the one public procedure of that
name in a standard module, or, qualified (``Module.Name``), a procedure of the
standard module named. It reads statement calls (``Name args``,
``Call Name(args)``, ``Module.Name args``) and calls with parentheses inside an
expression (``x = Helper(1)``), and reports:

- too many or too few arguments, counting ``Optional`` parameters, a
  ``ParamArray`` and arguments passed by name (``Name:=Value``);
- a named argument that the procedure does not declare;
- a qualified call to a ``Private`` procedure of another module.

When unsure it says nothing: a member call (``obj.Method``, ``.Add x`` in a
``With`` block), a name that a variable or array shadows, a ``Property``
procedure, a ``Declare``, a built-in, a name public in two standard modules
and a procedure whose parameters cannot be read are not resolved, and parentheses
after a function without parameters may index its result, so they are not
counted. tools/README.md, "VBA source checks", says what it cannot see.

Ported from tools/check_vba_calls.py of danielep71/K-TEMPLATE at
9fef90ddd20b8f8c43b39b964a63ed186079b7a4 (MPL-2.0, as this repository). There it is
advisory; here it blocks, because the tracked source has no finding.
"""

from __future__ import annotations

import argparse
import re
import sys
from collections.abc import Sequence
from dataclasses import dataclass
from pathlib import Path

from check_vba_conditionals import reachable_sources
from check_vba_identifiers import (
    ELSE,
    MEMBER_MARK,
    NAME,
    REDIM,
    SINGLE_LINE_IF,
    STATEMENT_WORDS,
    SUFFIX,
    Module,
    Parameter,
    Procedure,
    closing_parenthesis,
    scan,
    top_level_split,
)
from vba_lexer import (
    Finding,
    date_delimiter,
    read_component,
    report,
    strip_code,
    tracked_components,
)

ROOT = Path(__file__).resolve().parents[1]
HEAD = re.compile(rf"^\s*(?:(?P<call>Call)\s+)?(?:(?P<qualifier>{NAME})\.)?(?P<name>{NAME})"
                  rf"{SUFFIX}", re.I)
# A name, qualified or not, and the parenthesis of its arguments; not a member
# of an object or a With block.
EXPRESSION_CALL = re.compile(rf"(?<![\w.!])(?:(?P<qualifier>{NAME})\.)?(?P<name>{NAME}){SUFFIX}"
                             r"\s*\(")
STRING = re.compile(r'"(?:[^"]|"")*"?')
NAMED_ARGUMENT = re.compile(rf"^\s*(?P<name>{NAME})\s*:=")
# Statements whose operands are names or constant bounds, not expressions to
# read. ReDim bounds are evaluated at run time, so ReDim is read apart.
DECLARING = frozenset("const declare dim enum erase event static type".split())


@dataclass(frozen=True)
class Target:
    module: Module
    procedure: Procedure
    qualified: bool

    @property
    def label(self) -> str:
        return f"{Path(self.module.path).stem}.{self.procedure.name}"


class Resolver:
    """Resolves the calls of one compilation environment."""

    def __init__(self, modules: list[Module]) -> None:
        self.standard = {Path(module.path).stem.casefold(): module
                         for module in modules if module.standard}
        # Public non-procedure names of standard modules shadow a procedure;
        # a public procedure name in two standard modules is ambiguous.
        self.public_names: set[str] = set()
        self.public: dict[str, list[Target]] = {}
        for module in self.standard.values():
            self.public_names |= module.public
            for procedure in module.procedures:
                if procedure.public:
                    self.public.setdefault(procedure.name.casefold(), []).append(
                        Target(module, procedure, False))
        self.resolved = 0

    @staticmethod
    def own(module: Module, name: str, qualified: bool) -> Target | None:
        found = [procedure for procedure in module.procedures
                 if procedure.name.casefold() == name]
        if len(found) != 1 or found[0].kind.startswith("property") or \
                found[0].parameters is None:
            return None
        return Target(module, found[0], qualified)

    def resolve(self, module: Module, caller: Procedure, qualifier: str | None,
                name: str) -> Target | None:
        name = name.casefold()
        local = caller.names | module.names
        if qualifier is not None:
            home = self.standard.get(qualifier.casefold())
            if home is None or qualifier.casefold() in local:
                return None
            return self.own(home, name, True)
        if name in local:
            return None
        if any(procedure.name.casefold() == name for procedure in module.procedures):
            return self.own(module, name, False)
        candidates = self.public.get(name, [])
        if name in self.public_names or len(candidates) != 1:
            return None
        return self.own(candidates[0].module, name, False)

    def statement(self, module: Module, caller: Procedure, line: int, statement: str
                  ) -> list[Finding]:
        """The findings of one statement of ``caller``'s body."""
        single = SINGLE_LINE_IF.match(statement)
        if single:
            findings = self.expression(module, caller, line, single.group("condition"))
            for part in ELSE.split(single.group("tail")):
                if part.strip():
                    findings += self.statement(module, caller, line, part)
            return findings
        head = HEAD.match(statement)
        if not head:
            # A With-block member, as in .Add Helper(1), or nothing to read.
            return self.expression(module, caller, line, statement)
        name = head.group("name")
        if not head.group("call") and not head.group("qualifier") and \
                name.casefold() in STATEMENT_WORDS:
            if name.casefold() in DECLARING:
                return []
            if name.casefold() == "redim":
                return self.redim_bounds(module, caller, line, statement)
            return self.expression(module, caller, line, statement[head.end():])
        rest = statement[head.end():]
        arguments = self.statement_arguments(rest.lstrip(), bool(head.group("call")))
        target = None if arguments is None else \
            self.resolve(module, caller, head.group("qualifier"), name)
        if arguments is None or target is None:
            # An assignment, a member access or an unresolved call: read its
            # expressions, the name and any index included, as in Helper(1).Refresh.
            start = head.start("qualifier") if head.group("qualifier") else head.start("name")
            return self.expression(module, caller, line, statement[start:])
        return self.checked(module, caller, line, target, arguments, statement=True) + \
            self.expression(module, caller, line, arguments)

    def redim_bounds(self, module: Module, caller: Procedure, line: int, statement: str
                     ) -> list[Finding]:
        """The findings of the calls in the bounds of each array a ReDim
        resizes, as in ReDim Preserve Values(1 To Count(2), 3); the array
        itself is not a call."""
        redim = REDIM.match(statement)
        findings: list[Finding] = []
        for item in top_level_split(redim.group("list")) if redim else []:
            start = item.find("(")
            end = closing_parenthesis(item, start) if start >= 0 else None
            if end is not None:
                findings += self.expression(module, caller, line, item[start + 1:end])
        return findings

    @staticmethod
    def statement_arguments(rest: str, explicit: bool) -> str | None:
        """The argument list of a statement call, or None when the statement
        is an assignment or a member access instead."""
        if rest.startswith((".", "=")) or MEMBER_MARK.match(rest):
            return None
        if not rest.startswith("("):
            return None if explicit and rest else rest
        end = closing_parenthesis(rest, 0)
        if end is None:
            return None
        after = rest[end + 1:].lstrip()
        if after.startswith((".", "=")) or MEMBER_MARK.match(after):
            return None
        if explicit:
            return rest[1:end] if not after else None
        return rest

    def expression(self, module: Module, caller: Procedure, line: int, text: str
                   ) -> list[Finding]:
        """The findings of the calls with parentheses inside ``text``."""
        findings: list[Finding] = []
        for match in EXPRESSION_CALL.finditer(text):
            target = self.resolve(module, caller, match.group("qualifier"), match.group("name"))
            end = closing_parenthesis(text, match.end() - 1)
            if target is None or end is None:
                continue
            findings += self.checked(module, caller, line, target, text[match.end():end],
                                     statement=False)
        return findings

    def checked(self, module: Module, caller: Procedure, line: int, target: Target,
                arguments: str, *, statement: bool) -> list[Finding]:
        """The findings of one resolved call with its argument list."""
        self.resolved += 1
        problems: list[str] = []
        if target.qualified and not target.procedure.public and target.module is not module:
            problems.append(f"{target.procedure.name} is Private in "
                            f"{Path(target.module.path).stem}, so a call from another module "
                            "cannot reach it")
        parameters = target.procedure.parameters or []
        # Parentheses after a function without parameters may index its result.
        if statement or parameters:
            problems += argument_problems(parameters, arguments)
        return [Finding(module.path, line, f"call to {target.label} in {caller.name}: {problem}")
                for problem in problems]


def argument_problems(parameters: list[Parameter], arguments: str) -> list[str]:
    """What is wrong with ``arguments`` for ``parameters``; nothing when the
    list cannot be read, as with a positional argument after a named one."""
    positional: list[str] = []
    named: list[str] = []
    for item in top_level_split(arguments) if arguments.strip() else []:
        argument = NAMED_ARGUMENT.match(item)
        if argument:
            named.append(argument.group("name"))
        elif named:
            return []
        else:
            positional.append(item.strip())
    problems: list[str] = []
    if not any(parameter.paramarray for parameter in parameters) and \
            len(positional) > len(parameters):
        problems.append(f"{len(positional)} argument(s) passed, at most {len(parameters)} "
                        "accepted")
    # A ParamArray cannot be passed by name; when it is, VBA says so.
    names = {parameter.name.casefold() for parameter in parameters}
    problems += [f"no parameter is named {item}" for item in named
                 if item.casefold() not in names]
    given = {item.casefold() for item in named}
    missing = [parameter.name for index, parameter in enumerate(parameters)
               if not parameter.optional and not parameter.paramarray
               and parameter.name.casefold() not in given
               and (index >= len(positional) or not positional[index])]
    if missing:
        problems.append(f"required argument(s) {', '.join(missing)} not passed")
    return problems


def dates_as_values(code: str) -> str:
    """``code`` with each date literal, such as #January 1, 2020#, replaced by
    zeros, so that its comma separates no arguments. A # type suffix, as in
    Amount# or 2#, opens no literal; strings must already be replaced."""
    result = list(code)
    in_date = False
    for index, char in enumerate(code):
        if char == "#" and date_delimiter(code, index, in_date):
            in_date = not in_date
            result[index] = "0"
        elif in_date:
            result[index] = "0"
    return "".join(result)


def literals_as_values(source: str) -> str:
    """``source`` with each string and date literal of its code replaced by zeros.

    The identifier scan blanks string literals, so that ``Proc "a", , "c"``
    would read as three omitted arguments; a literal of the same length that
    is not a string keeps each argument visible and every column in place.
    A conditional-compilation line, such as #If VBA7 Then, is left as it is."""
    lines = []
    for raw in source.split("\n"):
        code = strip_code(raw)
        values = STRING.sub(lambda match: "0" * len(match.group()), code)
        if not values.lstrip().startswith("#"):
            values = dates_as_values(values)
        lines.append(values + raw[len(code):])
    return "\n".join(lines)


def environment_findings(modules: list[Module]) -> tuple[list[Finding], int]:
    """The findings of one environment, and how many calls it resolved."""
    resolver = Resolver(modules)
    findings: list[Finding] = []
    for module in modules:
        for procedure in module.procedures:
            for line, statement in procedure.body:
                findings += resolver.statement(module, procedure, line, statement)
    return findings, resolver.resolved


def check(root: Path) -> tuple[list[Finding], int, int]:
    components = tracked_components(root)
    by_environment: dict[str, list[Module]] = {}
    for path in components:
        # Unreadable directives are vba-conditionals findings, not repeated here.
        sources, _ = reachable_sources(path, read_component(root, path))
        for environment, source in sources.items():
            by_environment.setdefault(environment, []).append(
                scan(path, literals_as_values(source)))
    environments: dict[Finding, list[str]] = {}
    resolved = 0
    for environment, modules in by_environment.items():
        found, count = environment_findings(modules)
        resolved = max(resolved, count)
        for finding in found:
            names = environments.setdefault(finding, [])
            if environment not in names:
                names.append(environment)
    findings = [Finding(item.path, item.line, f"{item.message} [{', '.join(names)}]")
                for item, names in environments.items()]
    return findings, len(components), resolved


def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=(__doc__ or "").splitlines()[0])
    parser.add_argument("--root", type=Path, default=ROOT)
    args = parser.parse_args(argv)
    findings, components, resolved = check(args.root.resolve())
    return report("vba-calls", findings, f"{components} VBA component(s), {resolved} resolved "
                                         "call(s)")


if __name__ == "__main__":
    sys.exit(main())
