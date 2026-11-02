#!/usr/bin/env python3
"""Check VBA conditional compilation and Declare statements in every modeled host.

Evaluates every #If block of the tracked VBA components in each modeled
compilation environment and rejects what it cannot evaluate: an unknown
symbol, a malformed or duplicate module #Const, an unbalanced block. A
module-level #Const is supported; its value must be a Boolean expression over
the compiler constants and the module's earlier #Const symbols. A Declare
reachable under VBA7 must be PtrSafe. check_vba_visibility.py reads the
reachable source of each environment through reachable_sources().
"""

from __future__ import annotations

import argparse
import re
import sys
from collections.abc import Callable, Sequence
from dataclasses import dataclass, field
from pathlib import Path

from vba_lexer import (
    Finding,
    logical_units,
    read_component,
    report,
    split_statements,
    tracked_components,
)

ROOT = Path(__file__).resolve().parents[1]
DECLARE = re.compile(
    r"^\s*(?:(?:Public|Private|Global)\s+)?Declare\s+(?P<ptrsafe>PtrSafe\s+)?(?:Function|Sub)\b",
    re.IGNORECASE)
DIRECTIVE = re.compile(r"^\s*#\s*(If|ElseIf|Else|End\s+If)\b(.*)$", re.IGNORECASE)
CONST = re.compile(r"^\s*#\s*Const\b(.*)$", re.IGNORECASE)
CONST_BODY = re.compile(r"^\s+(?P<name>[A-Za-z]\w*)\s*=(?P<expression>.*)$")
TOKEN = re.compile(r"\s*(\(|\)|<>|=|-?\d+|[A-Za-z_]\w*)")

# Microsoft's compiler constants. VBA6 means version-6 compatible, so it is
# True under VBA7 as well; Win32 is True on 32-bit and 64-bit Windows, and
# Win64 on 64-bit Office as well; Win16 and Mac are False on every modeled
# host. 64-bit Excel (vba7-win64) is the host the Excel regression runs on
# (INSTALLATION.md); vba7-win32 and the pre-VBA7 vba6-win32 are modeled so that
# a branch written for them is still evaluated and its declarations checked.
ENVIRONMENTS: dict[str, dict[str, bool]] = {
    "vba6-win32": {"VBA6": True, "VBA7": False, "WIN16": False, "WIN32": True, "WIN64": False,
                   "MAC": False},
    "vba7-win32": {"VBA6": True, "VBA7": True, "WIN16": False, "WIN32": True, "WIN64": False,
                   "MAC": False},
    "vba7-win64": {"VBA6": True, "VBA7": True, "WIN16": False, "WIN32": True, "WIN64": True,
                   "MAC": False},
}
SYMBOLS = frozenset(ENVIRONMENTS["vba7-win64"])
VBA7_ENVIRONMENTS = tuple(name for name, symbols in ENVIRONMENTS.items() if symbols["VBA7"])


class ExpressionError(ValueError):
    """A conditional expression outside the supported model."""


class ExpressionParser:
    """Boolean expressions over the compiler constants, the module's #Const
    symbols, True, False, -1 and 0, with Not, And, Or, = and <>."""

    def __init__(self, expression: str, symbols: dict[str, bool]) -> None:
        self.symbols = symbols
        self.tokens = self.tokenize(expression)
        self.index = 0

    @staticmethod
    def tokenize(expression: str) -> list[str]:
        tokens: list[str] = []
        position = 0
        while position < len(expression):
            match = TOKEN.match(expression, position)
            if match is None:
                if not expression[position:].strip():
                    break
                raise ExpressionError(f"unsupported token near {expression[position:]!r}")
            tokens.append(match.group(1))
            position = match.end()
        if not tokens:
            raise ExpressionError("empty conditional expression")
        return tokens

    def peek(self) -> str:
        return self.tokens[self.index] if self.index < len(self.tokens) else ""

    def take(self) -> str:
        token = self.peek()
        if not token:
            raise ExpressionError("unexpected end of expression")
        self.index += 1
        return token

    def parse(self) -> bool:
        value = self.parse_or()
        if self.peek():
            raise ExpressionError(f"unexpected token {self.peek()!r}")
        return value

    def parse_or(self) -> bool:
        value = self.parse_and()
        while self.peek().casefold() == "or":
            self.take()
            right = self.parse_and()
            value = value or right
        return value

    def parse_and(self) -> bool:
        value = self.parse_not()
        while self.peek().casefold() == "and":
            self.take()
            right = self.parse_not()
            value = value and right
        return value

    def parse_not(self) -> bool:
        if self.peek().casefold() == "not":
            self.take()
            return not self.parse_not()
        left = self.parse_atom()
        operator = self.peek()
        if operator not in {"=", "<>"}:
            return left
        self.take()
        right = self.parse_atom()
        return left == right if operator == "=" else left != right

    def parse_atom(self) -> bool:
        token = self.take()
        if token == "(":  # noqa: S105 - a VBA expression token, not a password
            value = self.parse_or()
            if self.take() != ")":
                raise ExpressionError("missing closing parenthesis")
            return value
        if token.casefold() in {"true", "false"}:
            return token.casefold() == "true"
        if re.fullmatch(r"-?\d+", token):
            if int(token) not in {-1, 0}:
                raise ExpressionError("only the Boolean integer literals -1 and 0 are supported")
            return int(token) == -1
        if token.upper() not in self.symbols:
            raise ExpressionError(f"unsupported conditional-compilation symbol {token!r}; "
                                  "supported are VBA6, VBA7, Win16, Win32, Win64, Mac and "
                                  "a #Const defined earlier in the module")
        return self.symbols[token.upper()]


def evaluate(expression: str, symbols: dict[str, bool]) -> bool:
    return ExpressionParser(expression, symbols).parse()


@dataclass
class Frame:
    """One open #If block in one environment."""

    parent_active: bool
    branch_taken: bool
    current_active: bool
    else_seen: bool = False


@dataclass
class State:
    """The open #If blocks and the module's #Const symbols, per environment."""

    path: str
    findings: list[Finding] = field(default_factory=list)
    stacks: dict[str, list[Frame]] = field(
        default_factory=lambda: {name: [] for name in ENVIRONMENTS})
    constants: dict[str, dict[str, bool]] = field(
        default_factory=lambda: {name: {} for name in ENVIRONMENTS})

    def active(self, environment: str) -> bool:
        return all(frame.current_active for frame in self.stacks[environment])

    def symbols(self, environment: str) -> dict[str, bool]:
        return {**ENVIRONMENTS[environment], **self.constants[environment]}

    def reject(self, line: int | None, message: str) -> None:
        self.findings.append(Finding(self.path, line, message))


def condition_values(state: State, line: int, label: str, remainder: str,
                     eligible: dict[str, bool]) -> dict[str, bool]:
    """The condition's value in each environment that can enter the branch.

    An environment that cannot (its enclosing branch is inactive or, for an
    #ElseIf, an earlier branch was taken) does not evaluate the condition, as
    VBA does not: a #Const defined only inside the enclosing branch is then
    never looked up where it does not exist. An environment whose evaluation
    fails is reported and reads the condition as False.
    """
    match = re.match(r"^(.*?)\s+Then\s*$", remainder.strip(), re.IGNORECASE)
    if match is None or not match.group(1).strip():
        state.reject(line, f"indeterminate #{label} directive: #{label} must end with Then "
                           "and contain a condition")
        return {name: False for name in ENVIRONMENTS}
    values = {name: False for name in ENVIRONMENTS}
    errors: list[str] = []
    for name in ENVIRONMENTS:
        if not eligible[name]:
            continue
        try:
            values[name] = evaluate(match.group(1), state.symbols(name))
        except ExpressionError as error:
            if str(error) not in errors:
                errors.append(str(error))
    for message in errors:
        state.reject(line, f"indeterminate #{label} directive: {message}")
    return values


def handle_if(state: State, line: int, remainder: str) -> None:
    eligible = {name: state.active(name) for name in ENVIRONMENTS}
    values = condition_values(state, line, "If", remainder, eligible)
    for name, stack in state.stacks.items():
        selected = eligible[name] and values[name]
        stack.append(Frame(eligible[name], selected, selected))


def handle_elseif(state: State, line: int, remainder: str) -> None:
    if any(not stack for stack in state.stacks.values()):
        state.reject(line, "#ElseIf without #If")
        return
    if any(stack[-1].else_seen for stack in state.stacks.values()):
        state.reject(line, "#ElseIf after #Else")
        return
    eligible = {name: stack[-1].parent_active and not stack[-1].branch_taken
                for name, stack in state.stacks.items()}
    values = condition_values(state, line, "ElseIf", remainder, eligible)
    for name, stack in state.stacks.items():
        frame = stack[-1]
        selected = frame.parent_active and not frame.branch_taken and values[name]
        frame.current_active = selected
        frame.branch_taken = frame.branch_taken or selected


def handle_else(state: State, line: int, remainder: str) -> None:
    if remainder.strip():
        state.reject(line, "#Else must not contain trailing text")
    elif any(not stack for stack in state.stacks.values()):
        state.reject(line, "#Else without #If")
    elif any(stack[-1].else_seen for stack in state.stacks.values()):
        state.reject(line, "duplicate #Else")
    else:
        for stack in state.stacks.values():
            frame = stack[-1]
            frame.current_active = frame.parent_active and not frame.branch_taken
            frame.branch_taken = True
            frame.else_seen = True


def handle_end_if(state: State, line: int, remainder: str) -> None:
    if remainder.strip():
        state.reject(line, "#End If must not contain trailing text")
    elif any(not stack for stack in state.stacks.values()):
        state.reject(line, "#End If without #If")
    else:
        for stack in state.stacks.values():
            stack.pop()


HANDLERS: dict[str, Callable[[State, int, str], None]] = {
    "if": handle_if, "elseif": handle_elseif, "else": handle_else, "endif": handle_end_if,
}


def handle_const(state: State, line: int, body: str) -> None:
    """A module #Const: defined, from here on, in each environment that reaches it."""
    match = CONST_BODY.match(body)
    if match is None:
        state.reject(line, "malformed #Const; write #Const Name = <Boolean expression>")
        return
    name = match.group("name")
    if name.upper() in SYMBOLS:
        state.reject(line, f"#Const {name} redefines a compiler constant")
        return
    duplicate = False
    for environment in ENVIRONMENTS:
        if not state.active(environment):
            continue
        duplicate = duplicate or name.upper() in state.constants[environment]
        try:
            value = evaluate(match.group("expression"), state.symbols(environment))
        except ExpressionError as error:
            state.reject(line, f"indeterminate #Const {name}: {error}")
            return
        state.constants[environment][name.upper()] = value
    if duplicate:
        state.reject(line, f"#Const {name} is defined twice in one environment")


def handle_directive(state: State, line: int, code: str) -> None:
    const = CONST.match(code)
    if const:
        handle_const(state, line, const.group(1))
        return
    directive = DIRECTIVE.match(code)
    if directive is None:
        state.reject(line, f"unsupported conditional-compilation directive: {code}")
        return
    kind = "".join(directive.group(1).split()).casefold()
    HANDLERS[kind](state, line, directive.group(2))


def walk(path: str, text: str) -> tuple[list[tuple[int, int, str, list[str]]], list[Finding]]:
    """``(first line, last line, code, environments that reach it)`` per logical
    code line, and the directive findings, unclosed blocks included."""
    state = State(path)
    units: list[tuple[int, int, str, list[str]]] = []
    for start, end, kind, code in logical_units(text.split("\n")):
        if kind == "directive":
            handle_directive(state, start, code)
        else:
            units.append((start, end, code,
                          [name for name in ENVIRONMENTS if state.active(name)]))
    depth = max(len(stack) for stack in state.stacks.values())
    if depth:
        state.reject(None, f"{depth} conditional-compilation block(s) not closed")
    return units, state.findings


def reachable_sources(path: str, text: str) -> tuple[dict[str, str], list[Finding]]:
    """Each environment's source with inactive lines and directives blanked,
    so that line numbers are preserved; plus the directive findings."""
    lines = text.split("\n")
    sources = {name: [""] * len(lines) for name in ENVIRONMENTS}
    units, findings = walk(path, text)
    for start, end, _, environments in units:
        for name in environments:
            sources[name][start - 1:end] = lines[start - 1:end]
    return {name: "\n".join(source) for name, source in sources.items()}, findings


def analyze(path: str, text: str) -> list[Finding]:
    units, findings = walk(path, text)
    for start, _, code, environments in units:
        reachable = [name for name in VBA7_ENVIRONMENTS if name in environments]
        # Every statement of the line: a Declare may follow a colon.
        for statement in split_statements(code):
            declaration = DECLARE.match(statement)
            if declaration and reachable and not declaration.group("ptrsafe"):
                findings.append(Finding(path, start, "Declare reachable under VBA7 without "
                                                     f"PtrSafe in {', '.join(reachable)}"))
    return findings


def check(root: Path) -> tuple[list[Finding], int]:
    components = tracked_components(root)
    findings: list[Finding] = []
    for path in components:
        findings.extend(analyze(path, read_component(root, path)))
    return findings, len(components)


def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=(__doc__ or "").splitlines()[0])
    parser.add_argument("--root", type=Path, default=ROOT)
    args = parser.parse_args(argv)
    findings, components = check(args.root.resolve())
    return report("vba-conditionals", findings, f"{components} VBA component(s), "
                  f"environments {', '.join(ENVIRONMENTS)}")


if __name__ == "__main__":
    sys.exit(main())
