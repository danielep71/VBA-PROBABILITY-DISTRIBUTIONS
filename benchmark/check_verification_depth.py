#!/usr/bin/env python3
"""Fail closed when a release-blocking control lacks executable negative proof.

The inventory is intentionally about assurance controls, not every benchmark
utility. Each listed control must name a live checker, an executable proof
command that contains meaningful negative cases, and its current repository
state. Proof commands must also be wired into test_evidence_tools.py so they run
*before* the strict numerical gate; otherwise an intentionally red Phase-0 gate
could skip the very tests meant to establish that the gate machinery works.
"""
from __future__ import annotations

import ast
import json
import shlex
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
INVENTORY = HERE / "verification_depth.json"
SHIM = HERE / "test_evidence_tools.py"
ALLOWED_STATES = {
    "green",
    "green_checker",
    "green_fixture_proof",
    "green_fixture_proof_live_main_stale",
    "green_with_registered_transition_debt",
    "expected_red_until_excel_export",
    "fixture_green_live_holdout_unbound_until_excel_export",
}


def _script_from_command(command: str) -> str:
    parts = shlex.split(command)
    if not parts:
        raise ValueError("empty command")
    if parts[0] in {"python", "python3"}:
        if len(parts) < 2:
            raise ValueError(f"python command has no script: {command!r}")
        return parts[1]
    return parts[0]


def _command_tuple(command: str) -> tuple[str, ...]:
    parts = tuple(shlex.split(command))
    if not parts:
        raise ValueError("empty command")
    if parts[0] in {"python", "python3"}:
        if len(parts) < 2:
            raise ValueError(f"python command has no script: {command!r}")
        parts = parts[1:]
    return parts


def _shim_commands(text: str) -> set[tuple[str, ...]]:
    """Extract complete constant command tuples from `commands = (...)`."""
    tree = ast.parse(text)
    found: set[tuple[str, ...]] = set()
    for node in ast.walk(tree):
        if not isinstance(node, ast.Assign):
            continue
        if not any(isinstance(t, ast.Name) and t.id == "commands" for t in node.targets):
            continue
        if not isinstance(node.value, (ast.Tuple, ast.List)):
            continue
        for item in node.value.elts:
            if not isinstance(item, (ast.Tuple, ast.List)) or not item.elts:
                continue
            values = []
            for element in item.elts:
                if not isinstance(element, ast.Constant) or not isinstance(element.value, str):
                    values = []
                    break
                values.append(element.value)
            if values:
                found.add(tuple(values))
    return found


def main() -> int:
    failures: list[str] = []
    try:
        data = json.loads(INVENTORY.read_text(encoding="utf-8"))
    except Exception as exc:
        print(f"FAIL: verification-depth inventory unreadable: {exc}")
        return 1

    if data.get("schema_version") != 1:
        failures.append("schema_version must be 1")
    controls = data.get("controls")
    if not isinstance(controls, list) or not controls:
        failures.append("controls must be a non-empty list")
        controls = []

    try:
        shim_commands = _shim_commands(SHIM.read_text(encoding="utf-8"))
    except Exception as exc:
        failures.append(f"cannot parse test_evidence_tools.py: {exc}")
        shim_commands = set()

    seen: set[str] = set()
    for index, control in enumerate(controls, 1):
        if not isinstance(control, dict):
            failures.append(f"control #{index} is not an object")
            continue
        cid = str(control.get("id", "")).strip()
        if not cid:
            failures.append(f"control #{index} has no id")
            continue
        if cid in seen:
            failures.append(f"duplicate control id: {cid}")
        seen.add(cid)
        for field in ("live_command", "proof_command", "negative_claim", "current_state"):
            if not isinstance(control.get(field), str) or not control[field].strip():
                failures.append(f"{cid}: missing {field}")
        if control.get("current_state") not in ALLOWED_STATES:
            failures.append(f"{cid}: unsupported current_state {control.get('current_state')!r}")
        if len(str(control.get("negative_claim", "")).strip()) < 24:
            failures.append(f"{cid}: negative_claim is too vague")

        for field in ("live_command", "proof_command"):
            command = control.get(field)
            if not isinstance(command, str) or not command.strip():
                continue
            try:
                script = _script_from_command(command)
            except ValueError as exc:
                failures.append(f"{cid}: {exc}")
                continue
            path = HERE / script
            if not path.is_file():
                failures.append(f"{cid}: {field} script does not exist: {script}")

        proof = control.get("proof_command")
        if isinstance(proof, str) and proof.strip():
            try:
                proof_tuple = _command_tuple(proof)
            except ValueError as exc:
                failures.append(f"{cid}: {exc}")
                proof_tuple = ()
            if proof_tuple and proof_tuple not in shim_commands:
                failures.append(
                    f"{cid}: complete proof command {shlex.join(proof_tuple)} "
                    "is not wired into test_evidence_tools.py"
                )

    # These are the minimum v1.0.0 controls. Removing one from the JSON must be
    # an explicit policy change, not an accidental edit that makes the matrix green.
    required = {
        "contract-evaluator", "main-grid-coverage", "root-readme-integrity",
        "incomplete-gamma-parity", "student-t-coefficient-integrity",
        "public-api-drift", "manifest-content-binding", "manifest-provenance",
        "reference-helper-degradation", "generated-contract-table",
        "source-threshold-single-source", "holdout-analyzer-semantics",
        "excel-exact-sha-certification",
    }
    for cid in sorted(required - seen):
        failures.append(f"required release-blocking control missing from inventory: {cid}")

    if failures:
        print("FAIL: verification-depth contract")
        for failure in failures:
            print("  - " + failure)
        return 1
    print(
        f"PASS: verification-depth contract ({len(controls)} release-blocking controls; "
        "every negative proof exists and runs before the strict gate)"
    )
    return 0


if __name__ == "__main__":
    sys.exit(main())
