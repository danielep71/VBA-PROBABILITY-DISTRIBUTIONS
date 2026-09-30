"""Focused fixtures for verification-depth command wiring."""
import contextlib
import copy
import io
import json
import tempfile
from pathlib import Path
from unittest.mock import patch

import check_verification_depth as V

shim = """
commands = (
    ("proof.py",),
    ("proof_mode.py", "--negative", "strict"),
)
"""
found = V._shim_commands(shim)
assert ("proof.py",) in found
assert ("proof_mode.py", "--negative", "strict") in found
assert ("proof_mode.py",) not in found
assert V._command_tuple("python3 proof_mode.py --negative strict") == (
    "proof_mode.py", "--negative", "strict"
)
assert V._command_tuple("proof.py") == ("proof.py",)
print("PASS: verification-depth command wiring fixtures (complete tuples, arguments preserved)")

# Exercise the checker itself against independently damaged policy inputs.
original = json.loads(V.INVENTORY.read_text(encoding="utf-8"))
original_shim = V.SHIM.read_text(encoding="utf-8")


def validate(data, shim_text=original_shim):
    with tempfile.TemporaryDirectory() as tmp:
        inventory = Path(tmp) / "inventory.json"
        shim_path = Path(tmp) / "shim.py"
        inventory.write_text(json.dumps(data), encoding="utf-8")
        shim_path.write_text(shim_text, encoding="utf-8")
        output = io.StringIO()
        with patch.object(V, "INVENTORY", inventory), patch.object(V, "SHIM", shim_path), \
                contextlib.redirect_stdout(output):
            result = V.main()
        return result, output.getvalue()


assert validate(original)[0] == 0
missing = copy.deepcopy(original)
excel = next(c for c in missing["controls"] if c["id"] == "excel-exact-sha-certification")
missing["controls"].remove(excel)
commands = V._shim_commands(original_shim) - {V._command_tuple(excel["proof_command"])}
result, output = validate(missing, "commands = " + repr(tuple(sorted(commands))))
assert result == 1 and "required release-blocking control missing" in output
for state in ("expected_red_until_excel_export", "green_fixture_proof_live_main_stale",
              "fixture_green_live_holdout_unbound_until_excel_export", "unknown"):
    damaged = copy.deepcopy(original)
    damaged["controls"][0]["current_state"] = state
    result, output = validate(damaged)
    assert result == 1 and "unsupported current_state" in output
unwired = copy.deepcopy(original)
unwired["controls"][0]["proof_command"] += " --unwired-negative-mode"
result, output = validate(unwired)
assert result == 1 and "is not wired" in output
print("PASS: verification-depth missing-control, retired-state and unwired-argument mutations")
