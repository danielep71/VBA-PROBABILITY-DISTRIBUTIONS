"""Focused fixtures for verification-depth command wiring."""
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
