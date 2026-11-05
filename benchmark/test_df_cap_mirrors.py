#!/usr/bin/env python3
"""Negative and positive fixtures for check_df_cap_mirrors.py (#86).

Each fixture copies the live src/M_STATS_PROBDIST_TFAMILY.bas and
_contract_eval.py into a temporary root, applies one mutation and requires the
checker to fail with the expected reason: a changed VBA value (including a
one-ulp change), a changed Python value (including a one-ulp change), a missing
or renamed constant on either side, a duplicate on either side (in the module,
in a procedure and in another module), and unparsable or unverifiable
literals and declarations. Equivalent spellings of the same Double, and a
declaration split by a line continuation, must pass. Every mutation asserts
that its target text occurs exactly once, so a fixture cannot pass vacuously
after the source moves. The live repository must pass.

Run: python3 test_df_cap_mirrors.py   (exit 0 = pass, 1 = fail)
"""
from __future__ import annotations

import contextlib
import io
import math
import subprocess
import sys
import tempfile
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
import check_df_cap_mirrors as C  # noqa: E402

BAS = (C.ROOT / C.SOURCE).read_bytes().decode("ascii").replace("\r\n", "\n")
PY = C.PYTHON.read_text(encoding="utf-8")
F_LINE = "Private Const PROB_F_MAX_DF       As Double = 10000000000#"
T_LINE = "Private Const PROB_T_MAX_DF       As Double = 100000000#"
CHI_LINE = "Private Const PROB_CHI_MAX_DF     As Double = 100000000#"
FIRST_PROCEDURE = "\nPublic Function "
failures: list[str] = []


def replace_once(text: str, old: str, new: str) -> str:
    count = text.count(old)
    if count != 1:
        raise SystemExit(f"FAIL: fixture anchor {old!r} occurs {count} times, expected 1")
    return text.replace(old, new)


def insert_in_first_procedure(text: str, line: str) -> str:
    """Insert a statement as the first line of the first public procedure body."""
    start = text.index(FIRST_PROCEDURE) + 1
    end = text.index("\n", start) + 1
    while text[end - 2] == "_":  # header continued on the next physical line
        end = text.index("\n", end) + 1
    return text[:end] + line + "\n" + text[end:]


def run(bas: str = BAS, py: str = PY, others: dict[str, str] | None = None
        ) -> tuple[int, str]:
    with tempfile.TemporaryDirectory() as tmp:
        root = Path(tmp)
        (root / "src").mkdir()
        (root / C.SOURCE).write_bytes(bas.replace("\n", "\r\n").encode("ascii"))
        for name, text in (others or {}).items():
            (root / "src" / name).write_bytes(text.replace("\n", "\r\n").encode("ascii"))
        python = root / "_contract_eval.py"
        python.write_text(py, encoding="utf-8")
        output = io.StringIO()
        with contextlib.redirect_stdout(output):
            code = C.main(["--root", str(root), "--python", str(python)])
        return code, output.getvalue()


def expect_pass(label: str, bas: str = BAS, py: str = PY) -> None:
    code, output = run(bas, py)
    if code != 0:
        failures.append(f"{label}: expected PASS, got:\n{output}")


def expect_fail(label: str, reason: str, bas: str = BAS, py: str = PY,
                others: dict[str, str] | None = None) -> None:
    code, output = run(bas, py, others)
    if code != 1 or reason not in output:
        failures.append(f"{label}: expected FAIL mentioning {reason!r}, got exit {code}:\n"
                        f"{output}")


def vba(old: str, new: str) -> str:
    return replace_once(BAS, old, new)


def python(old: str, new: str) -> str:
    return replace_once(PY, old, new)


ULP_ABOVE_1E8 = math.nextafter(1e8, math.inf)
assert ULP_ABOVE_1E8 == 1e8 + 2.0 ** -26

# --- live repository and the copy harness ---------------------------------------
expect_pass("unmodified copy")
proc = subprocess.run([sys.executable, "check_df_cap_mirrors.py"], cwd=HERE,
                      capture_output=True, text=True)
if proc.returncode != 0:
    failures.append("live repository check did not pass: " + (proc.stdout + proc.stderr))

# --- accepted spellings: the same Double, bit for bit ----------------------------
for spelling in ("1E10", "1e+10#", "1D10", "10000000000", "10000000000.0", "1E10!",
                 "10000000000@", "10000000000.000#"):
    expect_pass(f"VBA F cap spelled {spelling}",
                bas=vba(F_LINE, F_LINE.replace("10000000000#", spelling)))
expect_pass("VBA T cap spelled 100000000& (Long)",
            bas=vba(T_LINE, T_LINE.replace("100000000#", "100000000&")))
expect_pass("declaration split by a line continuation, with a trailing comment",
            bas=vba(F_LINE, "Private Const PROB_F_MAX_DF As Double = _\n"
                            "    10000000000# ' continued"))
expect_pass("Public scope", bas=vba(CHI_LINE, CHI_LINE.replace("Private", "Public")))
expect_pass("Python mirror spelled 1e10",
            py=python("F_MAX_DF = 10000000000.0", "F_MAX_DF = 1e10"))

# --- changed VBA value -----------------------------------------------------------
expect_fail("VBA F cap raised", "differs from PROB_F_MAX_DF",
            bas=vba(F_LINE, F_LINE.replace("10000000000#", "20000000000#")))
expect_fail("VBA T cap one ulp higher", "differs from PROB_T_MAX_DF",
            bas=vba(T_LINE, T_LINE.replace(
                "100000000#", "100000000.0000000149011611938476562500#")))
expect_fail("VBA CHI cap lowered", "differs from PROB_CHI_MAX_DF",
            bas=vba(CHI_LINE, CHI_LINE.replace("100000000#", "10000000#")))

# --- changed Python value --------------------------------------------------------
expect_fail("Python F mirror lowered", "differs from PROB_F_MAX_DF",
            py=python("F_MAX_DF = 10000000000.0", "F_MAX_DF = 1000000000.0"))
expect_fail("Python T mirror one ulp higher", "differs from PROB_T_MAX_DF",
            py=python("T_MAX_DF = 100000000.0", f"T_MAX_DF = {ULP_ABOVE_1E8!r}"))
expect_fail("Python CHI mirror raised", "differs from PROB_CHI_MAX_DF",
            py=python("CHI_MAX_DF = 100000000.0", "CHI_MAX_DF = 1000000000.0"))
expect_fail("Python mirror changed after import", "not its literal",
            py=PY + "\nglobals()['T_MAX_DF'] = 1.0\n")

# --- missing or renamed constant --------------------------------------------------
expect_fail("VBA CHI cap deleted", "PROB_CHI_MAX_DF is not declared",
            bas=vba(CHI_LINE, ""))
expect_fail("VBA T cap renamed", "PROB_T_MAX_DF is not declared",
            bas=vba(T_LINE, T_LINE.replace("PROB_T_MAX_DF", "PROB_T_MAX_DF_V2")))
expect_fail("Python F mirror deleted", "F_MAX_DF is not assigned",
            py=python("F_MAX_DF = 10000000000.0", ""))
expect_fail("Python CHI mirror renamed", "CHI_MAX_DF is not assigned",
            py=python("CHI_MAX_DF = 100000000.0", "CHI_MAX_DF_V2 = 100000000.0"))
with tempfile.TemporaryDirectory() as empty:
    (Path(empty) / "src").mkdir()
    output = io.StringIO()
    with contextlib.redirect_stdout(output):
        if C.main(["--root", empty]) != 1 or "not found" not in output.getvalue():
            failures.append("a root without the VBA module did not fail")

# --- duplicate constant ------------------------------------------------------------
expect_fail("VBA T cap declared twice with the same value", "declared 2 times",
            bas=vba(T_LINE, T_LINE + "\n" + T_LINE))
expect_fail("VBA F cap shadowed in a procedure", "cannot verify",
            bas=insert_in_first_procedure(BAS, "    Const PROB_F_MAX_DF As Double = 1E10"))
expect_fail("VBA CHI cap shadowed by a local variable", "cannot verify",
            bas=insert_in_first_procedure(BAS, "    Dim PROB_CHI_MAX_DF As Double"))
expect_fail("VBA T cap in a multi-constant declaration", "unparsable literal",
            bas=vba(T_LINE, "Private Const PROB_T_MAX_DF As Double = 100000000#, "
                            "PROB_X As Double = 1"))
expect_fail("VBA F cap redeclared in another module", "duplicate declaration of PROB_F_MAX_DF",
            others={"M_OTHER.bas": 'Attribute VB_Name = "M_OTHER"\nOption Explicit\n'
                                   "Public Const PROB_F_MAX_DF As Double = 10000000000#\n"})
expect_fail("VBA T cap as an Enum member in another module", "cannot verify",
            others={"M_OTHER.bas": 'Attribute VB_Name = "M_OTHER"\nOption Explicit\n'
                                   "Private Enum E\n    PROB_T_MAX_DF = 1\nEnd Enum\n"})
expect_fail("Python T mirror assigned twice", "bound 2 times",
            py=python("T_MAX_DF = 100000000.0", "T_MAX_DF = 100000000.0\nT_MAX_DF = 100000000.0"))
expect_fail("Python CHI mirror rebound through global", "global statement",
            py=PY + "\ndef _rebind():\n    global CHI_MAX_DF\n    CHI_MAX_DF = 1.0\n")

# --- unparsable or unverifiable literal ---------------------------------------------
for literal, reason in (
        ("1E10 * 1", "unparsable literal"),
        ("PROB_T_MAX_DF * 100", "unparsable literal"),
        ("CDbl(1E10)", "unparsable literal"),
        ("&H2540BE400", "unparsable literal"),
        ('"1E10"', "unparsable literal"),
        ("", "unparsable literal"),
        ("1E400", "outside the Double range"),
        ("10000000000.1#", "not exactly representable as a binary64"),
        ("1E10%", "integer type suffix"),
        ("10000000000&", "overflows"),
        ("10000000001!", "not exactly representable as a Single"),
        ("10000000000.00001@", "not exactly a Currency value"),
        ("-0#", "negative zero")):
    expect_fail(f"VBA F cap literal {literal!r}", reason,
                bas=vba(F_LINE, F_LINE.replace("10000000000#", literal)))
expect_fail("VBA T cap declared As Single", "not As Double",
            bas=vba(T_LINE, T_LINE.replace("As Double", "As Single")))
expect_fail("VBA CHI cap without a declared type", "cannot verify",
            bas=vba(CHI_LINE, "Private Const PROB_CHI_MAX_DF = 100000000#"))
expect_fail("VBA CHI cap without Private or Public", "cannot verify",
            bas=vba(CHI_LINE, CHI_LINE.replace("Private ", "")))
expect_fail("VBA F cap inside conditional compilation", "conditional compilation",
            bas=vba(F_LINE, "#If VBA7 Then\n" + F_LINE + "\n#End If"))
expect_fail("Python F mirror as an expression", "finite float literal",
            py=python("F_MAX_DF = 10000000000.0", "F_MAX_DF = 10.0 ** 10"))
expect_fail("Python T mirror as an int", "finite float literal",
            py=python("T_MAX_DF = 100000000.0", "T_MAX_DF = 100000000"))
expect_fail("Python module unparsable", "unreadable", py=PY + "\ndef broken(:\n")

# --- direct conversion table --------------------------------------------------------
for literal, expected in (("100000000#", "4197d78400000000"),
                          ("10000000000#", "4202a05f20000000"),
                          (".5", "3fe0000000000000"), ("5.", "4014000000000000"),
                          ("-2.5#", "c004000000000000"), ("1E-1074", None),
                          ("4.9406564584124654E-324", None)):
    try:
        got = C.bits(C.vba_double(literal))
    except C.CapError:
        got = None
    if got != expected:
        failures.append(f"vba_double({literal!r}) gave bits {got}, expected {expected}")
# The exact minimum subnormal, written out, converts; its 17-digit decimal does not.
exact_min = format(2.0 ** -1074, ".1074f").rstrip("0")
if C.bits(C.vba_double(exact_min)) != "0000000000000001":
    failures.append("the exact decimal of the minimum subnormal did not convert")

if failures:
    print("FAIL: df-cap mirror fixtures")
    for failure in failures:
        print("  - " + failure)
    raise SystemExit(1)
print("PASS: df-cap mirror fixtures (changed VBA or Python values, one-ulp drift, missing, "
      "renamed and duplicate constants, and unparsable or inexact literals fail; equivalent "
      "spellings and line continuations pass)")
