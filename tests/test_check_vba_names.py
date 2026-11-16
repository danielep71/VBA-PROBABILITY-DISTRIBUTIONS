#!/usr/bin/env python3
"""Fixtures for the VBA identifier and call checks (#291).

Each rule has a fixture that the check must report: an undeclared assignment, a
call to a missing procedure, a keyword used as a name (``Scale`` among them), a
call with too many or too few arguments, a named argument that does not exist
and a qualified call to a ``Private`` procedure. The valid fixtures hold the
constructs the checks must not report.

The fixtures and the cases are those of tests/test_check_vba.py of
danielep71/K-TEMPLATE at 9fef90ddd20b8f8c43b39b964a63ed186079b7a4, without its
form, class and document-module cases: this repository tracks standard modules
only. The cases for ``Scale``, for the accepted name and for this repository's
own source are added here.
"""

from __future__ import annotations

import contextlib
import io
import shutil
import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "tools"))
sys.path.insert(0, str(ROOT / "tests"))

import check_vba_calls  # noqa: E402
import check_vba_identifiers  # noqa: E402
from test_check_vba import VALID, repository, run_check  # noqa: E402

FIXTURE_PATH = "tests/TEST_Fixture.bas"
OTHER_PATH = "tests/TEST_Other.bas"
CALLEE_PATH = "tests/TEST_Callee.bas"


def module(body: str) -> str:
    return f'Attribute VB_Name = "TEST_Fixture"\nOption Explicit\nOption Private Module\n{body}'


def standalone(source: str) -> dict[str, str | bytes]:
    """One test module holding ``source``."""
    return {FIXTURE_PATH: source}


IDENTIFIERS_VALID = {
    "declarations-of-every-kind": """Private mCount As Long, mNames() As String
Private Const LIMIT As Long = 3
Private Enum Slot
    [_First] = 0
    slA = 1
End Enum

Public Function Total( _
    ByVal Value As Long, _
    Optional ByRef Label As String = "a, b", _
    Optional ByVal Extra As Variant) _
    As Long

        Dim i As Long, Items As Collection
        Static Calls As Long
        Const STEP_SIZE As Long = 1
        ReDim Buffer(1 To LIMIT) As Long
        ReDim Preserve Buffer(1 To LIMIT + 1), Spare$(2)
        ReDim Items.Values(1 To 3), Items!Keys(1)
        ReDim Ratios!(1 To 3)
        Ratios!(1) = 1
        Items!Count = 3
        Call Items(1).Refresh
        Call Items!First.Refresh
        Items(1).Refresh
        With Items
            ReDim .Values(1 To 4)
        End With
        Dim Fixed As String * 8, Range2 As Excel.Range, Made As New Collection
        Set Items = New Collection
        For i = 1 To LIMIT Step STEP_SIZE: Buffer(i) = i: Next i
        For Each Extra In Items
        Next Extra
        If Value > 0 Then mCount = slA Else mCount = LIMIT
        If Value > 1 Then mCount = 1: mCount = 2 Else mCount = 3: Calls = 0
        Dim Settings As Long, Letter As Long, Reset As Long
        Settings = 1: Letter = 2
        Set Items = Nothing
        Mid$(Label, 1, 1) = "z"
        Calls = Calls + 1
        Total = Value
End Function
""",
    "calls-of-every-kind": """Public Sub Run()
        Helper 1
        Call Helper(2)
        Call TEST_Other.Thing(3)
        TEST_Other.Thing 4
        Debug.Print "Undeclared = 1"
        Err.Raise 5
        MsgBox "x"
        DoEvents
        Shared
Retry:
        On Error GoTo Retry
End Sub

Private Sub Helper(ByVal Value As Long)
        With Application
            .StatusBar = Value
        End With
        Application.StatusBar = False
End Sub
""",
    "a-declaration-in-each-active-branch": """Public Sub Run()
#If VBA7 Then
        Dim Handle As LongPtr
#Else
        Dim Handle As Long
#End If
        Handle = 0
End Sub
""",
}
IDENTIFIERS_INVALID = {
    "undeclared-assignment": ("Public Sub Run()\n        Total = 1\nEnd Sub\n",
                              "Total is assigned in Run but not declared"),
    "undeclared-set": ("Public Sub Run()\n        Set Items = Nothing\nEnd Sub\n",
                       "Items is assigned in Run but not declared"),
    "undeclared-loop-counter": ("Public Sub Run()\n        For i = 1 To 2\n        Next i\n"
                                "End Sub\n", "i is assigned in Run but not declared"),
    "assignment-after-then": ("Public Sub Run()\n        If True Then Missing = 1\nEnd Sub\n",
                              "Missing is assigned in Run but not declared"),
    "assignment-after-else-past-a-colon": (
        "Public Sub Run()\n        Dim x As Long\n"
        "        If True Then x = 1: x = 2 Else Missing = 3\nEnd Sub\n",
        "Missing is assigned in Run but not declared"),
    "call-after-else-past-a-colon": (
        "Public Sub Run()\n        Dim x As Long\n"
        "        If True Then x = 1: x = 2 Else Gone 3\nEnd Sub\n",
        "Gone is called in Run"),
    "a-declaration-only-in-an-inactive-branch": (
        "Public Sub Run()\n#If Not VBA7 Then\n        Dim Legacy As Long\n#End If\n"
        "        Legacy = 1\nEnd Sub\n",
        "Legacy is assigned in Run but not declared [vba7-win32, vba7-win64]"),
    "a-renamed-private-helper": (
        "Public Sub Run()\n        OldHelper 1\nEnd Sub\n\n"
        "Private Sub NewHelper(ByVal Value As Long)\nEnd Sub\n",
        "OldHelper is called in Run but no procedure of that name is reachable"),
    "an-explicit-call-to-nothing": ("Public Sub Run()\n        Call Gone(1)\nEnd Sub\n",
                                    "Gone is called in Run"),
    "an-explicit-call-without-arguments": ("Public Sub Run()\n        Call Gone\nEnd Sub\n",
                                           "Gone is called in Run"),
    "dim-empty": ("Public Sub Run()\n        Dim Empty() As Long\nEnd Sub\n",
                  "variable Empty is named like the VBA keyword Empty"),
    "keyword-parameter": ("Public Sub Run(ByVal Date As Long)\nEnd Sub\n",
                          "parameter Date is named like the VBA keyword Date"),
    "keyword-module-name": ("Private Name As String\n",
                            "variable Name is named like the VBA keyword Name"),
    "keyword-enum-member": ("Private Enum Slot\n    Error = 1\nEnd Enum\n",
                            "Enum member Error is named like the VBA keyword Error"),
    "unreadable-parameters": ("Public Sub Run(ByVal 1st As Long)\nEnd Sub\n",
                              "cannot read the parameter"),
    "unreadable-declaration": ("Public Sub Run()\n        Dim 2nd As Long\nEnd Sub\n",
                               "cannot read the declaration"),
    "unreadable-redim": ("Public Sub Run()\n        ReDim 3rd(1)\nEnd Sub\n",
                         "cannot read the declaration"),
    "single-line-if-without-then": ("Public Sub Run()\n        If True x = 1\nEnd Sub\n",
                                    "cannot read the If statement"),
}
OTHER = """Attribute VB_Name = "TEST_Other"
Option Explicit
Option Private Module

Public Shared As Long

Public Sub Thing(ByVal Value As Long)
End Sub

Public Sub Shared()
End Sub
"""


@unittest.skipIf(shutil.which("git") is None, "git is not installed")
class IdentifierTest(unittest.TestCase):
    def test_valid_fixtures_pass(self) -> None:
        self.assertEqual(run_check(check_vba_identifiers.check, VALID), [])
        for name, body in IDENTIFIERS_VALID.items():
            with self.subTest(name=name):
                files = {**standalone(module(body)), OTHER_PATH: OTHER}
                self.assertEqual(run_check(check_vba_identifiers.check, files), [])

    def test_invalid_fixtures_fail(self) -> None:
        for name, (body, text) in IDENTIFIERS_INVALID.items():
            with self.subTest(name=name):
                findings = run_check(check_vba_identifiers.check, standalone(module(body)))
                self.assertTrue(any(text in item for item in findings), findings)

    def test_public_names_of_other_standard_modules_resolve(self) -> None:
        caller = module("Public Sub Run()\n        Shared = 1\n        Thing 2\nEnd Sub\n")
        files = {**standalone(caller), OTHER_PATH: OTHER}
        self.assertEqual(run_check(check_vba_identifiers.check, files), [])
        files[OTHER_PATH] = OTHER.replace("Public Sub Thing", "Private Sub Thing")
        self.assertEqual(run_check(check_vba_identifiers.check, files),
                         [FIXTURE_PATH + ":6: Thing is called in Run but no procedure of that "
                          "name is reachable [vba6-win32, vba7-win32, vba7-win64]"])

    def test_a_comma_inside_a_date_literal_separates_no_declaration(self) -> None:
        # Valid VBA that the template's splitter reported as unreadable.
        body = ("Private Const START_DAY As Date = #January 1, 2020#, LIMIT As Long = 3\n\n"
                "Public Sub Run(Optional ByVal Stamp As Date = #January 1, 2020#, "
                "Optional ByVal Count As Long = 2)\n"
                "        Dim First As Date, Rate#, Total#\n"
                "        First = #1/2/2020#: Rate# = 1#: Total# = Rate# + 2#\n"
                "End Sub\n")
        self.assertEqual(run_check(check_vba_identifiers.check, standalone(module(body))), [])
        # The names after the literal are still declared, and still checked.
        keyword = body.replace("Optional ByVal Count As Long = 2", "Optional ByVal Date As Long = 2")
        findings = run_check(check_vba_identifiers.check, standalone(module(keyword)))
        self.assertEqual(len(findings), 1, findings)
        self.assertIn("parameter Date is named like the VBA keyword Date", findings[0])
        undeclared = body.replace("Dim First As Date, ", "Dim ")
        findings = run_check(check_vba_identifiers.check, standalone(module(undeclared)))
        self.assertEqual(len(findings), 1, findings)
        self.assertIn("First is assigned in Run but not declared", findings[0])

    def test_scale_is_refused_as_a_name(self) -> None:
        # The name that stopped the project from compiling in Excel (pull request 242).
        cases = {
            "parameter": ("Public Sub Run(ByVal Scale As Double)\nEnd Sub\n", "parameter Scale"),
            "variable": ("Public Sub Run()\n        Dim Scale As Double\nEnd Sub\n",
                         "variable Scale"),
            "procedure": ("Public Sub Scale()\nEnd Sub\n", "Scale is named like"),
        }
        for name, (body, text) in cases.items():
            with self.subTest(name=name):
                findings = run_check(check_vba_identifiers.check, standalone(module(body)))
                self.assertEqual(len(findings), 1, findings)
                self.assertIn(text, findings[0])
                self.assertIn("named like the VBA keyword Scale", findings[0])
        allowed = "Public Sub Run(ByVal ScaleParam As Double)\nEnd Sub\n"
        self.assertEqual(run_check(check_vba_identifiers.check, standalone(module(allowed))), [])

    def test_an_accepted_name_is_accepted_only_in_its_file(self) -> None:
        body = module("Public Sub Run()\n        Dim Base As Double\n        Base = 1\nEnd Sub\n")
        findings = run_check(check_vba_identifiers.check, standalone(body))
        self.assertEqual(len(findings), 1, findings)
        self.assertIn("variable Base is named like the VBA keyword Base", findings[0])
        for (path, name), reason in check_vba_identifiers.ACCEPTED_KEYWORD_NAMES.items():
            with self.subTest(path=path, name=name):
                self.assertTrue((ROOT / path).is_file(), path)
                self.assertEqual(name, name.casefold())
                self.assertIn(name, check_vba_identifiers.KEYWORDS)
                self.assertGreater(len(reason), 40)
                accepted = run_check(check_vba_identifiers.check, {path: body})
                self.assertEqual(accepted, [])
                other = body.replace("Dim Base", "Dim Empty").replace("Base = 1", "")
                self.assertEqual(len(run_check(check_vba_identifiers.check, {path: other})), 1)

    def test_every_accepted_name_is_still_needed(self) -> None:
        # An entry whose name has left the source must leave the list too.
        for path, name in check_vba_identifiers.ACCEPTED_KEYWORD_NAMES:
            text = (ROOT / path).read_text(encoding="ascii", errors="surrogateescape")
            declared = [line for line in text.splitlines()
                        if line.lstrip().lower().startswith("dim ")
                        and f" {name} as " in f" {line.lower()} ".replace(",", " , ")]
            self.assertTrue(declared, f"{name} is no longer declared in {path}")

    def test_repository_vba_has_no_findings(self) -> None:
        findings, components, procedures = check_vba_identifiers.check(ROOT)
        self.assertEqual([str(item) for item in findings], [])
        self.assertGreater(components, 6)
        self.assertGreater(procedures, 100)

    def test_main_exits_1_on_a_finding(self) -> None:
        bad = module("Public Sub Run()\n        Total = 1\nEnd Sub\n")
        good = module("Public Sub Run()\n        Dim Total As Long\n        Total = 1\nEnd Sub\n")
        for source, expected in ((good, 0), (bad, 1)):
            with self.subTest(expected=expected), repository(standalone(source)) as root, \
                    contextlib.redirect_stdout(io.StringIO()) as output:
                self.assertEqual(check_vba_identifiers.main(["--root", str(root)]), expected)
                self.assertIn("procedure(s); static check", output.getvalue())


CALLEE = """Attribute VB_Name = "TEST_Callee"
Option Explicit
Option Private Module

Public Sub Pair(ByVal First As Long, ByVal Second As Long)
End Sub

Public Function Scaled(ByVal Value As Double, Optional ByVal Factor As Double = 1) As Double
        Scaled = Value * Factor
End Function

Public Sub Collect(ByVal Label As String, ParamArray Values() As Variant)
End Sub

Public Function Grid() As Variant
        Grid = Array(1, 2)
End Function

Public Property Get Size(ByVal Index As Long) As Long
        Size = Index
End Property

Private Sub Hidden(ByVal Value As Long)
End Sub

Private Function Secret() As Long
End Function
"""
CALLS_VALID = """
Private Totals(1 To 3) As Double


Public Sub Run()
        Dim Values(1 To 2) As Double
        Dim Items As Collection
        Dim Pair(1 To 2) As Long
        Pair(1) = 1
        TEST_Callee.Pair 1, 2
        Call TEST_Callee.Pair(1, Second:=2)
        Debug.Print Scaled(1), Scaled(1, 2), TEST_Callee.Scaled(Value:=1, Factor:=2)
        Collect "a, b"
        Collect "a", 1, , "c", 3
        Call Collect(Label:="a")
        Values(1) = Scaled(Values(2))
        Totals(1) = Values(1)
        If Scaled(1) > 0 Then Collect "x", 1 Else Helper 1
        Debug.Print Grid(1), TEST_Callee.Size(1, 2)
        Set Items = New Collection
        With Items
            .Add Scaled(2), "Key"
            .Add 1, 2, 3, 4, 5
            .Scaled = 1
        End With
        Items.Scaled 1, 2, 3
        Helper Value:=1
        TEST_Fixture.Helper 1
        Debug.Print "Scaled(1, 2, 3)"
End Sub


Private Sub Helper(ByVal Value As Long)
End Sub
"""
CALLS_INVALID = {
    "too-few": ("TEST_Callee.Pair 1", "call to TEST_Callee.Pair in Run: required argument(s) "
                                      "Second not passed"),
    "too-many": ("Pair 1, 2, 3", "call to TEST_Callee.Pair in Run: 3 argument(s) passed, at "
                                 "most 2 accepted"),
    "omitted-required": ('Pair , "2"', "required argument(s) First not passed"),
    "too-many-in-an-expression": ("Debug.Print Scaled(1, 2, 3)",
                                  "call to TEST_Callee.Scaled in Run: 3 argument(s) passed"),
    "too-few-nested": ("Debug.Print Scaled(Scaled())",
                       "call to TEST_Callee.Scaled in Run: required argument(s) Value"),
    "unknown-named": ("Call Pair(1, Second:=2, Third:=3)", "no parameter is named Third"),
    "paramarray-without-its-required-argument": ("Collect", "required argument(s) Label"),
    "private-of-another-module": ("TEST_Callee.Hidden 1", "call to TEST_Callee.Hidden in Run: "
                                                          "Hidden is Private in TEST_Callee"),
    "private-in-an-expression": ("Debug.Print TEST_Callee.Hidden(1)",
                                 "Hidden is Private in TEST_Callee"),
    "private-without-parameters": ("Debug.Print TEST_Callee.Secret()",
                                   "Secret is Private in TEST_Callee"),
}


@unittest.skipIf(shutil.which("git") is None, "git is not installed")
class CallTest(unittest.TestCase):
    def files(self, body: str) -> dict[str, str | bytes]:
        return {**standalone(module(body)), CALLEE_PATH: CALLEE}

    def test_valid_calls_pass(self) -> None:
        self.assertEqual(run_check(check_vba_calls.check, self.files(CALLS_VALID)), [])
        self.assertEqual(run_check(check_vba_calls.check, VALID), [])

    def test_invalid_calls_are_reported(self) -> None:
        for name, (statement, text) in CALLS_INVALID.items():
            with self.subTest(name=name):
                body = f"\nPublic Sub Run()\n        {statement}\nEnd Sub\n"
                findings = run_check(check_vba_calls.check, self.files(body))
                self.assertEqual(len(findings), 1, findings)
                self.assertIn(text, findings[0])
                self.assertTrue(findings[0].startswith(FIXTURE_PATH + ":6: "))

    def test_date_literals_and_type_suffixes(self) -> None:
        # A comma inside a date literal separates no arguments; a # type
        # suffix opens no date literal, so the arguments after it still count.
        valid = ("\nPublic Sub Run()\n        Dim Amount As Double\n"
                 "        Pair #January 1, 2020#, 2\n        Pair Amount#, 2#\n"
                 "        Debug.Print Scaled(#1/2/2020#, 1#)\nEnd Sub\n")
        self.assertEqual(run_check(check_vba_calls.check, self.files(valid)), [])
        invalid = valid.replace("Pair Amount#, 2#", "Pair Amount#, 2#, #3/4/2020#")
        findings = run_check(check_vba_calls.check, self.files(invalid))
        self.assertEqual(len(findings), 1, findings)
        self.assertIn("3 argument(s) passed, at most 2 accepted", findings[0])

    def test_redim_bounds_are_expressions(self) -> None:
        # The bounds of a ReDim are evaluated at run time; its target is an
        # array, even when it is named like a procedure or indexed past the
        # parameters of one.
        valid = ("\nPublic Sub Run()\n        Dim Values() As Double\n        Dim Pair() As Long\n"
                 "        ReDim Values(1 To Scaled(2, 3))\n"
                 "        ReDim Preserve Values(Scaled(1), 0 To TEST_Callee.Scaled(4)), "
                 "Pair(1, 2, 3)\nEnd Sub\n")
        self.assertEqual(run_check(check_vba_calls.check, self.files(valid)), [])
        cases = {"ReDim Values(Scaled(1, 2, 3))": "3 argument(s) passed, at most 2 accepted",
                 "ReDim Preserve Values(1 To 2, 0 To TEST_Callee.Hidden(1))":
                     "Hidden is Private in TEST_Callee"}
        for statement, text in cases.items():
            with self.subTest(statement=statement):
                body = ("\nPublic Sub Run()\n        Dim Values() As Double\n"
                        f"        {statement}\nEnd Sub\n")
                findings = run_check(check_vba_calls.check, self.files(body))
                self.assertEqual(len(findings), 1, findings)
                self.assertIn(text, findings[0])

    def test_with_block_members_are_not_calls(self) -> None:
        body = ("\nPublic Sub Run()\n        With New Collection\n            .Pair = 1\n"
                "            .Pair 1, 2, 3\n            .Hidden\n"
                "            Debug.Print .Scaled(1, 2, 3)\n        End With\nEnd Sub\n")
        self.assertEqual(run_check(check_vba_calls.check, self.files(body)), [])

    def test_repository_vba_has_no_findings(self) -> None:
        findings, components, resolved = check_vba_calls.check(ROOT)
        self.assertEqual(findings, [])
        if components:
            self.assertGreater(resolved, 0)

    def test_main_exits_1_on_a_finding(self) -> None:
        body = "\nPublic Sub Run()\n        TEST_Callee.Hidden 1\nEnd Sub\n"
        for files, expected in ((self.files(CALLS_VALID), 0), (self.files(body), 1)):
            with self.subTest(expected=expected), repository(files) as root, \
                    contextlib.redirect_stdout(io.StringIO()) as output:
                self.assertEqual(check_vba_calls.main(["--root", str(root)]), expected)
                self.assertIn("resolved call(s); static check", output.getvalue())


if __name__ == "__main__":
    unittest.main()
