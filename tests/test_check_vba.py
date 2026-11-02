"""Positive and negative fixtures for the VBA source checks.

Covers tools/check_vba_source.py, tools/check_vba_conditionals.py,
tools/check_vba_visibility.py and their shared lexer, tools/vba_lexer.py. Each
fixture is a throw-away Git repository holding synthetic VBA components,
written with CRLF as the VBE exports them and staged under this repository's
.gitattributes. The components exist only to exercise the checks: they compute
nothing and stand for no library routine.
"""

from __future__ import annotations

import contextlib
import io
import shutil
import subprocess
import sys
import tempfile
import unittest
from collections.abc import Callable, Iterator, Mapping
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "tools"))

import check_vba_conditionals  # noqa: E402
import check_vba_source  # noqa: E402
import check_vba_visibility  # noqa: E402
from vba_lexer import logical_units, split_statements, strip_code  # noqa: E402

CORE_PATH = "src/M_STATS_PROBDIST_CORE.bas"
SPECIAL_PATH = "src/M_STATS_PROBDIST_SPECIALFUNCS.bas"
FAMILY_PATH = "src/M_STATS_PROBDIST_NORMALFAMILY.bas"
TEST_PATH = "tests/M_STATS_PROBDIST_TEST.bas"
EXPORT_PATH = "benchmark/study/M_STATS_PROBDIST_EXPORT.bas"

CORE = '''Attribute VB_Name = "M_STATS_PROBDIST_CORE"
Option Explicit
Option Private Module

'Internal numerical layer.
Public Const PROB_LIMIT As Double = 1E+300  'Synthetic guard

'Switch private to this module.
#Const PROB_WRITE_STATUS_BAR = False


Public Function PROB_Add( _
    ByVal A As Double, _
    ByVal B As Double) _
    As Double

        PROB_Add = A + B
End Function


Private Sub PROB_Status(ByVal Message As String)
#If PROB_WRITE_STATUS_BAR Then
        Application.StatusBar = Message
#End If
End Sub
'''
SPECIAL = '''Attribute VB_Name = "M_STATS_PROBDIST_SPECIALFUNCS"
Option Explicit
Option Private Module

Public Enum PROB_RatioClass
    PROB_RATIO_ZERO = 0
    PROB_RATIO_POSITIVE = 1
End Enum

Private Type PROB_Pair
    Hi As Double
    Lo As Double
End Type


Public Function PROB_Twice(ByVal X As Double) As Double
    Dim Pair                As PROB_Pair       'Local workspace
    Const SCALE_FACTOR      As Double = 2      'Local constant
        Pair.Hi = X
        PROB_Twice = SCALE_FACTOR * Pair.Hi
End Function
'''
FAMILY = '''Attribute VB_Name = "M_STATS_PROBDIST_NORMALFAMILY"
Option Explicit

Private Const NORMAL_NAME   As String = "Declare Function Bad Lib ""k"" () As Long"
Private mCalls              As Long            'Synthetic counter

'Declare Function Commented Lib "kernel32" () As Long
' A comment that continues _
  Private Declare Function Continued Lib "kernel32" () As Long
Rem Declare Function Remarked Lib "kernel32" () As Long


Public Function K_STATS_FIXTURE( _
    ByVal X As Double) _
    As Variant

        mCalls = mCalls + 1
        K_STATS_FIXTURE = PROB_Add(X, 1)
End Function
'''
TEST = '''Attribute VB_Name = "M_STATS_PROBDIST_TEST"
Option Explicit

Public Sub Test_RunAll()
        Debug.Print "It's ""quoted"": Public Sub PROB_Add", ['Sheet 1'!A1]
        Rem A comment with an open " quote
End Sub

Private Function Helper() As Long
        Helper = 1
End Function
'''
EXPORT = '''Attribute VB_Name = "M_STATS_PROBDIST_EXPORT"
Option Explicit

Public Sub Export_Study()
    Dim Bitness             As String          'Office bitness
    #If Win64 Then
        Bitness = "64"
    #Else
        Bitness = "32"
    #End If
        Debug.Print Bitness
End Sub

Private Function Helper() As Long
        Helper = 2
End Function
'''
VALID: dict[str, str | bytes] = {
    CORE_PATH: CORE,
    SPECIAL_PATH: SPECIAL,
    FAMILY_PATH: FAMILY,
    TEST_PATH: TEST,
    EXPORT_PATH: EXPORT,
}


@contextlib.contextmanager
def repository(files: Mapping[str, str | bytes], *, attributes: bool = True) -> Iterator[Path]:
    """A Git repository holding ``files``: text written with CRLF, then staged."""
    with tempfile.TemporaryDirectory(prefix="vba-check-") as temporary:
        root = Path(temporary)
        subprocess.run(["git", "init", "-q", str(root)], check=True)
        if attributes:
            shutil.copyfile(ROOT / ".gitattributes", root / ".gitattributes")
        for relative, content in files.items():
            path = root / relative
            path.parent.mkdir(parents=True, exist_ok=True)
            if isinstance(content, bytes):
                data = content
            elif relative.endswith((".bas", ".cls", ".frm")):
                data = content.replace("\n", "\r\n").encode("ascii", errors="surrogateescape")
            else:
                data = content.encode("utf-8")
            path.write_bytes(data)
        subprocess.run(["git", "-C", str(root), "-c", "core.autocrlf=false", "-c",
                        "core.safecrlf=false", "add", "--all"], check=True)
        yield root


def run_check(check: Callable[[Path], tuple[Any, ...]], files: Mapping[str, str | bytes],
              *, attributes: bool = True) -> list[str]:
    with repository(files, attributes=attributes) as root:
        return [str(finding) for finding in check(root)[0]]


def changed(path: str, old: str, new: str, files: Mapping[str, str | bytes] = VALID
            ) -> dict[str, str | bytes]:
    text = files[path]
    assert isinstance(text, str) and old in text, (path, old)
    return {**files, path: text.replace(old, new, 1)}


def renamed(source: str, name: str) -> str:
    return source.replace('VB_Name = "M_STATS_PROBDIST_TEST"', f'VB_Name = "{name}"', 1)


def module(body: str, name: str = "M_STATS_PROBDIST_FIXTURE") -> str:
    return f'Attribute VB_Name = "{name}"\nOption Explicit\n{body}'


def standalone(body: str) -> dict[str, str | bytes]:
    """The valid project plus one test module holding ``body``."""
    return {**VALID, "tests/M_STATS_PROBDIST_FIXTURE.bas": module(body)}


class Assertions(unittest.TestCase):
    check: Callable[[Path], tuple[Any, ...]]

    def findings(self, files: Mapping[str, str | bytes], *, attributes: bool = True
                 ) -> list[str]:
        return run_check(type(self).check, files, attributes=attributes)

    def assertFinding(self, files: Mapping[str, str | bytes], text: str) -> None:
        findings = self.findings(files)
        self.assertTrue(any(text in item for item in findings), findings)

    def assertClean(self, files: Mapping[str, str | bytes]) -> None:
        self.assertEqual(self.findings(files), [])


@unittest.skipIf(shutil.which("git") is None, "git is not installed")
class SourceTest(Assertions):
    check = staticmethod(check_vba_source.check)

    def test_valid_project_passes(self) -> None:
        self.assertClean(VALID)

    def test_option_private_module_set_is_the_two_internal_layers(self) -> None:
        self.assertEqual(check_vba_source.OPTION_PRIVATE_REQUIRED,
                         frozenset({CORE_PATH, SPECIAL_PATH}))

    def test_component_outside_its_homes(self) -> None:
        source = renamed(TEST, "M_STATS_PROBDIST_STRAY")
        for path in ("M_STATS_PROBDIST_STRAY.bas", "tools/M_STATS_PROBDIST_STRAY.bas",
                     "src/sub/M_STATS_PROBDIST_STRAY.bas", "tests/x/M_STATS_PROBDIST_STRAY.bas"):
            with self.subTest(path=path):
                self.assertFinding({**VALID, path: source}, "outside src/ and tests/")
        self.assertClean({**VALID, "benchmark/a/b/M_STATS_PROBDIST_STRAY.bas": source})

    def test_name_prefix_and_length(self) -> None:
        self.assertFinding({**VALID, "tests/M_TEST.bas": renamed(TEST, "M_TEST")},
                           "must start with M_STATS_PROBDIST_")
        name = "M_STATS_PROBDIST_" + "A" * 15
        self.assertFinding({**VALID, f"tests/{name}.bas": renamed(TEST, name)},
                           "longer than 31 characters")
        name = "M_STATS_PROBDIST_" + "A" * 14
        self.assertClean({**VALID, f"tests/{name}.bas": renamed(TEST, name)})

    def test_vb_name_must_equal_file_name(self) -> None:
        self.assertFinding(changed(TEST_PATH, 'VB_Name = "M_STATS_PROBDIST_TEST"',
                                   'VB_Name = "M_STATS_PROBDIST_TESTS"'),
                           'exactly one Attribute VB_Name = "M_STATS_PROBDIST_TEST"')
        self.assertFinding(changed(TEST_PATH, "Option Explicit\n",
                                   'Option Explicit\nAttribute VB_Name = "M_STATS_PROBDIST_TEST"\n'),
                           "exactly one Attribute VB_Name")

    def test_bas_starts_with_vb_name(self) -> None:
        self.assertFinding(changed(TEST_PATH, "Attribute VB_Name", "'Banner\nAttribute VB_Name"),
                           "starts with its Attribute VB_Name line")

    def test_names_are_case_insensitive(self) -> None:
        files = {**VALID, "benchmark/x/M_STATS_PROBDIST_Test.bas":
                 renamed(TEST, "M_STATS_PROBDIST_Test")}
        self.assertFinding(files, "component name also used by")

    def test_option_explicit(self) -> None:
        for path in VALID:
            with self.subTest(path=path):
                self.assertFinding(changed(path, "Option Explicit\n", ""),
                                   "missing Option Explicit")

    def test_option_private_module_required_in_internal_layers(self) -> None:
        for path in (CORE_PATH, SPECIAL_PATH):
            with self.subTest(path=path):
                self.assertFinding(changed(path, "Option Private Module\n", ""),
                                   "missing Option Private Module")
        without = {key: value for key, value in VALID.items() if key != SPECIAL_PATH}
        self.assertFinding(without, f"{SPECIAL_PATH}: listed in OPTION_PRIVATE_REQUIRED")

    def test_option_private_module_forbidden_beside_worksheet_functions(self) -> None:
        self.assertFinding(changed(FAMILY_PATH, "Option Explicit\n",
                                   "Option Explicit\nOption Private Module\n"),
                           "declares Public K_STATS_ worksheet functions")
        moved = changed(CORE_PATH, "Public Function PROB_Add",
                        "Public Function K_STATS_ADD(ByVal A As Double) As Double\n"
                        "End Function\n\nPublic Function PROB_Add")
        self.assertFinding(moved, "would hide them")
        # Neither required nor forbidden elsewhere.
        for path in (TEST_PATH, EXPORT_PATH):
            with self.subTest(path=path):
                self.assertClean(changed(path, "Option Explicit\n",
                                         "Option Explicit\nOption Private Module\n"))
        # A Private K_STATS_ helper or a comment does not expose a worksheet function.
        self.assertClean(changed(CORE_PATH, "Private Sub PROB_Status",
                                 "'Public Function K_STATS_NOTE()\n"
                                 "Private Function K_STATS_Hidden() As Long\nEnd Function\n\n"
                                 "Private Sub PROB_Status"))

    def test_option_base(self) -> None:
        for base in ("0", "1"):
            with self.subTest(base=base):
                self.assertFinding(changed(TEST_PATH, "Option Explicit\n",
                                           f"Option Explicit\nOption Base {base}\n"),
                                   "Option Base; declare explicit array bounds")

    def test_options_in_comments_do_not_count(self) -> None:
        continued = "Option Explicit\n' An example _\nOption Base 1\n"
        self.assertClean(changed(TEST_PATH, "Option Explicit\n", continued))
        for replacement in ("' Option Explicit\n", "' Note _\nOption Explicit\n",
                            "Rem Option Explicit\n"):
            with self.subTest(replacement=replacement):
                self.assertFinding(changed(TEST_PATH, "Option Explicit\n", replacement),
                                   "missing Option Explicit")
        self.assertFinding(changed(CORE_PATH, "Option Private Module\n",
                                   "' Note _\nOption Private Module\n"),
                           "missing Option Private Module")

    def test_ascii_only(self) -> None:
        self.assertFinding(changed(TEST_PATH, "Helper = 1", 'Helper = Len("\udce9")'),
                           "non-ASCII byte 0xE9")
        self.assertFinding(changed(TEST_PATH, "'Sheet 1'", "'Sheet \udcc3\udca9'"),
                           "non-ASCII byte 0xC3")
        bom = {**VALID, TEST_PATH: b"\xef\xbb\xbf" + TEST.replace("\n", "\r\n").encode("ascii")}
        self.assertFinding(bom, "byte-order mark")

    def test_tabs_and_trailing_white_space(self) -> None:
        self.assertFinding(changed(TEST_PATH, "        Helper =", "\tHelper ="), "tab character")
        self.assertFinding(changed(TEST_PATH, "End Function", "End Function  "),
                           "trailing white space")
        self.assertFinding(changed(TEST_PATH, "Option Explicit", "Option Explicit\t"),
                           "trailing white space")

    def test_unclosed_string_literal(self) -> None:
        self.assertFinding(changed(TEST_PATH, "Helper = 1", 'Helper = Len("open)'),
                           "string literal not closed")
        self.assertClean(changed(TEST_PATH, "Helper = 1", 'Helper = Len("""quoted""") \' "'))

    def test_comment_continuation_is_comment(self) -> None:
        self.assertClean(changed(TEST_PATH, "End Sub",
                                 "        ' A comment that continues _\n"
                                 "        on \" this line\nEnd Sub"))

    def test_storage_needs_lf_and_crlf_checkout(self) -> None:
        findings = self.findings(VALID, attributes=False)
        self.assertTrue(any("stored with crlf" in item for item in findings), findings)
        self.assertTrue(any("check it out with CRLF" in item for item in findings), findings)

    def test_main_exit_codes(self) -> None:
        invalid = changed(TEST_PATH, "Option Explicit\n", "")
        for files, expected in ((VALID, 0), (invalid, 1)):
            with self.subTest(expected=expected), repository(files) as root, \
                    contextlib.redirect_stdout(io.StringIO()) as output:
                self.assertEqual(check_vba_source.main(["--root", str(root)]), expected)
                self.assertIn("static check, nothing compiled or run", output.getvalue())


CONDITIONAL_VALID = {
    "nested": '''#If VBA7 Then
    #If Win64 Then
        Private Declare PtrSafe Function F64 Lib "kernel32" () As LongPtr
    #ElseIf Win32 Then
        Private Declare PtrSafe Function F32 Lib "kernel32" () As Long
    #End If
#Else
    Private Declare Function FLegacy Lib "kernel32" () As Long
#End If
''',
    "inactive-legacy": '''#If Not VBA7 Then
    #If Win64 Then
        Private Declare Function Impossible Lib "kernel32" () As Long
    #End If
#End If
''',
    "boolean-expression": '''#If Not VBA7 Then
    Private Declare Function Legacy Lib "kernel32" () As Long
#ElseIf VBA7 And (Win64 Or Win32) Then
    Private Declare PtrSafe Function Modern Lib "kernel32" () As LongPtr
#End If
''',
    "continued-declare": '''#If VBA7 Then
    Private Declare PtrSafe Function Modern Lib "kernel32" _
        () As LongPtr
#Else
    Private Declare Function Legacy Lib "kernel32" _
        () As Long
#End If
''',
    "colon-separated-ptrsafe": 'Private Const N As Long = 1: Private Declare PtrSafe Function B '
                               'Lib "kernel32" () As LongPtr\n',
    "colon-separated-legacy-under-vba6": '''#If Not VBA7 Then
    Private Const N As Long = 1: Private Declare Function B Lib "kernel32" () As Long
#End If
''',
    "literals": '''#If True And Not 0 Then
    Private Declare PtrSafe Function Modern Lib "kernel32" () As LongPtr
#End If
''',
    "mac-and-win16-are-false": '''#If Mac Or Win16 Then
    Private Declare Function Elsewhere Lib "libc.dylib" () As Long
#End If
''',
    "comments-and-strings-are-not-declares": '''' Private Declare Function Bad Lib "kernel32" () As Long
Rem Declare Function Bad Lib "kernel32" () As Long
' A comment that continues _
  Private Declare Function Bad Lib "kernel32" () As Long
Private Const TEXT As String = "Declare Function Bad Lib ""k"" () As Long"
''',
    "module-const-false": '''#Const PROB_WRITE_STATUS_BAR = False
Private Sub Status()
#If PROB_WRITE_STATUS_BAR Then
        Application.StatusBar = "x"
#End If
End Sub
#If PROB_WRITE_STATUS_BAR Then
Private Declare Function Bad Lib "kernel32" () As Long
#End If
''',
    "module-const-expression": '''#Const LEGACY = Not VBA7
#Const MODERN = Not LEGACY And (Win64 Or Win32)
#If LEGACY Then
    Private Declare Function Legacy Lib "kernel32" () As Long
#ElseIf MODERN Then
    Private Declare PtrSafe Function Modern Lib "kernel32" () As LongPtr
#End If
''',
    "module-const-per-environment": '''#If VBA7 Then
    #Const HOST = True
#Else
    #Const HOST = False
#End If
#If HOST Then
    Private Declare PtrSafe Function Modern Lib "kernel32" () As LongPtr
#End If
''',
}
CONDITIONAL_INVALID = {
    "declare-reachable-only-under-vba7": ('''#If VBA7 Then
    Private Declare Function Bad Lib "kernel32" () As Long
#Else
    Private Declare Function Legacy Lib "kernel32" () As Long
#End If
''', ":4: Declare reachable under VBA7 without PtrSafe in vba7-win32, vba7-win64"),
    "declare-without-ptrsafe-on-win64": ('''#If VBA7 And Win64 Then
    Private Declare Function Bad Lib "kernel32" () As Long
#End If
''', "without PtrSafe in vba7-win64"),
    "unconditional-declare": ('Private Declare Function Bad Lib "kernel32" () As Long\n',
                              "without PtrSafe in vba7-win32, vba7-win64"),
    "declare-sub": ('Private Declare Sub Bad Lib "kernel32" ()\n',
                    "without PtrSafe in vba7-win32, vba7-win64"),
    "win32-is-true-on-win64": ('''#If Win32 Then
    Private Declare Function WrongOrder Lib "kernel32" () As Long
#ElseIf Win64 Then
    Private Declare PtrSafe Function NeverReached Lib "kernel32" () As LongPtr
#End If
''', "vba7-win64"),
    "colon-separated-declare": ('Private Const N As Long = 1: Private Declare Function B Lib '
                                '"kernel32" () As Long\n',
                                "without PtrSafe in vba7-win32, vba7-win64"),
    "declare-after-string-with-colon": ('Private Const T As String = "a: b": Declare Function '
                                        'B Lib "kernel32" () As Long\n',
                                        "without PtrSafe in vba7-win32, vba7-win64"),
    "vba6-is-true-under-vba7": ('''#If VBA6 Then
    Private Declare Function Legacy Lib "kernel32" () As Long
#End If
''', "without PtrSafe in vba7-win32, vba7-win64"),
    "module-const-true-reaches-declare": ('''#Const ENABLED = True
#If ENABLED Then
    Private Declare Function Bad Lib "kernel32" () As Long
#End If
''', "without PtrSafe in vba7-win32, vba7-win64"),
    "unknown-symbol": ("#If DEBUG_BUILD Then\n#End If\n",
                       "unsupported conditional-compilation symbol 'DEBUG_BUILD'"),
    "const-used-before-definition": ("#If LATER Then\n#End If\n#Const LATER = True\n",
                                     "unsupported conditional-compilation symbol 'LATER'"),
    "const-defined-in-one-environment-only": ("#If VBA7 Then\n#Const HOST = True\n#End If\n"
                                              "#If HOST Then\n#End If\n",
                                              "unsupported conditional-compilation symbol"),
    "const-redefines-compiler-constant": ("#Const Win64 = False\n",
                                          "redefines a compiler constant"),
    "duplicate-const": ("#Const FLAG = True\n#Const FLAG = False\n", "defined twice"),
    "malformed-const": ("#Const FLAG\n", "malformed #Const"),
    "const-not-boolean": ("#Const LEVEL = 2\n", "indeterminate #Const LEVEL"),
    "const-unknown-symbol": ("#Const FLAG = PROJECT_ARG\n", "indeterminate #Const FLAG"),
    "integer-literal": ("#If 2 Then\n#End If\n", "-1 and 0"),
    "missing-then": ("#If VBA7\n#End If\n", "must end with Then"),
    "unclosed": ("#If VBA7 Then\n", "1 conditional-compilation block(s) not closed"),
    "stray-elseif": ("#ElseIf VBA7 Then\n", "#ElseIf without #If"),
    "elseif-after-else": ("#If VBA7 Then\n#Else\n#ElseIf Win64 Then\n#End If\n",
                          "#ElseIf after #Else"),
    "duplicate-else": ("#If VBA7 Then\n#Else\n#Else\n#End If\n", "duplicate #Else"),
    "stray-end-if": ("#End If\n", "#End If without #If"),
    "unsupported-directive": ("#Region Main\n", "unsupported conditional-compilation directive"),
}


@unittest.skipIf(shutil.which("git") is None, "git is not installed")
class ConditionalTest(Assertions):
    check = staticmethod(check_vba_conditionals.check)

    def test_valid_fixtures_pass(self) -> None:
        self.assertClean(VALID)
        for name, body in CONDITIONAL_VALID.items():
            with self.subTest(name=name):
                self.assertClean(standalone(body))

    def test_invalid_fixtures_fail(self) -> None:
        for name, (body, text) in CONDITIONAL_INVALID.items():
            with self.subTest(name=name):
                self.assertFinding(standalone(body), text)

    def test_only_the_vba7_branch_is_reported(self) -> None:
        # The legacy Declare on line 6 is reachable under VBA6 only, where
        # PtrSafe does not exist.
        body, text = CONDITIONAL_INVALID["declare-reachable-only-under-vba7"]
        self.assertEqual(self.findings(standalone(body)),
                         [f"tests/M_STATS_PROBDIST_FIXTURE.bas{text}"])

    def test_reachable_sources_blank_inactive_lines(self) -> None:
        sources, findings = check_vba_conditionals.reachable_sources(
            "x.bas", "A\n#If VBA7 Then\nB\n#Else\nC\n#End If")
        self.assertEqual(findings, [])
        self.assertEqual(sources["vba7-win64"].split("\n"), ["A", "", "B", "", "", ""])
        self.assertEqual(sources["vba6-win32"].split("\n"), ["A", "", "", "", "C", ""])
        sources, findings = check_vba_conditionals.reachable_sources(
            "x.bas", "#Const ON = Win64\n#If ON Then\nB\n#End If")
        self.assertEqual(findings, [])
        self.assertEqual(sources["vba7-win64"].split("\n"), ["", "", "B", ""])
        self.assertEqual(sources["vba7-win32"].split("\n"), ["", "", "", ""])

    def test_main_exit_codes(self) -> None:
        for files, expected in ((VALID, 0), (standalone("#If Unknown Then\n#End If\n"), 1)):
            with self.subTest(expected=expected), repository(files) as root, \
                    contextlib.redirect_stdout(io.StringIO()) as output:
                self.assertEqual(check_vba_conditionals.main(["--root", str(root)]), expected)
                self.assertIn("static check, nothing compiled or run", output.getvalue())


CLASS = '''VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "M_STATS_PROBDIST_CFIXTURE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Label As String

Public Function PROB_Add() As Long
End Function

Friend Sub Reset()
End Sub
'''
MODERN = 'Public Const PROB_HOST As String = "vba7"'
LEGACY = 'Public Const PROB_HOST As String = "vba6"'


@unittest.skipIf(shutil.which("git") is None, "git is not installed")
class VisibilityTest(Assertions):
    check = staticmethod(check_vba_visibility.check)

    def test_valid_project_passes(self) -> None:
        self.assertClean(VALID)

    def test_counts_public_declarations(self) -> None:
        with repository(VALID) as root:
            _, components, declarations = check_vba_visibility.check(root)
        # PROB_LIMIT, PROB_Add, PROB_RatioClass, PROB_Twice, K_STATS_FIXTURE,
        # Test_RunAll and Export_Study.
        self.assertEqual((components, declarations), (5, 7))

    def test_implicit_visibility_is_rejected(self) -> None:
        cases = {
            "Function K_STATS_FIXTURE(": "implicit visibility: Function K_STATS_FIXTURE",
            "Sub Test_RunAll(": "implicit visibility: Sub Test_RunAll",
            "Static Function Helper(": "implicit visibility: Function Helper",
        }
        paths = {"Function K_STATS_FIXTURE(": FAMILY_PATH, "Sub Test_RunAll(": TEST_PATH,
                 "Static Function Helper(": TEST_PATH}
        originals = {"Function K_STATS_FIXTURE(": "Public Function K_STATS_FIXTURE(",
                     "Sub Test_RunAll(": "Public Sub Test_RunAll(",
                     "Static Function Helper(": "Private Function Helper("}
        for replacement, text in cases.items():
            with self.subTest(replacement=replacement):
                self.assertFinding(changed(paths[replacement], originals[replacement],
                                           replacement), text)
        statements = {
            "Dim mOther As Long": "module-level Dim statement",
            "Const OTHER As Long = 1": "module-level Const statement",
            "Enum Mode\n    ModeA = 1\nEnd Enum": "Enum Mode",
            "Type Pair\n    Hi As Double\nEnd Type": "Type Pair",
            'Declare PtrSafe Function Tick Lib "kernel32" () As Long':
                "module-level Declare statement",
            "Property Get Value() As Long\nEnd Property": "Property Get Value",
        }
        for statement, text in statements.items():
            with self.subTest(statement=statement):
                self.assertFinding(standalone(statement + "\n"), f"implicit visibility: {text}")

    def test_explicit_visibility_passes(self) -> None:
        self.assertClean(standalone(
            "Private mOther As Long\nPrivate Const OTHER As Long = 1\n"
            "Private Enum Mode\n    ModeA = 1\nEnd Enum\n"
            "Public Type PROB_FixturePair\n    Hi As Double\nEnd Type\n"
            'Private Declare PtrSafe Function Tick Lib "kernel32" () As Long\n'
            "Public Property Get Fixture_Value() As Long\nEnd Property\n"
            "Public Property Let Fixture_Value(ByVal Value As Long)\nEnd Property\n"
            "Private Static Sub Counter()\n    Dim Count As Long\n    Const STEP As Long = 1\n"
            "End Sub\n"))

    def test_friend_only_in_class_modules(self) -> None:
        self.assertFinding(changed(TEST_PATH, "Private Function Helper", "Friend Function Helper"),
                           "Friend in a standard module: Function Helper")
        self.assertClean({**VALID, "tests/M_STATS_PROBDIST_CFIXTURE.cls": CLASS})

    def test_public_name_in_two_modules(self) -> None:
        twin = changed(TEST_PATH, "Private Function Helper", "Public Function PROB_Add")
        self.assertFinding(twin, f"{TEST_PATH}:9: public name 'PROB_Add' declared also in "
                                 f"{CORE_PATH} (one VBA project): M_STATS_PROBDIST_CORE.PROB_Add "
                                 "(Function) at line 12")
        self.assertFinding(twin, "ambiguous name in vba6-win32, vba7-win32, vba7-win64")

    def test_every_public_kind_collides(self) -> None:
        cases = {
            "Public Const PROB_Add As Long = 1": "M_STATS_PROBDIST_CORE.PROB_Add (Function)",
            "Public Enum PROB_LIMIT\n    PROB_LIMIT_A = 1\nEnd Enum":
                "M_STATS_PROBDIST_CORE.PROB_LIMIT (Const)",
            "Public Type PROB_RatioClass\n    Hi As Double\nEnd Type":
                "M_STATS_PROBDIST_SPECIALFUNCS.PROB_RatioClass (Enum)",
            "Public Sub PROB_Twice()\nEnd Sub":
                "M_STATS_PROBDIST_SPECIALFUNCS.PROB_Twice (Function)",
            "Public K_STATS_FIXTURE As Long": "K_STATS_FIXTURE (Function)",
            "Global PROB_Add As Long": "PROB_Add (Function)",
            "Public Const PROB_ONE As Long = 1, PROB_Add As Long = 2": "PROB_Add (Function)",
            "Public gFirst As Long, gSecond() As Double, PROB_Add As Variant":
                "PROB_Add (Function)",
            'Public Declare PtrSafe Function PROB_Add Lib "kernel32" () As Long':
                "PROB_Add (Function)",
        }
        for statement, text in cases.items():
            with self.subTest(statement=statement):
                self.assertFinding(standalone(statement + "\n"), text)

    def test_case_insensitive_names(self) -> None:
        self.assertFinding(standalone("Public Function prob_add() As Long\nEnd Function\n"),
                           "public name 'prob_add' declared also in")

    def test_benchmark_modules_share_the_collision_group(self) -> None:
        self.assertFinding(changed(EXPORT_PATH, "Public Sub Export_Study", "Public Sub Test_RunAll"),
                           f"{TEST_PATH}:4: public name 'Test_RunAll' declared also in "
                           f"{EXPORT_PATH} (one VBA project)")
        self.assertEqual(check_vba_visibility.COLLISION_GROUPS,
                         {"one VBA project": ("src/", "tests/", "benchmark/")})
        homes = check_vba_source.FLAT_HOMES + check_vba_source.NESTED_HOMES
        self.assertEqual(sorted(homes), sorted(*check_vba_visibility.COLLISION_GROUPS.values()))

    def test_same_module_twice(self) -> None:
        self.assertFinding(changed(TEST_PATH, "Private Function Helper", "Public Sub Test_RunAll"),
                           "public name 'Test_RunAll' declared twice in its module")

    def test_what_does_not_collide(self) -> None:
        # Private names repeat freely, as VBA allows; Helper is Private in two modules.
        self.assertClean(changed(TEST_PATH, "Private Function Helper", "Private Function PROB_Add"))
        # Property Get, Let and Set of one module share one name.
        self.assertClean(standalone("Public Property Get Fixture_Value() As Long\nEnd Property\n"
                                    "Public Property Let Fixture_Value(ByVal V As Long)\n"
                                    "End Property\n"
                                    "Public Property Set Fixture_Value(ByVal V As Object)\n"
                                    "End Property\n"))
        # Class members are not in the project's global namespace.
        self.assertClean({**VALID, "tests/M_STATS_PROBDIST_CFIXTURE.cls": CLASS})
        # Comments and strings declare nothing.
        self.assertClean(standalone("' Public Function PROB_Add() As Long\n"
                                    "Rem Public Sub PROB_Twice()\n"
                                    'Private Const NOTE As String = "Public Sub PROB_Add()"\n'))
        # Enum members are not compared.
        self.assertClean(standalone("Public Enum Fixture_Mode\n    PROB_Add = 1\nEnd Enum\n"))

    def test_exclusive_branches(self) -> None:
        both = f"#If VBA7 Then\n{MODERN}\n#Else\n{LEGACY}\n#End If\n"
        self.assertClean(standalone(both))
        self.assertFinding(standalone(both + MODERN + "\n"),
                           "declared twice in its module")
        modern_only = f"#If VBA7 Then\n{MODERN}\n#End If\n"
        legacy_only = f"#If Not VBA7 Then\n{LEGACY}\n#End If\n"
        other = module(legacy_only, "M_STATS_PROBDIST_OTHER")
        self.assertClean({**standalone(modern_only), "tests/M_STATS_PROBDIST_OTHER.bas": other})
        overlap = module(f"#If Win64 Then\n{LEGACY}\n#End If\n", "M_STATS_PROBDIST_OTHER")
        self.assertFinding({**standalone(modern_only), "tests/M_STATS_PROBDIST_OTHER.bas": overlap},
                           "ambiguous name in vba7-win64")
        dead = f"#If False Then\n{MODERN.replace('PROB_HOST', 'PROB_Add')}\n#End If\n"
        self.assertClean(standalone(dead))

    def test_directive_findings_are_reported(self) -> None:
        self.assertFinding(standalone("#If Unknown Then\n#End If\n"), "indeterminate #If")

    def test_main_exit_codes(self) -> None:
        invalid = changed(TEST_PATH, "Public Sub Test_RunAll", "Sub Test_RunAll")
        for files, expected in ((VALID, 0), (invalid, 1)):
            with self.subTest(expected=expected), repository(files) as root, \
                    contextlib.redirect_stdout(io.StringIO()) as output:
                self.assertEqual(check_vba_visibility.main(["--root", str(root)]), expected)
                self.assertIn("static check, nothing compiled or run", output.getvalue())


class LexerTest(unittest.TestCase):
    def test_comments_end_the_code(self) -> None:
        cases = {
            "x = 1 ' note": "x = 1",
            "Rem whole line": "",
            "  Rem": "",
            "Debug.Print 1: Rem GoTo Missing": "Debug.Print 1:",
            'Debug.Print "a: Rem b"': 'Debug.Print "a: Rem b"',
            "Call Run(Remainder:=1)": "Call Run(Remainder:=1)",
            "Report Name:= Rem": "Report Name:= Rem",
            "Remove 1": "Remove 1",
            "100 Rem GoTo Missing": "100",
            "  10 x = 1 ' note": "  10 x = 1",
            "10: Rem note": "10:",
            "x = ['Sheet 1'!A1] ' note": "x = ['Sheet 1'!A1]",
            "x = 1: Remark = 2": "x = 1: Remark = 2",
            '"it\'s": GoTo Label': '"it\'s": GoTo Label',
            "'Declare Function F Lib \"k\" ()": "",
        }
        for raw, expected in cases.items():
            with self.subTest(raw=raw):
                self.assertEqual(strip_code(raw), expected)
        self.assertEqual(strip_code('Debug.Print "GoTo x": GoTo y', blank_strings=True),
                         "Debug.Print         : GoTo y")

    def test_statements_split_at_colons_outside_literals(self) -> None:
        self.assertEqual(split_statements('a = "x: y": b = #12:30:00 PM#: c = [A1:B2]: Rem d: e'),
                         ['a = "x: y"', "b = #12:30:00 PM#", "c = [A1:B2]"])
        self.assertEqual(split_statements("Call Run(Name:=1): x = 2"),
                         ["Call Run(Name:=1)", "x = 2"])

    def test_continued_comment_lines_are_comment(self) -> None:
        lines = ["x = 1 ' note _", "  GoTo Missing _", "  MsgBox 1", "y = 2", "' plain _"]
        self.assertEqual(logical_units(lines), [(1, 3, "code", "x = 1"), (4, 4, "code", "y = 2"),
                                                (5, 5, "code", "")])
        self.assertEqual(logical_units(["x = 1 + _", "  2"]), [(1, 2, "code", "x = 1 + 2")])
        self.assertEqual(logical_units(["  #If VBA7 Then ' note"]),
                         [(1, 1, "directive", "#If VBA7 Then")])


if __name__ == "__main__":
    unittest.main()
