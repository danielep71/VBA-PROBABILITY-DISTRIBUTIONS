"""Tests for tools/format_vba.py: layout rules, real modules and the baseline ratchet.

Each rule case states one line or a few lines as the formatter must leave
them, so that a failure names the rule. The literal and apostrophe cases are
the ones a line-based formatter is most likely to get wrong. The real-module
cases run on every tracked module's text and only read it.
"""

from __future__ import annotations

import contextlib
import io
import json
import re
import shutil
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "tools"))

import format_vba  # noqa: E402
import vba_lexer  # noqa: E402

HEAD = ['Attribute VB_Name = "M_STATS_PROBDIST_PROBE"', "Option Explicit",
        "Option Private Module", ""]
HAS_GIT = shutil.which("git") is not None


def module(*body: str) -> str:
    return "".join(f"{line}\r\n" for line in [*HEAD, *body])


def body_of(text: str) -> list[str]:
    return text.split("\r\n")[len(HEAD):-1]


# A token is a string literal, a comment (without trailing white space) or a run
# of code characters; white space and line breaks separate tokens.
TOKEN = re.compile(r'"(?:[^"\r\n]|"")*"|\'[^\r\n]*|[^\s"\']+')


def tokens(text: str) -> list[str]:
    return [token.rstrip() if token.startswith("'") else token
            for token in TOKEN.findall(text)]


def logical_lines(text: str) -> tuple[list[tuple[str, str]], list[str]]:
    """The logical lines and statements tools/vba_lexer.py reads in a module.

    Continuations (in code and in comments) are joined and comments removed,
    so a blank line added to or removed from a logical line, or comment text
    turned into code, changes the result. Blank logical lines are left out.
    """
    lines = text.replace("\r\n", "\n").replace("\r", "\n").split("\n")
    units = [(kind, code) for _, _, kind, code in vba_lexer.logical_units(lines)
             if code.strip()]
    return units, [statement for _, statement in vba_lexer.statements("\n".join(lines))
                   if statement]


def real_modules() -> dict[str, str]:
    if not HAS_GIT:
        return {}
    return {name: format_vba.read(ROOT / name) for name in format_vba.tracked_modules(ROOT)}


class RuleTest(unittest.TestCase):
    def assertFormats(self, before: list[str], after: list[str]) -> None:
        formatted = format_vba.format_text(module(*before))
        self.assertEqual(body_of(formatted), after)
        self.assertEqual(format_vba.format_text(formatted), formatted, "not idempotent")
        self.assertEqual(logical_lines(formatted), logical_lines(module(*before)),
                         "logical lines changed")

    def test_formatted_module_is_unchanged(self) -> None:
        text = module("Public Function K_STATS_Probe() As Double", "'", "'Declare",
                      "    Dim Count As Long", "'Return", "    K_STATS_Probe = 1#",
                      "End Function")
        self.assertEqual(format_vba.format_text(text), text)

    def test_line_endings_and_end_of_file(self) -> None:
        expected = module("Private Sub Run()", "End Sub")
        lf = "\n".join([*HEAD, "Private Sub Run()", "End Sub", "", "", ""])
        cr = "\r".join([*HEAD, "Private Sub Run()", "End Sub", ""])
        mixed = "\r\n".join(HEAD) + "\r\nPrivate Sub Run()\nEnd Sub\r"
        no_final = expected.removesuffix("\r\n")
        for label, text in (("LF", lf), ("CR", cr), ("mixed", mixed), ("no final", no_final),
                            ("trailing CRLF blank lines", expected + "\r\n\r\n")):
            with self.subTest(label):
                self.assertEqual(format_vba.format_text(text), expected)

    def test_trailing_white_space(self) -> None:
        self.assertFormats(["Private Sub Run()  ", "    Count = 1\t ", "    ' note  ",
                            "End Sub"],
                           ["Private Sub Run()", "    Count = 1", "    ' note", "End Sub"])

    def test_indentation_tabs(self) -> None:
        self.assertFormats(["Private Sub Run()", "\tIf Ready Then", "\t\tCount = 1",
                            "  \tCount = 2", "\tEnd If", "End Sub"],
                           ["Private Sub Run()", "    If Ready Then", "        Count = 1",
                            "      Count = 2", "    End If", "End Sub"])

    def test_indented_comment_headings_keep_their_indentation(self) -> None:
        # The repository indents statements under their 'Declare / 'Invert headings.
        lines = ["Public Function K_STATS_Probe() As Double", "'Declare",
                 "    Dim X As Double", "    'Invert", "        X = 1# / 2#",
                 "  ' a two-space comment", "End Function"]
        self.assertFormats(lines, lines)

    def test_procedure_and_label_lines_at_outer_level(self) -> None:
        self.assertFormats(["    Public Function Total() As Double", "Clean_Exit:",
                            "    Exit Function", "  Err_Handler:  ' handler",
                            "    End Function"],
                           ["Public Function Total() As Double", "Clean_Exit:",
                            "    Exit Function", "Err_Handler:  ' handler", "End Function"])
        self.assertFormats(["  Friend Property Get Name() As String", "  End Property"],
                           ["Friend Property Get Name() As String", "End Property"])
        self.assertFormats(["   Private Static Sub Tick()", " End Sub"],
                           ["Private Static Sub Tick()", "End Sub"])
        self.assertFormats(["\tSub Bare()", "\tEnd Sub  ' done"],
                           ["Sub Bare()", "End Sub  ' done"])

    def test_statement_keywords_are_not_labels(self) -> None:
        lines = ["Private Sub Run()", "    If Ready Then", "        Count = 1",
                 "    Else:", "        Count = 2", "    End If", "    Do", "    Loop:",
                 "End Sub"]
        self.assertFormats(lines, lines)

    def test_continuation_lines_keep_their_indentation(self) -> None:
        self.assertFormats(["Public Function Total( _", "    ByVal Amount As Double, _",
                            "    ByVal Rate As Double) _", "    As Double", "End Function"],
                           ["Public Function Total( _", "    ByVal Amount As Double, _",
                            "    ByVal Rate As Double) _", "    As Double", "End Function"])
        # A continued statement whose next line looks like a label is left alone.
        self.assertFormats(["Private Sub Run()", "    Debug.Print _", "        Done:",
                            "End Sub"],
                           ["Private Sub Run()", "    Debug.Print _", "        Done:",
                            "End Sub"])
        # A continuation at the end of a comment continues the comment.
        self.assertFormats(["Private Sub Run()", "    ' note _", "        Done:", "End Sub"],
                           ["Private Sub Run()", "    ' note _", "        Done:", "End Sub"])
        # A tab before the underscore is a continuation too.
        self.assertFormats(["Private Sub Run()", "    Debug.Print\t_", "        Done:",
                            "End Sub"],
                           ["Private Sub Run()", "    Debug.Print\t_", "        Done:",
                            "End Sub"])

    def test_end_with_a_continued_comment(self) -> None:
        # The blank lines after End wait until its continued comment has ended:
        # inserted earlier, they would end the comment and make its next
        # physical line executable.
        before = ["Private Sub One()", "End Sub ' explanation _", "    Debug.Print 1",
                  "Private Sub Two()", "End Sub"]
        self.assertFormats(before, ["Private Sub One()", "End Sub ' explanation _",
                                    "    Debug.Print 1", "", "", "Private Sub Two()",
                                    "End Sub"])
        self.assertSameLogicalLines(module(*before))
        self.assertFormats(["Private Sub One()", "End Sub ' a _", "  b _", "  c",
                            "Private Sub Two()", "End Sub"],
                           ["Private Sub One()", "End Sub ' a _", "  b _", "  c", "", "",
                            "Private Sub Two()", "End Sub"])

    def test_blank_lines_inside_a_logical_line_are_kept(self) -> None:
        # A blank line after a continuation ends that logical line: it is
        # neither removed nor joined to what follows, even after an End or at
        # the end of the file.
        cases = [
            ["Private Sub One()", "    X = 1 _", "", "    Y = 2", "End Sub"],
            ["Private Sub One()", "End Sub ' a _", "", "", "", "Private Sub Two()",
             "End Sub"],
            ["Private Sub One()", "End Sub ' a _", ""],
            ["Private Sub One()", "    ' a _", "", "End Sub"],
        ]
        for before in cases:
            with self.subTest(before=before):
                text = module(*before)
                formatted = format_vba.format_text(text)
                self.assertEqual(format_vba.format_text(formatted), formatted)
                self.assertSameLogicalLines(text)
        self.assertFormats(["Private Sub One()", "End Sub ' a _", ""],
                           ["Private Sub One()", "End Sub ' a _", ""])

    def assertSameLogicalLines(self, text: str) -> None:
        self.assertEqual(logical_lines(format_vba.format_text(text)), logical_lines(text))

    def test_two_blank_lines_after_end(self) -> None:
        self.assertFormats(["Private Sub One()", "End Sub", "'Header of Two",
                            "Private Sub Two()", "End Sub"],
                           ["Private Sub One()", "End Sub", "", "", "'Header of Two",
                            "Private Sub Two()", "End Sub"])
        self.assertFormats(["Private Sub One()", "End Sub", "", "Private Sub Two()",
                            "End Sub"],
                           ["Private Sub One()", "End Sub", "", "", "Private Sub Two()",
                            "End Sub"])
        self.assertFormats(["Private Sub One()", "End Sub", "", "", "", "",
                            "'=====", "' SUITE", "Private Sub Two()", "End Sub"],
                           ["Private Sub One()", "End Sub", "", "", "'=====", "' SUITE",
                            "Private Sub Two()", "End Sub"])
        # Blank lines inside a procedure and between declarations are not touched.
        lines = ["Private Const A As Long = 1", "", "", "", "Private Sub One()", "",
                 "    Count = 1", "", "", "", "    Count = 2", "End Sub"]
        self.assertFormats(lines, lines)

    def test_no_blank_lines_at_the_end(self) -> None:
        self.assertFormats(["Private Sub One()", "End Sub", "", ""],
                           ["Private Sub One()", "End Sub"])
        self.assertFormats(["Private Const A As Long = 1", "", "  ", ""],
                           ["Private Const A As Long = 1"])

    def test_other_lines_are_not_moved(self) -> None:
        lines = ["Private Sub Run()", "    If Ready Then Exit Sub", "    Select Case Mode",
                 "        Case Else: Count = 0", "    End Select", "    End",
                 "    Debug.Print Count: Count = 1", "  Private Declare PtrSafe Function "
                 "GetTickCount Lib \"kernel32\" () As Long", "End Sub"]
        self.assertFormats(lines, lines)


class LiteralTest(unittest.TestCase):
    """No rule changes a string literal, or mistakes one for code."""

    CASES = [
        ("a label name in a literal is data", '    Message = "Fail:"'),
        ("a procedure keyword in a literal is data", '    Message = "End Sub"'),
        ("an apostrophe in a literal does not start a comment",
         "    Const Quote As String = \"a 'b' c\""),
        ("a doubled quote does not unbalance the scan",
         "    Const Said As String = \"he said \"\"go\"\" 'x' now\""),
        ("a tab inside a literal is kept", '    Message = "a\tb"'),
        ("spaces inside a literal are kept", '    Message = "a     b"'),
    ]

    def test_literals_are_unchanged(self) -> None:
        for label, line in self.CASES:
            with self.subTest(label):
                text = module("Private Sub Run()", line, "End Sub")
                self.assertIn(line, body_of(format_vba.format_text(text)))

    def test_comment_after_a_literal_with_an_apostrophe(self) -> None:
        line = "Done:  ' after \"a 'quoted' word\""
        self.assertEqual(format_vba.split_code_comment(line),
                         ("Done:  ", "' after \"a 'quoted' word\""))
        self.assertEqual(format_vba.split_code_comment("    X = \"it's\" ' note"),
                         ('    X = "it\'s" ', "' note"))


class HeaderTest(unittest.TestCase):
    def test_vbe_header_is_kept_byte_for_byte(self) -> None:
        header = ["VERSION 5.00", "Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmRun ",
                  "   Caption         =   \"Run\"   ", "\tClientHeight    =   3015",
                  "   Begin VB.CommandButton cmdOk ", "      Caption = \"OK\" ", "   End",
                  "End", "Attribute VB_Name = \"frmRun\"", "Attribute VB_Exposed = False"]
        text = "".join(f"{line}\r\n" for line in [*header, "Option Explicit", "",
                                                    "  Private Sub cmdOk_Click()  ",
                                                    "  End Sub"])
        lines = format_vba.format_text(text).split("\r\n")
        self.assertEqual(lines[:len(header)], header)
        self.assertEqual(lines[len(header):-1],
                         ["Option Explicit", "", "Private Sub cmdOk_Click()", "End Sub"])

    def test_class_header(self) -> None:
        header = ["VERSION 1.0 CLASS", "BEGIN", "  MultiUse = -1  'True", "END",
                  'Attribute VB_Name = "CResult"', "Attribute VB_PredeclaredId = False"]
        lines = [*header, "Option Explicit"]
        self.assertEqual(format_vba.format_body([]), [])
        self.assertEqual(format_vba.header_length(lines), len(header))


def deform(text: str) -> str:
    """Real module text with every kind of drift the formatter must undo."""
    out: list[str] = []
    previous = ""
    for line in text.split("\r\n")[:-1]:
        code = line.strip()
        if format_vba.outer_level(line) and not previous.endswith(" _"):
            line = "  " + line                          # outer-level line indented
        elif line.startswith("    "):
            line = "\t" + line[4:]                      # tab indentation
        if code and not code.startswith(("'", "Attribute")):
            line += " \t"                               # trailing white space
        out.append(line)
        previous = line.rstrip()
        if format_vba.PROCEDURE_END.match(code):
            out.append("")                              # an extra blank line after End
    return "\n".join(out) + "\n\n\n"                    # LF endings, blank lines at the end


@unittest.skipUnless(HAS_GIT, "git is not installed")
class RealModuleTest(unittest.TestCase):
    """The formatter on the text of every tracked module, read only."""

    modules: dict[str, str]

    @classmethod
    def setUpClass(cls) -> None:
        cls.modules = real_modules()

    def test_modules_are_found(self) -> None:
        names = set(self.modules)
        self.assertIn("src/M_STATS_PROBDIST_CORE.bas", names)
        self.assertIn("tests/M_STATS_PROBDIST_TEST.bas", names)
        self.assertTrue(any(name.startswith("benchmark/") for name in names))

    def test_token_stream_is_unchanged(self) -> None:
        for name, text in self.modules.items():
            with self.subTest(name):
                self.assertEqual(tokens(format_vba.format_text(text)), tokens(text))

    def test_logical_lines_are_unchanged(self) -> None:
        # The sequence of logical lines and statements read by tools/vba_lexer.py
        # is the same before and after formatting, as committed and with drift.
        for name, text in self.modules.items():
            with self.subTest(name):
                self.assertTrue(logical_lines(text)[1])
                self.assertEqual(logical_lines(format_vba.format_text(text)),
                                 logical_lines(text))
                drifted = deform(text)
                self.assertEqual(logical_lines(format_vba.format_text(drifted)),
                                 logical_lines(text))

    def test_idempotent(self) -> None:
        for name, text in self.modules.items():
            with self.subTest(name):
                once = format_vba.format_text(text)
                self.assertEqual(format_vba.format_text(once), once)

    def test_today_only_blank_lines_differ(self) -> None:
        # Every committed difference is a blank line added or removed; the text
        # of every non-blank line, indentation included, is already in layout.
        for name, text in self.modules.items():
            with self.subTest(name):
                formatted = format_vba.format_text(text)
                self.assertEqual([line for line in formatted.split("\r\n") if line],
                                 [line for line in text.split("\r\n") if line])

    def test_drift_is_undone(self) -> None:
        for name, text in self.modules.items():
            with self.subTest(name):
                drifted = deform(text)
                self.assertNotEqual(drifted, text)
                self.assertEqual(tokens(drifted), tokens(text))
                formatted = format_vba.format_text(drifted)
                self.assertEqual(tokens(formatted), tokens(text))
                self.assertEqual(formatted, format_vba.format_text(text))

    def test_token_check_detects_a_change_inside_a_literal(self) -> None:
        # The token comparison itself rejects white space changed inside a literal.
        self.assertNotEqual(tokens('X = "a  b"'), tokens('X = "a b"'))
        self.assertNotEqual(tokens("X = a b"), tokens("X = ab"))
        self.assertEqual(tokens("  X = 1 ' note  \r\n"), tokens("X  =  1\t' note"))


@unittest.skipUnless(HAS_GIT, "git is not installed")
class MainTest(unittest.TestCase):
    """The command line, in temporary Git repositories."""

    def setUp(self) -> None:
        directory = tempfile.TemporaryDirectory(prefix="format-vba-")
        self.addCleanup(directory.cleanup)
        self.root = Path(directory.name)
        subprocess.run(["git", "init", "-q"], cwd=self.root, check=True)

    def track(self, files: dict[str, str]) -> None:
        for name, text in files.items():
            (self.root / name).parent.mkdir(parents=True, exist_ok=True)
            (self.root / name).write_bytes(text.encode("latin-1"))
        subprocess.run(["git", "add", "-A"], cwd=self.root, check=True)

    def baseline(self, files: list[str], **extra: object) -> None:
        path = self.root / format_vba.BASELINE
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(json.dumps({"schema_version": 1, "files": files, **extra}),
                        encoding="utf-8")

    def main(self, *args: str) -> tuple[int, str]:
        with contextlib.redirect_stdout(io.StringIO()) as printed:
            code = format_vba.main([*args, "--root", str(self.root)])
        return code, printed.getvalue()

    GOOD = module("Private Sub Run()", "End Sub")
    BAD = module("  Private Sub Run()", "End Sub", "")

    def test_no_module_passes(self) -> None:
        code, printed = self.main("--check")
        self.assertEqual(code, 0)
        self.assertIn("vba-layout: PASS (no VBA module found)", printed)

    def test_repository_passes_with_its_baseline(self) -> None:
        with contextlib.redirect_stdout(io.StringIO()) as printed:
            self.assertEqual(format_vba.main(["--check", "--root", str(ROOT)]), 0)
        self.assertIn("vba-layout: PASS", printed.getvalue())

    def test_only_tracked_vba_modules_are_checked(self) -> None:
        self.track({"src/A.bas": self.GOOD, "src/B.cls": self.GOOD, "notes/C.txt": self.BAD})
        (self.root / "src/Untracked.bas").write_text(self.BAD, encoding="latin-1")
        code, printed = self.main("--check")
        self.assertEqual(code, 0, printed)
        self.assertIn("(2 module(s), 0 baselined", printed)

    def test_module_out_of_layout_fails(self) -> None:
        self.track({"src/A.bas": self.GOOD, "benchmark/study/B.bas": self.BAD})
        code, printed = self.main("--check")
        self.assertEqual(code, 1)
        self.assertIn("FAIL: benchmark/study/B.bas: not in layout", printed)
        self.assertNotIn("src/A.bas", printed)

    def test_baselined_module_out_of_layout_passes(self) -> None:
        self.track({"src/A.bas": self.GOOD, "src/B.bas": self.BAD})
        self.baseline(["src/B.bas"])
        code, printed = self.main("--check")
        self.assertEqual(code, 0, printed)
        self.assertIn("baselined: src/B.bas", printed)
        self.assertIn("(2 module(s), 1 baselined", printed)

    def test_baselined_module_in_layout_fails(self) -> None:
        # The ratchet: a module that regains layout must leave the baseline.
        self.track({"src/A.bas": self.GOOD, "src/B.bas": self.BAD})
        self.baseline(["src/A.bas", "src/B.bas"])
        code, printed = self.main("--check")
        self.assertEqual(code, 1)
        self.assertIn("FAIL: src/A.bas: in layout but still listed in "
                      "tools/vba_layout_baseline.json; remove its entry", printed)

    def test_baseline_entry_that_is_not_a_tracked_module_fails(self) -> None:
        self.track({"src/A.bas": self.GOOD})
        self.baseline(["src/Gone.bas"])
        code, printed = self.main("--check")
        self.assertEqual(code, 1)
        self.assertIn("FAIL: src/Gone.bas: listed in tools/vba_layout_baseline.json but not "
                      "a tracked VBA module", printed)

    def test_write_then_the_baseline_must_shrink(self) -> None:
        self.track({"src/A.bas": self.GOOD, "src/B.bas": self.BAD})
        self.baseline(["src/B.bas"])
        code, printed = self.main("--write")
        self.assertEqual(code, 0)
        self.assertIn("formatted: src/B.bas", printed)
        self.assertIn("1 of 2 module(s) rewritten", printed)
        self.assertEqual((self.root / "src/B.bas").read_bytes(), self.GOOD.encode("latin-1"))
        self.assertEqual(self.main("--check")[0], 1)
        self.baseline([])
        self.assertEqual(self.main("--check")[0], 0)
        # Once the list is empty, the baseline file is deleted: the default is optional.
        (self.root / format_vba.BASELINE).unlink()
        self.assertEqual(self.main("--check")[0], 0)

    def test_no_baseline_checks_every_module(self) -> None:
        self.track({"src/B.bas": self.BAD})
        self.baseline(["src/B.bas"])
        code, printed = self.main("--check", "--no-baseline")
        self.assertEqual(code, 1)
        self.assertIn("FAIL: src/B.bas: not in layout", printed)

    def test_malformed_baselines(self) -> None:
        self.track({"src/A.bas": self.GOOD, "src/B.bas": self.BAD})
        cases = {
            "not JSON": "{",
            "wrong schema": json.dumps({"schema_version": 2, "files": []}),
            "not a list": json.dumps({"schema_version": 1, "files": "src/B.bas"}),
            "unsorted": json.dumps({"schema_version": 1, "files": ["src/B.bas", "src/A.bas"]}),
            "duplicate": json.dumps({"schema_version": 1, "files": ["src/B.bas", "src/B.bas"]}),
            "absolute": json.dumps({"schema_version": 1, "files": ["/src/B.bas"]}),
            "parent": json.dumps({"schema_version": 1, "files": ["../src/B.bas"]}),
        }
        path = self.root / format_vba.BASELINE
        path.parent.mkdir(parents=True, exist_ok=True)
        for label, content in cases.items():
            with self.subTest(label):
                path.write_text(content, encoding="utf-8")
                code, printed = self.main("--check")
                self.assertEqual(code, 2)
                self.assertIn("vba-layout: ERROR:", printed)

    def test_explicit_baseline_must_exist(self) -> None:
        self.track({"src/A.bas": self.GOOD})
        code, printed = self.main("--check", "--baseline", "missing.json")
        self.assertEqual(code, 2)
        self.assertIn("cannot read baseline", printed)

    def test_explicit_paths(self) -> None:
        self.track({"src/A.bas": self.GOOD})
        (self.root / "elsewhere").mkdir()
        (self.root / "elsewhere/B.bas").write_text(self.BAD, encoding="latin-1")
        code, printed = self.main("--check", "elsewhere/B.bas")
        self.assertEqual(code, 1)
        self.assertIn("FAIL: elsewhere/B.bas", printed)
        self.assertIn("(1 module(s)", printed)
        # With explicit paths, baseline entries outside them are not reported.
        self.baseline(["src/Gone.bas"])
        self.assertEqual(self.main("--check", "src/A.bas")[0], 0)

    def test_bytes_outside_ascii_round_trip(self) -> None:
        text = module("Private Sub Run()", "    ' caf\xe9", "End Sub")
        self.track({"src/A.bas": text})
        code, _ = self.main("--write")
        self.assertEqual(code, 0)
        self.assertEqual((self.root / "src/A.bas").read_bytes(), text.encode("latin-1"))


if __name__ == "__main__":
    unittest.main()
