"""
Fixtures for carried certification (#301): vba_code_identity.py and the carry
path of _manifest.verify_source_binding / verify_holdout_binding.

  1. Comments and layout are not code. A changed or added apostrophe comment,
     a Rem line, a Rem after a colon, a comment continued with " _",
     indentation, trailing blanks, blank lines, CRLF line ends, and a logical
     line re-wrapped with continuations all leave the code stream unchanged.
  2. NEGATIVE: code is code. Each of these is reported as a code change: a
     changed numeric literal; a changed string literal; whitespace or an
     apostrophe changed inside a string literal; a statement that only looks
     like a comment (an apostrophe inside a string, "Remainder = 1"); a
     changed statement inside #If; a changed #If or #Const directive; a
     statement commented out; a comment uncommented; a changed [...]
     expression; `a=1` against `a = 1`; a reordered pair of statements.
  3. NEGATIVE: the comparison fails closed on a non-ASCII byte.
  4. carried() on a real repository: a comment-only edit of a committed
     module is carried from the commit; a code edit, a recorded digest that is
     not the commit's, an unknown commit, a missing commit name and a deleted
     module are refused.
  5. The manifest verifiers: with a `carried` list a comment-only edit is
     reported there and is not a mismatch, for the main and the holdout
     binding; a code edit stays a mismatch and names the changed statement;
     with carried=None (strict) the comment-only edit is a mismatch too.

Run: python3 test_vba_code_identity.py   (exit 0 = pass)
"""
import os
import subprocess
import sys
import tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import _manifest as M  # noqa: E402
import vba_code_identity as V  # noqa: E402

fails = []


def check(cond, msg):
    if not cond:
        fails.append(msg)


BASE = (
    'Attribute VB_Name = "M_FIXTURE"\n'
    "Option Explicit\n"
    "' Module comment\n"
    "#Const FAST = 1\n"
    "Public Function F(ByVal X As Double) As Double\n"
    "    Dim Total As Double, Remainder As Double   ' locals\n"
    '    Const LABEL As String = "it\'s  a \'quoted\' label"\n'
    "    Rem an old-style comment\n"
    "#If FAST Then\n"
    "    Total = X * 2#\n"
    "#Else\n"
    "    Total = X + X\n"
    "#End If\n"
    "    Remainder = 1: Rem trailing remark\n"
    "    F = Total + _\n"
    "        Remainder + [A1 : B2]\n"
    "End Function\n"
)


def same(new, old=BASE):
    return V.first_difference(old.encode("ascii"), new.encode("ascii")) is None


# 1. Comments and layout.
SAME = {
    "a changed comment": BASE.replace("' Module comment", "' Another comment entirely"),
    "an added comment line": BASE.replace("Option Explicit\n", "Option Explicit\n' new\n"),
    "a removed trailing comment": BASE.replace("   ' locals", ""),
    "a changed Rem line": BASE.replace("Rem an old-style comment", "Rem rewritten"),
    "a changed Rem after a colon": BASE.replace("Rem trailing remark", "Rem other"),
    "a comment continued with _": BASE.replace(
        "' Module comment", "' Module comment _\n    Total = 99 is still comment"),
    "indentation": BASE.replace("    Total = X * 2#", "\t\tTotal = X * 2#"),
    "trailing blanks": BASE.replace("Option Explicit\n", "Option Explicit   \n"),
    "blank lines": BASE.replace("Option Explicit\n", "Option Explicit\n\n\n"),
    "CRLF line ends": BASE.replace("\n", "\r\n"),
    "a joined continuation": BASE.replace("F = Total + _\n        Remainder", "F = Total + Remainder"),
    "a new continuation": BASE.replace("Dim Total As Double, Remainder",
                                       "Dim Total As Double, _\n        Remainder"),
    "runs of spaces between tokens": BASE.replace("Total = X + X", "Total  =  X   +   X"),
}
for label, text in SAME.items():
    check(text != BASE and same(text), f"1: {label} must not be a code change")

# 2. Code.
DIFFERENT = {
    "a changed numeric literal": BASE.replace("X * 2#", "X * 3#"),
    "a changed string literal": BASE.replace("label", "lAbel"),
    "whitespace inside a string": BASE.replace("it's  a", "it's a"),
    "an apostrophe inside a string": BASE.replace("it's", "its"),
    "the text after an apostrophe in a string": BASE.replace("'quoted' label", "'quoted' LABEL"),
    "a statement starting with Rem...": BASE.replace("Remainder = 1:", "Remainder = 2:"),
    "a statement inside #If": BASE.replace("Total = X + X", "Total = X + X + 0"),
    "an #If condition": BASE.replace("#If FAST Then", "#If Not FAST Then"),
    "a #Const value": BASE.replace("#Const FAST = 1", "#Const FAST = 0"),
    "a statement commented out": BASE.replace("    Total = X * 2#", "    ' Total = X * 2#"),
    "a statement turned into Rem": BASE.replace("    Total = X * 2#", "    Rem Total = X * 2#"),
    "a comment uncommented": BASE.replace("' Module comment", "Dim Leak As Long"),
    "a [...] expression": BASE.replace("[A1 : B2]", "[A1:B2]"),
    "whitespace removed between tokens": BASE.replace("Total = X + X", "Total=X+X"),
    "two statements swapped": BASE.replace(
        "    Total = X * 2#\n#Else\n    Total = X + X\n",
        "    Total = X + X\n#Else\n    Total = X * 2#\n"),
    "a removed statement": BASE.replace("    Remainder = 1: Rem trailing remark\n", ""),
    "an added statement": BASE.replace("End Function\n", "End Function\nPublic Sub G()\nEnd Sub\n"),
    "a continuation that hides code in a comment": BASE.replace(
        "    Dim Total As Double, Remainder As Double   ' locals\n",
        "    Dim Total As Double   ' locals _\n    Dim Remainder As Double\n"),
}
for label, text in DIFFERENT.items():
    check(text != BASE and not same(text), f"2: {label} must be a code change")

# 3. Fail closed.
try:
    V.code_stream(BASE.replace("Module comment", "Module commént").encode("latin-1"))
    check(False, "3: a non-ASCII byte must refuse the comparison")
except V.CarryRefused:
    pass


# 4 and 5. A real repository.
def git(root, *args):
    return subprocess.run(["git", "-C", root, "-c", "user.name=fixture",
                           "-c", "user.email=fixture@example.invalid",
                           "-c", "commit.gpgsign=false"] + list(args),
                          check=True, capture_output=True, text=True).stdout.strip()


def write(root, rel, text):
    path = os.path.join(root, *rel.split("/"))
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", encoding="ascii", newline="") as f:
        f.write(text)


def refused(root, path, recorded, commit):
    try:
        V.carried(root, path, recorded, commit)
    except V.CarryRefused as exc:
        return str(exc)
    return None


with tempfile.TemporaryDirectory(prefix="carried-") as root:
    git(root, "init", "-q")
    SRC, EXPORTER = "src/M_FIXTURE.bas", "benchmark/holdout/M_HOLDOUT.bas"
    write(root, SRC, BASE)
    write(root, EXPORTER, BASE.replace("M_FIXTURE", "M_HOLDOUT"))
    write(root, "benchmark/grid.csv", "function,arg1\nF,1\n")
    write(root, "benchmark/contracts.csv", "contract_id\nF.all\n")
    git(root, "add", "-A")
    git(root, "commit", "-qm", "executed")
    commit = git(root, "rev-parse", "HEAD")
    recorded = V.normalized_hash_bytes(BASE.encode())
    grid, contracts = (os.path.join(root, "benchmark", n) for n in ("grid.csv", "contracts.csv"))
    manifest = {"commit_sha": commit[:7], "source_binding": M.source_hashes(root),
                "grid_schema_version": M.GRID_SCHEMA_VERSION,
                "grid_schema_columns": M.grid_columns(grid),
                "observation_grid_sha256": M.normalized_hash(grid),
                "contract_registry_sha256": M.normalized_hash(contracts)}
    holdout = dict(manifest, manifest_version=M.MANIFEST_VERSION,
                   evidence_kind="independent_holdout", observation_row_count=1,
                   grid_schema_version=M.HOLDOUT_GRID_SCHEMA_VERSION,
                   source_binding=M.source_hashes(root, M._HOLDOUT_BAS_GLOBS))
    check(M.verify_source_binding(root, manifest, grid, contracts, carried=[]) == []
          and M.verify_holdout_binding(root, holdout, grid, contracts, carried=[]) == [],
          "5: the untouched fixture is bound")

    write(root, SRC, SAME["a changed comment"].replace("\n", "\r\n"))
    check(refused(root, SRC, recorded, commit) is None, "4: a comment-only edit is carried")
    check(refused(root, SRC, recorded, commit[:7]) is None, "4: a short commit name works")
    check("not the recorded" in (refused(root, SRC, "sha256:" + "0" * 64, commit) or ""),
          "4: a recorded digest that is not the commit's is refused")
    check(refused(root, SRC, recorded, "f" * 40) is not None, "4: an unknown commit is refused")
    check(refused(root, SRC, recorded, "") is not None, "4: a missing commit name is refused")
    for verify, record, label in ((M.verify_source_binding, manifest, "main"),
                                  (M.verify_holdout_binding, holdout, "holdout")):
        seen = []
        check(verify(root, record, grid, contracts, carried=seen) == [] and seen == [SRC],
              f"5: {label}: a comment-only edit is carried, not a mismatch")
        strict = verify(root, record, grid, contracts)
        check(len(strict) == 1 and SRC in strict[0],
              f"5: {label}: without a carried list every byte difference is a mismatch")

    write(root, SRC, DIFFERENT["a changed numeric literal"])
    check("code changed" in (refused(root, SRC, recorded, commit) or ""), "4: a code edit is refused")
    for verify, record, label in ((M.verify_source_binding, manifest, "main"),
                                  (M.verify_holdout_binding, holdout, "holdout")):
        seen = []
        problems = verify(root, record, grid, contracts, carried=seen)
        check(len(problems) == 1 and seen == [] and "code changed" in problems[0]
              and "X * 3#" in problems[0],
              f"5: {label}: a code edit stays a mismatch and names the statement: {problems}")

    os.unlink(os.path.join(root, "src", "M_FIXTURE.bas"))
    check("cannot read" in (refused(root, SRC, recorded, commit) or ""),
          "4: a deleted module is refused")

if fails:
    print(f"FAIL: {len(fails)} carried-certification check(s)")
    for m in fails:
        print("  - " + m)
    raise SystemExit(1)
print(f"PASS: {len(SAME)} comment and layout edits keep the code stream, {len(DIFFERENT)} code "
      "edits change it, non-ASCII input is refused, and the manifest verifiers carry a "
      "comment-only edit from the executed commit while a code edit stays a mismatch")
