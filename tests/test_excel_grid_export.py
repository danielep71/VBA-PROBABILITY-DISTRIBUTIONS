"""Study-grid exports on the self-hosted runner (#225).

The unattended export (#181) covered five grids, each with its own boolean
workflow input. #225 adds the positive_ratio_subnormal probe CSVs, selected by
name through one text input, export_studies, and checked as "generated" grids:
the probe writes the whole CSV, so the export must keep the committed header,
line count and each row's key fields.

These checks read the workflow, the PowerShell module and the VBA modules as
text, so they need no YAML parser and no Excel:

* the study specs of ci/ExcelGridExport.psm1 and benchmark/apply_excel_export.py
  agree name for name (path, exporter, entry point, field and key counts), and
  the workflow's input description lists exactly those names;
* each entry point is a Public Sub of its exporter module that writes the CSV
  beside the workbook with no dialog, through the same private writer the
  interactive Probe_ macro uses;
* the workflow passes export_studies to the adapter through the environment
  only, never through cmd, and lets it start the export step and the longer
  timeout.

When PowerShell (pwsh) is on PATH, the module's own functions also run on
fixtures, and their counts must equal apply_excel_export.verify_grid's.
GitHub's Ubuntu runners have pwsh; without it those cases are skipped.

Run: python3 -m unittest tests.test_excel_grid_export
"""

from __future__ import annotations

import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
MODULE = ROOT / "ci" / "ExcelGridExport.psm1"
ADAPTER = ROOT / "ci" / "Export-ExcelGrids.ps1"
WORKFLOW = ROOT / ".github" / "workflows" / "excel-vba-regression.yml"

sys.path.insert(0, str(ROOT / "benchmark"))
import apply_excel_export as A  # noqa: E402

GENERATED_CALL = re.compile(
    r'"(?P<name>\w+)" \{\s*return New-GeneratedGridSpec -Name "(?P=name)" '
    r'-FileName "(?P<file>[^"]+)" `\s*-Exporter "(?P<exporter>[^"]+)" '
    r'-Macro "(?P<macro>\w+)" `\s*-Columns (?P<columns>\d+) -KeyColumns (?P<keys>\d+) '
    r'-SuccessPrefix "(?P<prefix>[^"]+)"')


def module_text() -> str:
    return MODULE.read_text(encoding="utf-8")


def study_names_in_module(text: str) -> list[str]:
    match = re.search(r"\$script:StudyGridNames = @\((?P<body>.*?)\)", text, re.S)
    if match is None:
        return []
    return re.findall(r'"(\w+)"', match.group("body"))


def generated_specs_in_module(text: str) -> dict[str, dict[str, object]]:
    folder = "benchmark/positive_ratio_subnormal/"
    return {m.group("name"): {"path": folder + m.group("file"),
                              "exporter": folder + m.group("exporter"),
                              "macro": m.group("macro"), "columns": int(m.group("columns")),
                              "keys": int(m.group("keys")), "prefix": m.group("prefix")}
            for m in GENERATED_CALL.finditer(text)}


def python_study_specs() -> dict[str, A.GridSpec]:
    return {name: spec for name, spec in A.GRIDS.items() if spec.kind == "generated"}


def sub_body(text: str, name: str) -> str | None:
    """The text of `Sub name(` up to its End Sub, or None."""
    match = re.search(rf"^(?:Public |Private )?Sub {re.escape(name)}\((?P<args>[^)]*)\)"
                      r"(?P<body>.*?)^End Sub", text, re.M | re.S)
    return None if match is None else match.group(0)


def step(text: str, name: str) -> str:
    match = re.search(rf"^(?P<indent>\s+)- name: {re.escape(name)}\n(?P<body>.*?)"
                      r"(?=^(?P=indent)- name: |^  \S|\Z)", text, re.M | re.S)
    if match is None:
        raise AssertionError(f"step {name!r} not found")
    return match.group(0)


class StudySpecParityTests(unittest.TestCase):
    def test_module_and_apply_list_the_same_study_grids(self):
        text = module_text()
        names = study_names_in_module(text)
        module_specs = generated_specs_in_module(text)
        python_specs = python_study_specs()
        self.assertTrue(names, "ci/ExcelGridExport.psm1 lists no study grid")
        self.assertEqual(sorted(names), sorted(module_specs),
                         "every listed study grid needs a New-GeneratedGridSpec")
        self.assertEqual(sorted(names), sorted(python_specs),
                         "apply_excel_export.py must know every study grid the runner exports")
        for name in names:
            with self.subTest(name=name):
                ours, theirs = module_specs[name], python_specs[name]
                self.assertEqual(ours["path"], theirs.path)
                self.assertEqual(ours["exporter"], theirs.exporter)
                self.assertEqual(ours["macro"], theirs.macros["full"])
                self.assertEqual(ours["columns"], theirs.min_columns)
                self.assertEqual(ours["keys"], theirs.key_columns)
                self.assertGreater(theirs.key_columns, 0)
                self.assertLess(theirs.key_columns, theirs.min_columns)

    def test_committed_study_grids_have_the_declared_shape(self):
        for name, spec in python_study_specs().items():
            with self.subTest(name=name):
                raw = (ROOT / spec.path).read_bytes()
                lines = A.grid_lines(raw)
                self.assertGreater(len(lines), 1)
                self.assertEqual({len(line.split(b",")) for line in lines}, {spec.min_columns})
                # The committed bytes (LF) verify against themselves as Excel
                # would write them (CRLF), with nothing changed.
                result = A.verify_grid(spec, "full", raw, A.printed(lines))
                self.assertEqual(result["problems"], [])
                self.assertEqual(result["changed_observations"], 0)
                self.assertEqual(result["filled_rows"], len(lines) - 1)
                keys = [tuple(line.split(b",")[:spec.key_columns]) for line in lines[1:]]
                self.assertEqual(len(keys), len(set(keys)), "key fields must identify each row")

    def test_success_messages_match_the_vba(self):
        for name, spec in generated_specs_in_module(module_text()).items():
            with self.subTest(name=name):
                vba = (ROOT / spec["exporter"]).read_text(encoding="ascii")
                prefix = spec["prefix"]
                self.assertIn(f'MsgBox "{prefix.rstrip()}', vba,
                              "the success prefix must be the start of the writer's message")


class EntryPointTests(unittest.TestCase):
    def test_each_entry_point_writes_beside_the_workbook_without_a_dialog(self):
        for name, spec in python_study_specs().items():
            with self.subTest(name=name):
                vba = (ROOT / spec.exporter).read_text(encoding="ascii").replace("\r\n", "\n")
                macro = spec.macros["full"]
                suffix = macro[len("Export_"):]
                entry = sub_body(vba, macro)
                self.assertIsNotNone(entry, f"{spec.exporter} has no Sub {macro}")
                assert entry is not None
                self.assertTrue(entry.startswith(f"Public Sub {macro}()"))
                self.assertNotIn("GetSaveAsFilename", entry)
                self.assertNotIn("ResolveOutputPath", entry)
                self.assertIn("ThisWorkbook.Path", entry)
                self.assertIn(f'"{os.path.basename(spec.path)}"', entry)
                self.assertRegex(entry, rf"\n\s+Write{suffix} ThisWorkbook\.Path")
                probe = sub_body(vba, f"Probe_{suffix}")
                self.assertIsNotNone(probe, "the interactive Probe_ macro must stay")
                assert probe is not None
                self.assertIn("ResolveOutputPath()", probe)
                self.assertRegex(probe, rf"\n\s+Write{suffix} Path\n")
                writer = sub_body(vba, f"Write{suffix}")
                self.assertIsNotNone(writer)
                assert writer is not None
                self.assertTrue(writer.startswith(f"Private Sub Write{suffix}(ByVal Path As String)"))
                self.assertNotIn("ResolveOutputPath", writer)
                self.assertIn("Open Path For Output", writer)


class WorkflowWiringTests(unittest.TestCase):
    def setUp(self):
        self.text = WORKFLOW.read_text(encoding="utf-8")

    def test_export_studies_is_a_text_input_listing_the_study_grids(self):
        match = re.search(r"^      export_studies:\n        description: \"(?P<d>[^\"]+)\"\n"
                          r"        type: string\n        default: \"\"\n", self.text, re.M)
        self.assertIsNotNone(match, "workflow_dispatch needs the export_studies text input")
        assert match is not None
        listed = re.findall(r"\b(prs_\w+)\b", match.group("d"))
        self.assertEqual(sorted(listed), sorted(study_names_in_module(module_text())))

    def test_studies_reach_the_adapter_only_through_the_environment(self):
        export = step(self.text, "Export selected grids through Excel COM")
        self.assertIn("EXPORT_STUDIES: ${{ inputs.export_studies }}", export)
        run = re.search(r"^\s+run: (?P<line>.*)$", export, re.M)
        assert run is not None
        self.assertNotIn("EXPORT_STUDIES", run.group("line"),
                         "cmd must never expand the free-text input")
        self.assertNotIn("inputs.export_studies", run.group("line"))
        adapter = ADAPTER.read_text(encoding="utf-8")
        self.assertRegex(adapter, r"\[string\]\$Studies = \$env:EXPORT_STUDIES")

    def test_studies_start_the_export_step_and_the_long_timeout(self):
        export = step(self.text, "Export selected grids through Excel COM")
        condition = "inputs.export_studies != ''"
        self.assertIn(condition, re.search(r"^\s+if: (.*)$", export, re.M).group(1))
        timeout = re.search(r"^    timeout-minutes: (.*)$", self.text, re.M)
        assert timeout is not None
        self.assertIn(condition, timeout.group(1))


PWSH = shutil.which("pwsh")

PS_SCRIPT = r"""
param([string]$Fixtures)
$ErrorActionPreference = "Stop"
Import-Module (Join-Path $env:REPO_ROOT "ci/ExcelGridExport.psm1") -Force
$out = [ordered]@{}
foreach ($case in (Get-Content -LiteralPath (Join-Path $Fixtures "cases.json") -Raw | ConvertFrom-Json)) {
    $spec = Get-ExportGridSpec -Name $case.grid
    $base = [System.IO.File]::ReadAllBytes((Join-Path $Fixtures ($case.id + ".base")))
    $exp = [System.IO.File]::ReadAllBytes((Join-Path $Fixtures ($case.id + ".exported")))
    $r = Test-ExportedGrid -Spec $spec -BaselineBytes $base -ExportedBytes $exp
    $out[$case.id] = [ordered]@{
        rows = $r.rows; in_scope_rows = $r.in_scope_rows; filled_rows = $r.filled_rows
        error_rows = $r.error_rows; out_of_scope_rows = $r.out_of_scope_rows
        changed_observations = $r.changed_observations; problems = @($r.problems)
    }
}
$selection = [ordered]@{}
foreach ($text in @("prs_gamma_probe, prs_chisquare_probe", "", "main", "prs_gamma_probe prs_gamma_probe")) {
    try { $selection[$text] = [string[]](ConvertTo-StudySelection -Value $text) }
    catch { $selection[$text] = "REFUSED: " + $_.Exception.Message }
}
[ordered]@{ grids = $out; selection = $selection; names = [string[]](Get-StudyGridNames) } | ConvertTo-Json -Depth 6
"""


def regenerate(spec: A.GridSpec, baseline: bytes, *, change_key: bool = False,
               drop_field: bool = False, header: bytes | None = None,
               crlf: bool = True) -> bytes:
    """What a probe writes: the committed rows with the last field changed."""
    lines = A.grid_lines(baseline)
    for i in range(1, len(lines)):
        cols = lines[i].split(b",")
        cols[-1] = b"9.9E-001;0E+000"
        if change_key and i == 1:
            cols[0] = cols[0] + b"x"
        if drop_field and i == 2:
            cols = cols[:-1]
        lines[i] = b",".join(cols)
    if header is not None:
        lines[0] = header
    raw = A.printed(lines)
    return raw if crlf else raw.replace(b"\r\n", b"\n")


@unittest.skipIf(PWSH is None, "pwsh is not on PATH; the PowerShell module cases are skipped")
class PowerShellModuleTests(unittest.TestCase):
    """The module's functions on fixtures, against apply_excel_export.py."""

    @classmethod
    def setUpClass(cls):
        spec = A.GRIDS["prs_gamma_q_complement"]
        base = (ROOT / spec.path).read_bytes()
        cls.cases = {
            "same": (spec, base, A.printed(A.grid_lines(base))),
            "measured": (spec, base, regenerate(spec, base)),
            "key": (spec, base, regenerate(spec, base, change_key=True)),
            "fields": (spec, base, regenerate(spec, base, drop_field=True)),
            "header": (spec, base, regenerate(spec, base, header=b"point_id,x,shape,other")),
            "lf": (spec, base, regenerate(spec, base, crlf=False)),
            "short": (spec, base, A.printed(A.grid_lines(base)[:-1])),
        }
        work = tempfile.mkdtemp(prefix="excel-grid-export-")
        try:
            for case, (_, b, e) in cls.cases.items():
                Path(work, case + ".base").write_bytes(b)
                Path(work, case + ".exported").write_bytes(e)
            Path(work, "cases.json").write_text(json.dumps(
                [{"id": c, "grid": "prs_gamma_q_complement"} for c in cls.cases]))
            script = Path(work, "run.ps1")
            script.write_text(PS_SCRIPT)
            proc = subprocess.run([PWSH, "-NoProfile", "-NonInteractive", "-File", str(script),
                                   "-Fixtures", work], capture_output=True, text=True,
                                  env=dict(os.environ, REPO_ROOT=str(ROOT)), timeout=300)
            if proc.returncode:
                raise AssertionError(proc.stdout + proc.stderr)
            cls.result = json.loads(proc.stdout)
        finally:
            shutil.rmtree(work, ignore_errors=True)

    def test_counts_and_problems_match_python(self):
        for case, (spec, base, exported) in self.cases.items():
            with self.subTest(case=case):
                ours = A.verify_grid(spec, "full", base, exported)
                theirs = self.result["grids"][case]
                for key in ("rows", "in_scope_rows", "filled_rows", "error_rows",
                            "out_of_scope_rows", "changed_observations"):
                    self.assertEqual(theirs[key], ours[key], key)
                self.assertEqual(bool(theirs["problems"]), bool(ours["problems"]),
                                 (theirs["problems"], ours["problems"]))

    def test_expected_verdicts(self):
        grids = self.result["grids"]
        self.assertEqual(grids["same"]["problems"], [])
        self.assertEqual(grids["same"]["changed_observations"], 0)
        self.assertEqual(grids["measured"]["problems"], [])
        self.assertEqual(grids["measured"]["changed_observations"], grids["measured"]["rows"])
        self.assertIn("line 2: the point's key fields changed", grids["key"]["problems"])
        self.assertIn("line 3: 11 fields, not 12", grids["fields"]["problems"])
        self.assertIn("the header line changed", grids["header"]["problems"])
        self.assertIn("the export is not CRLF-terminated lines as VBA Print # writes them",
                      grids["lf"]["problems"])
        self.assertTrue(any(p.startswith("line count changed") for p in grids["short"]["problems"]))

    def test_study_selection(self):
        selection = self.result["selection"]
        self.assertEqual(selection["prs_gamma_probe, prs_chisquare_probe"],
                         ["prs_gamma_probe", "prs_chisquare_probe"])
        self.assertEqual(selection[""], [])
        self.assertTrue(selection["main"].startswith("REFUSED: unknown study grid 'main'"))
        self.assertIn("named twice", selection["prs_gamma_probe prs_gamma_probe"])
        self.assertEqual(sorted(self.result["names"]), sorted(python_study_specs()))


if __name__ == "__main__":
    unittest.main()
