#!/usr/bin/env python3
"""The plan of the Home and README sheets: their texts, written in desktop Excel.

Run by tools/demobuild/explorer_build.py, which writes home-plan.json beside the
package; tools/demobuild/Evolve-Home.ps1 applies it after the other sheets.

Each sheet keeps its two title bands (rows 1 to 3) and gets a new subtitle and
new tables from row 5: a heading, then rows of a label in column C and a text in
column D, in the template's own table styles. Three kinds of table:

  text      a label and its text
  links     a sheet's tab name, linked to its defined name (nav_<CodeName>), and
            its section and tip: the links work with macros disabled
  legend    a label shown in the format of the role it names (an input, a
            result, a heading, a check), and what that role means

A text may hold three fields that only the build knows: {version}, the library
version (VERSION); {commit}, the source commit; and {excel}, the Excel that
applied the plan, with its build and bitness, which Evolve-Home.ps1 reads.
"""
from __future__ import annotations

import json
import sys
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(Path(__file__).resolve().parent))
import build_demo  # noqa: E402

FIRST_ROW = 5
TEXT_PER_LINE = build_demo.TEXT_PER_LINE
HEIGHTS = (build_demo.ROW_LINE, build_demo.ROW_TWO_LINES, build_demo.ROW_THREE_LINES)
FIELDS = ("{version}", "{commit}", "{excel}")
# The longest each field can be once filled: a 40-digit commit, a version, an Excel line.
FIELD_WIDTH = {"{version}": 12, "{commit}": 40, "{excel}": 40}
SWATCHES = ("input", "result", "heading", "pass", "warning", "fail")

TAB_LEGEND = ("Navy: instructions and reference (Home, Diagnostics, README). Gold: sheets to explore. "
              "Green: accuracy comparisons. Blue: applications.")


def version() -> str:
    return (ROOT / "VERSION").read_text(encoding="utf-8").strip()


def height(text: str) -> int:
    """The row height for a text of column D, with its fields at their widest."""
    length = len(text)
    for field in FIELDS:
        length += text.count(field) * (FIELD_WIDTH[field] - len(field))
    if length > 3 * TEXT_PER_LINE:
        raise ValueError(f"a text holds three lines at most: {text[:60]}...")
    return HEIGHTS[0] if length <= TEXT_PER_LINE else HEIGHTS[1] if length <= 2 * TEXT_PER_LINE else HEIGHTS[2]


def table(heading: str, rows: list[dict[str, Any]], kind: str = "text") -> dict[str, Any]:
    if len(rows) < 2:
        raise ValueError("a table has two rows at least: the template has no one-row style")
    for row in rows:
        row["height"] = height(row["text"])
    return {"heading": heading, "kind": kind, "rows": rows}


def row(label: str, text: str, **extra: Any) -> dict[str, Any]:
    return {"label": label, "text": text, **extra}


def home() -> dict[str, Any]:
    menu = build_demo.sections()
    others = [sheet for _, items in menu[1:] for sheet in items]
    project = build_demo.PROJECT
    return {
        "sheet": "Home",
        "subtitle": ("Start here: what this workbook shows, what it needs, how to use it and a link to every "
                     f"sheet. The commands are on the {project} tab of the Ribbon."),
        "tables": [
            table("What this workbook shows", [
                row("Shows", "The K_STATS_ worksheet functions of VBA Probability Distributions at work: densities, "
                             "probabilities, quantiles and moments of the continuous and discrete families, beside "
                             "Excel's own functions and independent references."),
                row("Demonstration", "The sheets are worked examples. They compare the library with Excel and with "
                                     "references computed outside Excel (mpmath, SciPy) to show how to read the "
                                     "results. They do not validate the library."),
                row("Validation", "The accuracy of the library rests on the evidence in its repository: the "
                                  "certified Excel runs, the reference studies and their thresholds "
                                  "(docs/EXCEL_CERTIFICATION.md). Report a difference you find here as an issue "
                                  "there."),
                row("Version", "VBA Probability Distributions {version}, built from commit {commit}. README gives "
                               "the build record."),
            ]),
            table("What it needs", [
                row("Excel", "Microsoft 365, or Excel 2024 or later, for Windows: the sheets use dynamic arrays, "
                             "LET, LAMBDA, MAP and VSTACK. Tested on {excel}. Excel for Mac, older versions and "
                             "32-bit Excel are not tested."),
                row("Macros", "Enable macros: the K_STATS_ functions, the results of every sheet and the Ribbon "
                              "commands need them. With macros disabled the workbook shows the values it was saved "
                              "with, and the links below and the bar on the left still open every sheet."),
                row("Downloaded file", "If Windows blocks the macros of a downloaded copy, close it, open its "
                                       "Properties in File Explorer, tick Unblock and open it again."),
                row("Python", "Not needed: the mpmath and SciPy references are stored in the workbook. With Python "
                              "in Excel (Microsoft 365), the Comparison sheet also grades a value of yours from an "
                              "=PY() formula."),
            ]),
            table("How to use", [
                row("1. Choose a sheet", f"Use Go to on the {project} tab, the bar on the left or the links below. "
                                         "The bar highlights the current sheet."),
                row("2. Change the inputs", "Type in the input cells, under Editable inputs. The results, the "
                                            "charts and the checks follow at once. A value outside its limits is "
                                            "refused with a message that gives them."),
                row("3. Recalculate", f"Recalculate on the {project} tab recalculates this workbook only."),
                row("4. Reset inputs", "Reset inputs on the same tab puts the inputs of the active sheet back to "
                                       "its example, after a confirmation. It is not available here or on README."),
                row("5. Hide the bar", "Navigation bar on the same tab hides or shows the bar on the left for this "
                                       "session. The workbook is always saved with the bar shown."),
            ]),
            table("Colour legend", [
                row("Input", "A cell you can change: a choice from a list or a number within the limits its "
                             "message gives.", swatch="input"),
                row("Result", "A value computed by a formula of the workbook: a K_STATS_ call, an Excel function, "
                              "or a figure derived from them.", swatch="result"),
                row("Heading", "A section or the heading of a table.", swatch="heading"),
                row("As documented", "A check that holds: the value agrees with its reference or with its "
                                     "documented outcome.", swatch="pass"),
                row("Attention", "A value to read with care, near a limit or below a threshold; the sheet says "
                                 "which.", swatch="warning"),
                row("UNEXPECTED", "A check that does not hold: the value departs from its reference or from its "
                                  "documented outcome.", swatch="fail"),
                row("Tabs", TAB_LEGEND),
            ], kind="legend"),
            table("Sheets", [row(sheet.tab, f"{sheet.section}. {sheet.tip}", target=sheet.name, code=sheet.code)
                             for sheet in others], kind="links"),
        ],
    }


def readme() -> dict[str, Any]:
    readme_sheet = next(sheet for sheet in build_demo.sheet_map() if sheet.code == "shReadme")
    return {
        "sheet": readme_sheet.tab,
        "subtitle": readme_sheet.tip,
        "tables": [
            table("Conventions", [
                row("Functions", "Every library function starts with K_STATS_. The repository documents the public "
                                 "API and the map from each Excel function in docs/PUBLIC_API.md."),
                row("Sheet colours", TAB_LEGEND),
                row("Navigation", "Go to, the bar on the left and the links on Home open a sheet by its name. The "
                                  "bar and the links work with macros disabled."),
                row("Protection", "Home and README are protected without a password, and the workbook structure is "
                                  "locked so that the links keep working. This is an editing aid, not a security "
                                  "boundary."),
            ]),
            table("Build", [
                row("Library", "VBA Probability Distributions, version {version}: the production modules of "
                               "src/modules, embedded as the repository holds them."),
                row("Source commit", "{commit}"),
                row("Built by", "tools/demobuild/explorer_build.py puts a VBA project built from the sources of this "
                                "commit into src/workbook/PROBDIST_Template.xlsx, without Excel. The Evolve scripts "
                                "of tools/demobuild then write each sheet in desktop Excel."),
                row("Built in", "{excel}"),
            ]),
            table("Validation and limits", [
                row("Workbook checks", "The Distribution Explorer workflow checks every sheet in desktop Excel "
                                       "against its stored references or documented outcomes (ci/Test-*.ps1), with "
                                       "the Ribbon, the navigation, the resets and the save. Its reports record what "
                                       "that Excel returned."),
                row("Library validation", "The numerical certification of the library is separate: "
                                          "docs/EXCEL_CERTIFICATION.md and the evidence of the repository. This "
                                          "workbook is educational and integration evidence, not a certification."),
                row("Tested", "{excel}. Other builds, older versions, 32-bit Excel and Excel for Mac are not "
                              "tested."),
                row("Layout", "The page, the styles, the navigation bar and the Ribbon structure follow the K-SACCR "
                              "workbook template, under the MPL-2.0."),
            ]),
        ],
    }


def plan() -> dict[str, Any]:
    return {"first_row": FIRST_ROW, "version": version(), "swatches": list(SWATCHES), "sheets": [home(), readme()]}


def main() -> int:
    json.dump(plan(), sys.stdout, indent=1)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
