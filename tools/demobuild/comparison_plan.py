"""The Comparison sheet (Excel vs VBA vs Python) as data: every formula, name, validation rule and text.

Same method as probabilities_plan.py: examples/demo/comparison-cases.json is the one source of the
cases, and this module writes from it the plan that tools/demobuild/Evolve-Comparison.ps1 applies
in desktop Excel. The cases are the 29 rows of the benchmark/excel_comparison study, with the
mpmath reference at 50 digits and the SciPy call and value that comparison_references.py adds.

Layout, in the style of the Explorer:

    C6:D9     the inputs: the case shown in detail, the two grade thresholds (correct digits) and
              an optional cell for the value of an =PY() formula
    D19:D20   the input status and the macros
    C21:I35   the selected case: the Excel formula and value, the library formula and value, the
              reference, the Python code and its stored SciPy value, each implementation against
              the reference (difference, relative difference, correct digits, grade), why the case
              is in the study, the Excel that computed the sheet and the provenance
    F5:I17    the notes; K5:O35 a bar chart of the correct digits of each case
    C37:O66   every case: the reference, the three values, their correct digits and grades
    R:T       hidden: the reference as 50-digit text, the Python code and the reason for the case

Each implementation is graded against the mpmath reference only: Excel is one of the three, never
the yardstick. Correct digits are -log10 of the relative difference, capped at 16 (a Double holds
about 16), and 16 when the value equals the reference in Double. A zero reference has no relative
difference: there only an exact 0 counts, as 16 digits, and anything else as 0. The grades are the
study's: reference from 12 digits, report from 6, inadequate below, wrong at 0; the two thresholds
are inputs.

The Excel and library columns are live formulas (F and G), written exactly as the study lists them;
a case where Excel has no equivalent function holds the text "none". The SciPy column is a stored
value: the sheet works without Python. The Python code of the selected case can be pasted into an
=PY() formula in D9, which the sheet then grades like the others.
"""
from __future__ import annotations

import json
from decimal import Decimal
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parents[2]
CASES_FILE = ROOT / "examples" / "demo" / "comparison-cases.json"

SHEET = "Comparison"
CASE, REFERENCE_GRADE, REPORT_GRADE, PYTHON_VALUE = "$D$6", "$D$7", "$D$8", "$D$9"
STATUS = "$D$19"
MACROS = "D20"
HEADER_ROW = 37
FIRST_ROW = 38
FULL_DIGITS = 16
DETAIL_ROWS = {
    "excel_formula": 21, "excel_value": 22, "library_formula": 23, "library_value": 24,
    "reference": 25, "python_code": 26, "scipy_value": 27, "excel_grade": 28,
    "library_grade": 29, "scipy_grade": 30, "python_grade": 31, "why": 32, "excel_build": 34,
    "provenance": 35,
}
# Visible columns of the table, and the hidden ones that hold the texts of each case.
COLUMNS = {"label": "C", "reference": "D", "excel": "F", "library": "G", "scipy": "H",
           "excel_digits": "I", "library_digits": "K", "scipy_digits": "L",
           "excel_grade": "M", "library_grade": "N", "scipy_grade": "O",
           "reference_text": "R", "python": "S", "why": "T"}
NAMES = {"label": "ComparisonLabels", "reference": "ComparisonReference", "excel": "ComparisonExcel",
         "library": "ComparisonLibrary", "scipy": "ComparisonSciPy",
         "excel_digits": "ComparisonExcelDigits", "library_digits": "ComparisonLibraryDigits",
         "scipy_digits": "ComparisonSciPyDigits", "excel_grade": "ComparisonExcelGrade",
         "library_grade": "ComparisonLibraryGrade", "scipy_grade": "ComparisonSciPyGrade",
         "reference_text": "ComparisonReferenceText", "python": "ComparisonPython", "why": "ComparisonWhy"}
NONE_TEXT = "none"


def load() -> dict[str, Any]:
    data: dict[str, Any] = json.loads(CASES_FILE.read_text(encoding="utf-8"))
    labels = [case["label"] for case in data["cases"]]
    if len(labels) != len(set(labels)):
        raise ValueError("duplicate case label")
    if data["default_case"] not in labels:
        raise ValueError("the default case is not a case")
    return data


def ref(address: str) -> str:
    return f"{SHEET}!{address}"


def last_row(data: dict[str, Any]) -> int:
    return FIRST_ROW + len(data["cases"]) - 1


def digits_formula(value: str, reference: str) -> str:
    """Correct significant digits of a value against the reference: -log10 of the relative
    difference, from 0 to 16; #N/A when the value is not a number (no function, or an error)."""
    return (f"LET(obs,{value},truth,{reference},IF(NOT(ISNUMBER(obs)),NA(),"
            f"LET(rel,IF(truth=0,IF(obs=0,0,1),ABS(obs-truth)/ABS(truth)),"
            f"IF(rel=0,{FULL_DIGITS},MIN({FULL_DIGITS},MAX(0,-LOG10(rel)))))))")


def grade_formula(value: str, digits: str) -> str:
    """The study's grade of a value from its correct digits and the two thresholds."""
    return (f'IF({STATUS}<>"OK","-",IF(ISTEXT({value}),"no function",IF(ISERROR({value}),"error",'
            f'IF({digits}<=0,"wrong",IF({digits}>={REFERENCE_GRADE},"reference",'
            f'IF({digits}>={REPORT_GRADE},"report","inadequate"))))))')


def scientific(expression: str) -> str:
    """A number as 6.22E-16, with the regional decimal separator: TEXT would need a format code
    in the language of the user's settings."""
    return (f'IF({expression}=0,"0",LET(ex,FLOOR.MATH(LOG10(ABS({expression}))),m,ROUND({expression}/10^ex,2),'
            f'IF(ABS(m)>=10,ROUND(m/10,2)&"E"&(ex+1),m&"E"&ex)))')


def summary_formula(value: str, digits: str, grade: str) -> str:
    """One line on a value against the reference: the difference, the relative difference, the
    correct digits and the grade. A zero reference has no relative difference."""
    truth = "INDEX(ComparisonReference,ComparisonRow)"
    return (f'=IFERROR(LET(obs,{value},truth,{truth},dig,{digits},grade,{grade},'
            f'IF(ISNUMBER(obs),"difference "&{scientific("ABS(obs-truth)")}&", relative "&'
            f'IF(truth=0,"undefined (zero reference)",{scientific("ABS(obs-truth)/ABS(truth)")})&": "&'
            f'ROUND(dig,1)&" correct digits, "&grade,grade)),"")')


def detail_formulas() -> list[dict[str, str]]:
    row = "ComparisonRow"

    def pick(key: str) -> str:
        return f"INDEX({NAMES[key]},{row})"

    python_digits = digits_formula(PYTHON_VALUE, pick("reference"))
    python_grade = grade_formula(PYTHON_VALUE, "dig")
    at = {key: f"D{number}" for key, number in DETAIL_ROWS.items()}
    return [
        {"address": at["excel_formula"],
         "formula": (f'=IFERROR(IF(ISTEXT({pick("excel")}),"none: the study lists no equivalent Excel function",'
                     f'FORMULATEXT({pick("excel")})),"")')},
        {"address": at["excel_value"], "formula": f'=IFERROR({pick("excel")},"")'},
        {"address": at["library_formula"], "formula": f'=IFERROR(FORMULATEXT({pick("library")}),"")'},
        {"address": at["library_value"], "formula": f'=IFERROR({pick("library")},"")'},
        {"address": at["reference"], "formula": f'=IFERROR({pick("reference_text")},"")'},
        {"address": at["python_code"], "formula": f'=IFERROR({pick("python")},"")'},
        {"address": at["scipy_value"], "formula": f'=IFERROR({pick("scipy")},"")'},
        {"address": at["excel_grade"],
         "formula": summary_formula(pick("excel"), pick("excel_digits"), pick("excel_grade"))},
        {"address": at["library_grade"],
         "formula": summary_formula(pick("library"), pick("library_digits"), pick("library_grade"))},
        {"address": at["scipy_grade"],
         "formula": summary_formula(pick("scipy"), pick("scipy_digits"), pick("scipy_grade"))},
        {"address": at["python_grade"],
         "formula": (f'=IF(ISBLANK({PYTHON_VALUE}),"not used: the stored SciPy value above stands in",'
                     f'IF(NOT(ISNUMBER({PYTHON_VALUE})),"D9 is not a number: set the =PY() output to Excel value",'
                     + summary_formula(PYTHON_VALUE, python_digits, python_grade)[1:] + "))")},
        {"address": at["why"], "formula": f'=IFERROR({pick("why")},"")'},
    ]


def table(data: dict[str, Any]) -> tuple[list[dict[str, Any]], list[dict[str, str]]]:
    """The values written once (labels, references, stored SciPy values, texts) and the formulas
    of every row: the live Excel and library calls, the correct digits and the grades."""
    cells: list[dict[str, Any]] = []
    formulas: list[dict[str, str]] = []
    for index, case in enumerate(data["cases"]):
        row = FIRST_ROW + index
        at = {key: f"{column}{row}" for key, column in COLUMNS.items()}
        cells += [
            {"address": at["label"], "value": case["label"]},
            {"address": at["reference"], "value": float(Decimal(case["reference"]))},
            {"address": at["scipy"], "value": float(case["scipy"])},
            # Texts that Excel would read as numbers keep the apostrophe of a text constant.
            {"address": at["reference_text"], "value": "'" + case["reference"]},
            {"address": at["python"], "value": case["python"]},
            {"address": at["why"], "value": case["why"]},
        ]
        if case["excel"] is None:
            cells.append({"address": at["excel"], "value": NONE_TEXT})
        else:
            formulas.append({"address": at["excel"], "formula": "=" + case["excel"]})
        formulas.append({"address": at["library"], "formula": "=" + case["library"]})
        for side in ("excel", "library", "scipy"):
            formulas.append({"address": at[f"{side}_digits"],
                             "formula": "=" + digits_formula(at[side], f"${COLUMNS['reference']}{row}")})
            formulas.append({"address": at[f"{side}_grade"],
                             "formula": "=" + grade_formula(at[side], at[f"{side}_digits"])})
    return cells, formulas


def status_formula() -> str:
    return (f'=IF(ISNA(ComparisonRow),"Check: case",'
            f'IF(NOT(AND(ISNUMBER({REFERENCE_GRADE}),ISNUMBER({REPORT_GRADE}))),"Check: grade thresholds",'
            f'IF(OR({REPORT_GRADE}<1,{REFERENCE_GRADE}>{FULL_DIGITS},{REPORT_GRADE}>{REFERENCE_GRADE}),'
            f'"Check: grade thresholds","OK")))')


def static_messages(data: dict[str, Any]) -> dict[str, list[str]]:
    """Input title, input message, error title and error message of the input cells."""
    count = len(data["cases"])
    return {
        "D6": ["Case", f"Choose one of the {count} cases of the study: its formulas, its Python code and the "
               "three implementations against the reference are shown below.",
               "Unknown case", "Choose a case from the list."],
        "D7": ["Reference grade from", "Correct significant digits from which a value is reference grade "
               f"(the study uses 12). From 1 to {FULL_DIGITS}, not below the report threshold.",
               "Invalid threshold", f"Enter a number of digits from 1 to {FULL_DIGITS}, not below D8."],
        "D8": ["Report grade from", "Correct significant digits from which a value is fit to report "
               f"(the study uses 6). From 1 to {FULL_DIGITS}, not above the reference threshold.",
               "Invalid threshold", f"Enter a number of digits from 1 to {FULL_DIGITS}, not above D7."],
        "D9": ["Your =PY() value", "Optional, with Python in Excel (Microsoft 365): type =PY( here, paste the "
               "Python code of D26, press Ctrl+Enter and set the output to Excel value. Empty, the stored "
               "SciPy value stands in.", "", ""],
    }


def plan() -> dict[str, Any]:
    data = load()
    end = last_row(data)
    cells, table_formulas = table(data)
    names = [{"name": name, "refers": f"={ref(f'${COLUMNS[key]}${FIRST_ROW}:${COLUMNS[key]}${end}')}"}
             for key, name in NAMES.items()]
    names += [
        {"name": "ComparisonRow", "refers": f"=IFERROR(MATCH({ref(CASE)},ComparisonLabels,0),NA())"},
        {"name": "ComparisonStatusBad", "refers": f'={ref(STATUS)}<>"OK"'},
    ]
    formulas = [
        {"address": STATUS.replace("$", ""), "formula": status_formula()},
        {"address": MACROS, "formula": '=IF(ISERROR(K_STATS_NormalStandard_Cumulative(0)+0*COUNTA($D$6:$D$9)),"DISABLED","ENABLED")'},
        {"address": "F4", "formula": f'=IF(${MACROS[0]}${MACROS[1:]}="ENABLED","","MACROS ARE DISABLED: the results are not live")'},
        *detail_formulas(),
        *table_formulas,
    ]
    versions = data["scipy"]
    provenance = (f'Reference: {data["reference"]["method"]} (mpmath {versions["mpmath"]}). SciPy {versions["scipy"]}, '
                  f'NumPy {versions["numpy"]}, Python {versions["python"]}: stored by examples/demo/comparison_references.py.')
    texts = [
        {"address": "B1", "value": "Excel vs VBA vs Python"},
        {"address": "B2", "value": "Excel's own functions, the library and SciPy on the same inputs, each against a 50-digit reference."},
        {"address": "C3", "value": "Choose a case to read its formulas and Python code. Every library value is a call to the embedded "
                                   "VBA API (K_STATS_ functions); the Excel values are Excel's own functions. Macros are required."},
        {"address": "C6", "value": "Case"},
        {"address": "C7", "value": "Reference grade from (digits)"},
        {"address": "C8", "value": "Report grade from (digits)"},
        {"address": "C9", "value": "Your =PY() value (optional)"},
        {"address": "C19", "value": "Input status"},
        {"address": "C20", "value": "Macros"},
        {"address": "C21", "value": "Excel formula"},
        {"address": "C22", "value": "Excel value"},
        {"address": "C23", "value": "Library formula"},
        {"address": "C24", "value": "Library value"},
        {"address": "C25", "value": "Reference (mpmath, 50 digits)"},
        {"address": "C26", "value": "Python code (SciPy)"},
        {"address": "C27", "value": "SciPy value (stored)"},
        {"address": "C28", "value": "Excel against the reference"},
        {"address": "C29", "value": "Library against the reference"},
        {"address": "C30", "value": "SciPy against the reference"},
        {"address": "C31", "value": "=PY() against the reference"},
        {"address": "C32", "value": "Why this case"},
        {"address": "C34", "value": "Computed by"},
        {"address": "C35", "value": "Provenance"},
        {"address": "D35", "value": provenance},
        {"address": "F5", "value": "Read the results"},
        {"address": "F6", "value": "Yellow cells are editable. Choose a case to read its Excel formula, its library call and "
                                   "the Python code with the same inputs and tail."},
        {"address": "F8", "value": "Each value is graded against the mpmath reference only. Excel is one of the three "
                                   "implementations, not the yardstick."},
        {"address": "F10", "value": f"Correct digits: -log10 of the relative difference, at most {FULL_DIGITS}. A zero "
                                    "reference has no relative difference: only an exact 0 counts."},
        {"address": "F12", "value": "The SciPy values are stored, so no Python is needed. With Python in Excel, paste the "
                                    "code of D26 into an =PY() formula in D9."},
        {"address": "F14", "value": "A grade below reference is a difference from this reference, not a certified defect. "
                                    "The library's contracts are in docs/."},
        {"address": "F16", "value": "Needs macros and Excel 365 or 2024. If Macros says DISABLED, unblock the file and reopen."},
        {"address": f"C{HEADER_ROW - 1}", "value": "Every case: values, correct digits and grades"},
        {"address": f"C{HEADER_ROW}", "value": "Case"},
        {"address": f"D{HEADER_ROW}", "value": "Reference (mpmath)"},
        {"address": f"F{HEADER_ROW}", "value": "Excel"},
        {"address": f"G{HEADER_ROW}", "value": "Library (K_STATS_)"},
        {"address": f"H{HEADER_ROW}", "value": "SciPy (stored)"},
        {"address": f"I{HEADER_ROW}", "value": "Excel digits"},
        {"address": f"K{HEADER_ROW}", "value": "Library digits"},
        {"address": f"L{HEADER_ROW}", "value": "SciPy digits"},
        {"address": f"M{HEADER_ROW}", "value": "Excel grade"},
        {"address": f"N{HEADER_ROW}", "value": "Library grade"},
        {"address": f"O{HEADER_ROW}", "value": "SciPy grade"},
    ]
    detail = [f"D{number}" for number in DETAIL_ROWS.values()]
    return {
        "schema": "comparison-plan-1",
        "sheet": SHEET,
        "names": names,
        "cells": cells,
        "formulas": formulas,
        "texts": texts,
        "defaults": {"case": data["default_case"], "reference_grade": data["grades"]["reference"],
                     "report_grade": data["grades"]["report"]},
        "messages": static_messages(data),
        "layout": {"case": "D6", "reference_grade": "D7", "report_grade": "D8", "python": "D9",
                   "status": STATUS.replace("$", ""), "macros": MACROS, "inputs": "D6:D9",
                   "detail_rows": DETAIL_ROWS, "detail_merge": "D{row}:I{row}", "why_merge": "D32:I33",
                   "header_row": HEADER_ROW, "first_row": FIRST_ROW, "last_row": end,
                   "table": f"C{HEADER_ROW}:O{end}", "values": f"D{FIRST_ROW}:D{end},F{FIRST_ROW}:H{end}",
                   "digits": f"I{FIRST_ROW}:I{end},K{FIRST_ROW}:L{end}", "grades": f"M{FIRST_ROW}:O{end}",
                   "hidden_texts": f"R{FIRST_ROW}:T{end}", "chart": "K5:O35",
                   "outputs": ["D19:D20", "D21:I35", f"C{HEADER_ROW}:O{end}"], "detail": detail,
                   "spacers": ["C10:D11", "C12:D13", "C14:D15", "C16:D16"]},
        "validation": {"case_list": "=ComparisonLabels", "digits": [1, FULL_DIGITS]},
        "conditional": {"status_bad": "=ComparisonStatusBad"},
        "chart": {"name": "ComparisonDigits", "title": "Correct digits against the 50-digit reference",
                  "categories": "ComparisonLabels",
                  "series": [{"name": f"I{HEADER_ROW}", "values": "ComparisonExcelDigits"},
                             {"name": f"K{HEADER_ROW}", "values": "ComparisonLibraryDigits"},
                             {"name": f"L{HEADER_ROW}", "values": "ComparisonSciPyDigits"}]},
        "full_digits": FULL_DIGITS,
    }


def formula_lengths(result: dict[str, Any]) -> list[tuple[str, int]]:
    """Length of every formula and name, to check them against Excel's limit of 8192 characters."""
    sizes = [(item["address"], len(item["formula"])) for item in result["formulas"]]
    sizes += [(item["name"], len(item["refers"])) for item in result["names"]]
    return sorted(sizes, key=lambda item: -item[1])


if __name__ == "__main__":
    for name, size in formula_lengths(plan())[:6]:
        print(size, name)
