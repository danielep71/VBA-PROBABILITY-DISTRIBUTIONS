# The JSON spec of `xlsm_tool.py build`

Paths inside the spec are relative to the spec file; the `-o` output is relative to the current directory. Every key except `sheets[].name` is optional.

```json
{
  "sheets": [
    {
      "name": "Home",
      "code_name": "shHome",
      "cells": {"A1": "Value", "B1": 41, "A2": "Plus one", "B2": "=AddOne(B1)", "C1": true},
      "bold": ["A1", "A2"],
      "formats": {"B1": "#,##0.00", "B3": "0.00%", "B4": "yyyy-mm-dd"},
      "widths": {"A": 18},
      "code": "shHome.vba"
    },
    {"name": "Data"}
  ],
  "modules": ["src/*.bas"],
  "workbook_code": "ThisWorkbook.vba",
  "ribbon": "customUI14.xml",
  "defined_names": {"Start": "'Home'!$A$1"},
  "project_id": "{00000000-0000-0000-0000-000000000000}"
}
```

| Key | Meaning |
| --- | --- |
| `sheets[].name` | Tab name: 1 to 31 characters, none of `[ ] : * ? / \`, unique. |
| `sheets[].code_name` | The sheet's VBA CodeName. Default `Sheet1`, `Sheet2`, … by position. |
| `sheets[].cells` | Cell values by A1 reference. A number is a number, `true`/`false` a Boolean, a string starting with `=` a formula (English function names, comma separators), any other string text. |
| `sheets[].bold` | References written in bold. |
| `sheets[].formats` | Number format code by reference, as Excel's Format Cells > Custom shows it (`0.00%`, `#,##0.00`, `yyyy-mm-dd hh:mm`). Bold and number formats are the only formatting of a plain workbook. |
| `sheets[].widths` | Column widths in characters, by column letter. |
| `sheets[].code` | File with the code of the sheet's module (event procedures). Every sheet gets a module and a CodeName whether or not it has code, so other modules can refer to it. |
| `modules` | Standard modules, as paths or glob patterns, in project order. |
| `workbook_code` | File with the code of `ThisWorkbook`. |
| `ribbon` | RibbonX file. |
| `defined_names` | Workbook-level names and what they refer to. |
| `project_id` | Project GUID in braces. Default: derived from the code, so the same code gives the same file. |

A formula that calls a function of the embedded VBA shows `#NAME?` until Excel has the macros
enabled and recalculates; that is expected in a file that Excel has never opened. LibreOffice
shows `#VALUE!` or `#NAME?` for the same cells, because it does not run the VBA.

For anything richer than plain cells (formats, charts, tables, validation, images), build the
workbook with a spreadsheet library or start from an existing `.xlsx`, then use `inject`.
