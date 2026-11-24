---
name: excel-vba-workbook
description: Build, modify and inspect macro-enabled Excel workbooks (.xlsm) that contain VBA, without Excel installed. Use this whenever the task involves VBA code inside a workbook - creating an .xlsm from .bas modules, adding macros, user-defined functions (UDFs), workbook or sheet event code, or a custom Ribbon tab (customUI / RibbonX) to an .xlsx; turning a formatted spreadsheet into a macro-enabled one; extracting or reading the VBA of an existing .xlsm; or checking why Excel repairs, refuses or ignores a generated workbook. Also use it when the user says "macro", "xlsm", "vbaProject", "ThisWorkbook", "CodeName" or "Ribbon callback" about an Excel file, even if they do not ask for a build. Do not use it for spreadsheets without VBA (plain .xlsx or .csv), or for Word or PowerPoint macros.
---

# Excel workbooks with VBA, without Excel

An `.xlsm` is a zip of XML parts plus one binary part, `xl/vbaProject.bin`, that holds the VBA.
Spreadsheet libraries can copy that binary from an existing file but cannot write one from
source. `scripts/xlsm_tool.py` can: it writes the VBA project from `.bas` text, following the
[MS-OVBA] and [MS-CFB] specifications, as compressed source with no compiled cache, so Excel
compiles it when the workbook opens. It needs only the Python standard library.

## What to keep in mind the whole time

**Nothing here runs Excel, so nothing here proves the VBA works.** The tool proves that the
package is well formed and that names agree with each other. Whether the code compiles, whether
a formula recalculates, whether a Ribbon button does its job: only Excel on Windows or macOS
can say. Tell the user this plainly, every time you hand over a workbook, and say what they
should try first (open it, enable macros, Debug > Compile VBAProject in the VBA editor). Do not
write "tested" or "verified" about VBA behaviour you have not seen run.

Because nothing compiles the code for you, read your VBA as a compiler would before you ship it:
every variable declared (`Option Explicit` in every module), every `Sub`/`Function` closed,
every called procedure present, no `Dim` after use of the name in the same scope, line
continuations ending in ` _`. Avoid parameter and variable names that shadow VBA's own
(`Rate`, `Date`, `Name`, `Address`, `Value`, `Count`): they are legal, and they produce the
confusing errors. The tool checks only that each procedure is closed and that `Option
Explicit` is present.

## Choosing the command

```
python scripts/xlsm_tool.py inject  BOOK.xlsx -o OUT.xlsm -m MODULES...   # VBA into an existing workbook
python scripts/xlsm_tool.py build   SPEC.json -o OUT.xlsm                 # plain workbook + VBA from a spec
python scripts/xlsm_tool.py inspect BOOK.xlsm [--check] [--extract DIR]   # list, check, extract
```

- The workbook should **look like something** (formats, charts, tables, a template the user
  gave you): make or take the `.xlsx` first, by whatever means suit it (a spreadsheet library,
  the user's file), then `inject`. Every part of the original is kept byte for byte; only the
  content type, one relationship and the sheet CodeNames change.
- The workbook is **mostly a carrier for code** (labelled cells, numbers, formulas, bold and
  number formats, and the macros): `build` from a JSON spec. Read `references/spec.md` for the
  keys. Modules, sheet code, CodeNames and the Ribbon are options of `inject` and keys of the
  spec; `build` takes only the spec and `-o`.
- You were **given an .xlsm** and need to know what is in it, or to edit its code: `inspect
  --extract DIR` writes every module to a file. Edit the `.bas` files, then `inject` them back
  into the same workbook. Injection replaces the whole VBA project: pass every module that
  should remain, not only the one you changed.

Each build ends with the same report `inspect` prints. Read it: lines starting `PROBLEM:` are
things Excel would refuse, repair or report, and the exit code is 1 when there is one.

## Writing the VBA

**Standard modules** are `.bas` files. Start each with `Attribute VB_Name = "ModuleName"` (if
the line is missing the tool uses the file name). Names are a letter followed by letters,
digits or underscores, 31 characters at most, and unique across modules, sheets and
`ThisWorkbook`.

**Code for ThisWorkbook and for sheets** (event procedures such as `Workbook_Open` or
`Worksheet_Change`) goes in its own text file, passed with `--workbook-code FILE` or
`--sheet-code "Sheet name=FILE"` (`workbook_code` and `sheets[].code` in a spec). Write only
the procedures: the tool adds the module header and `Option Explicit`, and drops an `Option
Explicit` of your own, so there is never a duplicate. An exported `.cls` of a document module
is accepted as it is. Every sheet gets a module whether or not it has code.

In a `Worksheet_Change` handler remember that `Target` can be many cells, that a handler
which writes to its own sheet must wrap the write in `Application.EnableEvents = False` /
`True` (restored in an error handler too), and that any write by a macro empties the user's
Undo history; say so if it matters to them.

**A function for worksheet formulas** (a UDF) must be `Public`, in a standard module, and that
module must not have `Option Private Module`, or the formula bar will not see it.

**Text encoding.** VBA modules are stored in Windows-1252. The tool refuses a character outside
it and names the line. Keep source ASCII where you can; for other characters in a string use
`ChrW(code)`.

**Not supported:** class modules and UserForms. If the design needs one, say so; do not fake it
with a standard module that changes the behaviour silently. Chart sheets are refused too: put
charts on worksheets.

**Library references.** The project references OLE Automation and the Microsoft Office object
library, so `IRibbonUI`, `IRibbonControl` and the `mso…` constants are available early-bound.
Anything else (Scripting.Dictionary, ADODB, Outlook) must be late-bound:
`CreateObject("Scripting.Dictionary")` with variables `As Object`.

## CodeNames

Every sheet has a tab name the user sees and a CodeName the VBA sees (`shHome.Range("A1")`).
The tool keeps a sheet's existing CodeName, or assigns `Sheet1`, `Sheet2`, … by position; set
your own with `--code-name "Tab name=shHome"` (or `code_name` in a spec). Code that refers to
sheets by CodeName survives the user renaming a tab; code that uses `Worksheets("Home")` does
not. Prefer CodeNames, and prefer `ThisWorkbook` to `ActiveWorkbook` so that the code never
touches another open file.

## A Ribbon tab

Pass a RibbonX file with `--ribbon FILE` (key `ribbon` in a spec). The root is `<customUI>` in
the namespace `http://schemas.microsoft.com/office/2009/07/customui`. Every callback named in
it (`onAction`, `getEnabled`, `onLoad`, …) must be a `Public Sub` in a standard module with the
signature Excel expects. A minimal file:

```xml
<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui" onLoad="RibbonOnLoad">
  <ribbon><tabs>
    <tab id="myTab" label="My tools">
      <group id="myGroup" label="Actions">
        <button id="myRun" label="Run" size="large" imageMso="CalculateNow" onAction="RibbonAction"/>
      </group>
    </tab>
  </tabs></ribbon>
</customUI>
```

Every `id` is unique in the file. `imageMso` names one of Office's built-in icons (`FileOpen`,
`CalculateNow`, `Undo`, `Help`); a wrong name gives a blank icon, not an error. The tool checks
that the file is well-formed and that each callback has a procedure; it does not validate
elements or attributes against the RibbonX schema, and a schema error makes Excel drop the
whole tab silently (File > Options > Advanced > "Show add-in user interface errors" makes it
say why). The signatures:

```vb
Public Sub RibbonOnLoad(ByVal Ribbon As IRibbonUI)
Public Sub RibbonAction(ByVal Control As IRibbonControl)                              ' button onAction
Public Sub RibbonToggle(ByVal Control As IRibbonControl, ByVal Pressed As Boolean)    ' toggleButton onAction
Public Sub RibbonGetEnabled(ByVal Control As IRibbonControl, ByRef ReturnedVal As Variant)   ' any get…
```

A wrong signature is not caught here: Excel reports it when the control is used. The tool does
catch a callback with no procedure of that name. Keep callbacks thin: read `Control.Id`, call
an ordinary procedure, and handle errors with a message, because an unhandled error in a
callback disables the control for the session.

## Mistakes that make Excel repair or refuse a file

`inspect --check` looks for these; they are worth knowing when you edit package XML yourself.

- **Macros in a file named .xlsx.** Excel refuses it. The output must end in `.xlsm`.
- **Content type and VBA part disagree.** A `vbaProject.bin` without the macro-enabled content
  type, or the reverse.
- **A sheet whose CodeName has no document module**, or the reverse. Excel invents modules and
  may rename them, and event code ends up unbound.
- **`mc:Ignorable` naming a prefix that is not declared.** Common after an XML library
  rewrites a part and drops unused namespaces. XML parsers accept it; Excel calls the file
  corrupt. When you rewrite a sheet with an XML library, either keep the declarations or
  remove the `mc:Ignorable` attribute together with the elements it covered.
- **Element order.** In a worksheet, `sheetPr` comes first and `hyperlinks` before
  `pageMargins`; in a workbook, `workbookPr` before `bookViews`. Excel repairs or refuses a
  part whose children are out of schema order even when the XML is well formed.

## Checking the result

1. Run `inspect --check` on the output (a build already does). Fix every `PROBLEM`.
2. If LibreOffice is available, convert to PDF
   (`soffice --headless --convert-to pdf --outdir DIR OUT.xlsm`) and look at the pages. This
   shows cells, formats and shapes; it does not run VBA, so a cell that calls one of the
   workbook's functions shows `#VALUE!` or `#NAME?` there, and it does not show the Ribbon.
3. Extract and compare when you edited existing code: `inspect --extract DIR`, then diff
   against your sources, so that what is in the workbook is what you meant to put there.
   Extracted files have CRLF line endings (strip `\r` before a diff against LF sources), and
   the `.cls` of a sheet or of ThisWorkbook carries the header the tool added.

## Handing over

Give the user the `.xlsm` and say, briefly:

- what is in it (sheets, modules, Ribbon);
- that the VBA has not been compiled or run, and the first things to try in Excel;
- that a file downloaded from the web or received by mail opens with macros blocked: on
  Windows, right-click the file, Properties, tick Unblock, or put it in a trusted location;
- that a formula calling one of the workbook's own functions shows `#NAME?` until macros are
  enabled and the sheet recalculates.

If they report a compile error or a repair message, ask for the exact text and the line Excel
highlights, fix the source, and rebuild: the build is deterministic, so the same sources give
the same file.

## Files

- `scripts/xlsm_tool.py` - the three commands. `-h` on each for options.
- `scripts/vba_project.py` - the VBA project writer and reader the tool uses. Import it
  directly (`build_vba_project`, `Module`, `document_module`, `CompoundFile`, `decompress`)
  when a project needs its own builder rather than the tool.
- `references/spec.md` - the JSON spec of `build`. Read it before writing a spec.
