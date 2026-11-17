#!/usr/bin/env python3
"""Build the development demo workbook, without Excel or external packages.

Run from the repository: python tools/demobuild/build_demo.py
Output is ignored under test-results/demo/. This is not a certified release.

The workbook follows the K-SACCR workbook template
(https://github.com/danielep71/K-SACCR, src/workbook/SACCR_Template.xlsx):

* the style table and the theme are that template's own parts, kept under
  examples/demo/reference/ and written unchanged;
* every sheet has the template's page: a black band with the project name, a
  blue band with the sheet title, a subtitle, and bordered two-column tables
  under section headings, with the template's cell styles, row heights,
  column widths, frozen rows 1 to 3 and column A, and landscape print setup;
* every sheet carries the template's navigation bar in column A: linked
  shapes in the sections of the Ribbon's Go to menu, the first section after
  Start in the frozen rows, the others below with Start last, the current
  sheet highlighted, a strip in each sheet's tab colour, not printed;
* the Ribbon tab has the template's Navigate group: a Go to split button
  with a sectioned menu and the Navigation bar toggle;
* each role has the template's tab colour, sheets are protected without a
  password and the workbook structure is locked.

examples/demo/sheets.json is the one map of sections, tab names, CodeNames,
tooltips and issues: the Ribbon menu, the navigation bar, the Home table and
the defined names of the links are all written from it.
"""
from __future__ import annotations

import hashlib
import io
import json
import re
import sys
import zipfile
from dataclasses import dataclass
from pathlib import Path
from xml.sax.saxutils import escape, quoteattr

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT / "tools" / "devbuild"))
from build_workbook import (  # noqa: E402
    MAIN,
    PACKAGE_RELATIONSHIPS,
    RELATIONSHIPS,
    XML_HEADER,
    ZIP_TIME,
    git,
    module_text,
    project_id,
)
from vba_project import DOCUMENT, Module, build_vba_project, document_module  # noqa: E402

DEMO = ROOT / "examples" / "demo"
PROJECT = "Probability Distributions"
WORKBOOK_NAME = "Probability-Distributions-Demo.xlsm"
RIBBON_TAB = "pdTab"
GO_MENU = "pdGoMenu"
CUSTOM_UI = "http://schemas.microsoft.com/office/2009/07/customui"
CONTENT_TYPES = "http://schemas.openxmlformats.org/package/2006/content-types"
VBA_TYPE = "http://schemas.microsoft.com/office/2006/relationships/vbaProject"
RIBBON_TYPE = "http://schemas.microsoft.com/office/2007/relationships/ui/extensibility"
DRAWING_NS = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
ART_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"

# Tab colour of each role, as docs/workbook/sheet-style.json of the template gives them.
TAB_COLORS = {
    "Instructions": "1F3864",
    "Reference": "1F3864",
    "Input": "BF9000",
    "Output": "2E75B6",
    "Test": "548235",
}

# ----------------------------------------------- cell styles of the template ----
# Indices into the cellXfs of examples/demo/reference/styles.xml, as the
# template's Home sheet uses them.
S_PROJECT = (24, 25, 26)        # black band: name, filler, filler
S_TITLE = (28, 29, 30)          # blue band: sheet title, filler, filler
S_SUBTITLE = 2
S_HEADING = (31, 35)            # section heading and its rule
S_SPACER = (33, 34)
S_LABEL = (104, 105, 106)       # first, middle and last row of a table, column C
S_LINK = (168, 169, 170)        # the same for a linked label
S_TEXT = (108, 109, 110)        # the same for column D
S_NAV_COLUMN = 171              # column A under the navigation bar

ROW_PROJECT = "24"
ROW_TITLE = "26.1"
ROW_SUBTITLE = "33.950000000000003"
ROW_GAP = "9.9499999999999993"
ROW_LINE = 30                   # one line of text
ROW_TWO_LINES = 46
ROW_THREE_LINES = 62
TEXT_PER_LINE = 150             # characters of column D on one line

COLS = ('<cols><col min="1" max="1" width="21.7" style="%d" customWidth="1"/>'
        '<col min="2" max="2" width="4" customWidth="1"/>'
        '<col min="3" max="3" width="30" customWidth="1"/>'
        '<col min="4" max="4" width="130" customWidth="1"/>'
        '<col min="5" max="5" width="4.75" customWidth="1"/>'
        '<col min="6" max="24" width="0" hidden="1" customWidth="1"/>'
        '<col min="25" max="16384" width="9.125" hidden="1"/></cols>') % S_NAV_COLUMN

# ------------------------------------- navigation bar design of the template ----
PX = 9525                        # EMU per pixel at 96 dpi
BAR_X = 8                        # left margin of the bar inside column A, px
BAR_W = 141                      # bar width, px
PAD_TOP = 6
PAD_BOTTOM = 3
HEAD_H = 17
BTN_H = 21
BTN_GAP = 2
SECTION_GAP = 7
STRIP_W = 3
TOP_ROWS = 3                     # rows frozen on every sheet
FONT = "Aptos"
HEAD_COLOR = "FFC000"            # gold of the wordmark
BTN_FILL = "262626"
BTN_TEXT = "D9D9D9"
CUR_FILL = "265A99"              # title band blue
CUR_TEXT = "FFFFFF"
CUR_STRIP = "FFC000"
DARK_STRIP = "8FAADC"            # used when a tab colour would vanish on the button
GROUP_TOP = "NavBarTop"
GROUP_REST = "NavBar"

THIS_WORKBOOK_EVENTS = """
Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    'Save the navigation bar shown even while it is hidden.
        M_STATS_PROBDIST_DEMO.NavigationBarBeforeSave

End Sub

Private Sub Workbook_AfterSave(ByVal Success As Boolean)
    'Hide the navigation bar again if it was hidden before the save.
        M_STATS_PROBDIST_DEMO.NavigationBarAfterSave

End Sub
"""


@dataclass(frozen=True)
class Sheet:
    tab: str
    code: str
    title: str
    tip: str
    issue: int
    section: str
    role: str

    @property
    def control(self) -> str:
        """Id of the sheet's Go to entry: pdGo and the CodeName without its sh."""
        return "pdGo" + self.code[2:]

    @property
    def name(self) -> str:
        """Defined name of the sheet's link target."""
        return "nav_" + self.code

    @property
    def color(self) -> str:
        return TAB_COLORS[self.role]


Block = list[tuple[str, list[Sheet]]]
Row = tuple[float | str, list[tuple[str, int, str]]]


def sections() -> Block:
    """The Go to menu: [(section title, [sheets]), ...], Start first."""
    data = json.loads((DEMO / "sheets.json").read_text(encoding="utf-8"))
    found: Block = []
    for section in data["sections"]:
        if section["role"] not in TAB_COLORS:
            raise ValueError(f'unknown role {section["role"]!r}')
        found.append((section["title"], [
            Sheet(item["tab"], item["code"], item["title"], item["tip"], int(item["issue"]),
                  section["title"], section["role"]) for item in section["sheets"]]))
    sheets = [sheet for _, items in found for sheet in items]
    for field in ("tab", "code"):
        values = [getattr(sheet, field).lower() for sheet in sheets]
        if len(values) != len(set(values)):
            raise ValueError(f"duplicate {field}")
    for sheet in sheets:
        if not re.fullmatch(r"sh[A-Z][A-Za-z0-9]*", sheet.code):
            raise ValueError(f"{sheet.code}: a CodeName is sh and a capitalized word")
        if not re.fullmatch(r"[A-Za-z0-9_]{1,31}", sheet.tab):
            raise ValueError(f"{sheet.tab}: a tab name is letters, digits and underscores")
    if len(found) < 2 or [sheet.code for sheet in found[0][1]] != ["shHome"]:
        raise ValueError("the first section is Start, with Home alone")
    return found


def sheet_map() -> list[Sheet]:
    """Every sheet in tab order: the order of the Go to menu."""
    return [sheet for _, items in sections() for sheet in items]


def blocks(menu: Block) -> tuple[Block, Block]:
    """The bar's two blocks: the first section after Start in the frozen rows; the rest, Start last."""
    return [menu[1]], menu[2:] + [menu[0]]


# ------------------------------------------------------------------ Ribbon ----
def ribbon(menu: Block) -> str:
    entries = []
    for title, items in menu:
        entries.append('              <menuSeparator id="pdGoSep%s" title=%s/>' % (
            re.sub(r"\W", "", title.title()), quoteattr(title)))
        entries.extend(
            '              <button id="%s" label=%s onAction="RibbonGoTo" supertip=%s/>' % (
                sheet.control, quoteattr(sheet.tab), quoteattr(sheet.tip)) for sheet in items)
    return f"""{XML_HEADER}<!--
  Probability Distributions demo Ribbon. Generated by tools/demobuild/build_demo.py
  from examples/demo/sheets.json; the structure follows the K-SACCR tab. The
  callbacks are in examples/demo/M_STATS_PROBDIST_DEMO.bas and hold no numerical
  logic: RibbonAction runs a command, RibbonGoTo activates a sheet of this
  workbook by its CodeName, RibbonToggle shows or hides the navigation bar, and
  RibbonGetEnabled and RibbonGetPressed only read the workbook.
-->
<customUI xmlns="{CUSTOM_UI}" onLoad="RibbonOnLoad">
  <ribbon>
    <tabs>
      <tab id="{RIBBON_TAB}" label={quoteattr(PROJECT)} keytip="PD">

        <group id="pdGrpCalculation" label="Calculation">
          <button id="pdRecalculate" label="Recalculate" size="large" imageMso="CalculateNow" keytip="R"
                  screentip="Recalculate"
                  supertip="Recalculate every sheet of this workbook. No other open workbook is recalculated."
                  onAction="RibbonAction" getEnabled="RibbonGetEnabled"/>
          <button id="pdResetInputs" label="Reset inputs" imageMso="Undo" keytip="A"
                  screentip="Reset inputs"
                  supertip="Put the example inputs back to their defaults. Not available until the thematic sheets define editable inputs."
                  onAction="RibbonAction" getEnabled="RibbonGetEnabled"/>
        </group>

        <group id="pdGrpNavigate" label="Navigate">
          <splitButton id="pdGoSplit" size="large" keytip="G">
            <button id="pdGo" label="Go to" imageMso="FileOpen"
                    screentip="Go to"
                    supertip="Open Home. The arrow lists every sheet."
                    onAction="RibbonGoTo"/>
            <menu id="{GO_MENU}" itemSize="normal">
{chr(10).join(entries)}
            </menu>
          </splitButton>
          <toggleButton id="pdNavigationBar" label="Navigation bar" keytip="B"
                        screentip="Navigation bar"
                        supertip="Show or hide the bar of sheet links on the left of every sheet, for this session. The workbook is always saved with the bar shown."
                        getPressed="RibbonGetPressed" onAction="RibbonToggle"
                        getEnabled="RibbonGetEnabled"/>
        </group>

      </tab>
    </tabs>
  </ribbon>
</customUI>
"""


# ------------------------------------------------------------------ sheets ----
def text_height(text: str) -> int:
    if len(text) > 2 * TEXT_PER_LINE:
        return ROW_THREE_LINES
    return ROW_TWO_LINES if len(text) > TEXT_PER_LINE else ROW_LINE


def table(rows: list[tuple[str, str]], linked: bool = False) -> list[Row]:
    """A bordered table: label in column C, text in column D. Two rows at least."""
    if len(rows) < 2:
        raise ValueError("a table has two rows at least: the template has no one-row style")
    found: list[Row] = []
    for index, (label, text) in enumerate(rows):
        place = 0 if index == 0 else 2 if index == len(rows) - 1 else 1
        found.append((text_height(text), [("C", (S_LINK if linked else S_LABEL)[place], label),
                                          ("D", S_TEXT[place], text)]))
    return found


def heading(text: str) -> list[Row]:
    return [(ROW_LINE, [("C", S_HEADING[0], text), ("D", S_HEADING[1], "")])]


def spacer() -> list[Row]:
    return [(ROW_GAP, [("C", S_SPACER[0], ""), ("D", S_SPACER[1], "")])]


def band(title: str, subtitle: str) -> list[Row]:
    return [(ROW_PROJECT, [("B", S_PROJECT[0], PROJECT), ("C", S_PROJECT[1], ""), ("D", S_PROJECT[2], "")]),
            (ROW_TITLE, [("B", S_TITLE[0], title), ("C", S_TITLE[1], ""), ("D", S_TITLE[2], "")]),
            (ROW_SUBTITLE, [("C", S_SUBTITLE, subtitle)]),
            (ROW_GAP, [])]


def home_rows(menu: Block) -> tuple[list[Row], list[Sheet]]:
    others = [sheet for _, items in menu[1:] for sheet in items]
    rows = band("Home", "Start here: what this workbook shows, how to use it and a link to every sheet. "
                f"The commands are on the {PROJECT} tab of the Ribbon; its Go to and the bar on the left "
                "open any sheet.")
    rows += heading("What this workbook shows")
    rows += table([
        ("Shows", "The K_STATS_ worksheet functions of VBA Probability Distributions at work: densities, "
                  "probabilities, quantiles and moments of the continuous and discrete families, beside "
                  "Excel's own functions and independent references."),
        ("Does not cover", "Certification. The accuracy of the library is established by the evidence in its "
                           "repository, not by this workbook; the examples are educational."),
        ("Status", "Development shell. The thematic sheets are placeholders, and each names the issue that "
                   "supplies its content. README gives the build and the validation status."),
    ])
    rows += spacer() + heading("How to use")
    rows += table([
        ("1. Choose a sheet", f"Use Go to on the {PROJECT} tab, the bar on the left or the links below. "
                              "The bar highlights the current sheet."),
        ("2. Read the example", "Each thematic sheet will separate inputs, formulas, results and independent "
                                "references."),
        ("3. Recalculate", f"Recalculate on the {PROJECT} tab recalculates this workbook only."),
        ("4. Hide the bar", "Navigation bar on the same tab hides or shows the bar on the left for this "
                            "session. The workbook is always saved with the bar shown."),
        ("Reset inputs", "Not available until the thematic sheets define editable inputs."),
        ("Macros", "The K_STATS_ functions and the commands need macros enabled. With macros disabled the "
                   "workbook can still be read, and the links below and the bar on the left open every "
                   "sheet."),
    ])
    rows += spacer() + heading("Sheets")
    rows += table([(sheet.tab, f"{sheet.section}. {sheet.tip}") for sheet in others], linked=True)
    return rows, others


def readme_rows(sheet: Sheet, commit: str) -> list[Row]:
    version = (ROOT / "VERSION").read_text(encoding="utf-8").strip()
    rows = band(sheet.title, sheet.tip)
    rows += heading("Conventions")
    rows += table([
        ("Functions", "Every library function starts with K_STATS_. The repository documents the public API "
                      "and the map from each Excel function in docs/PUBLIC_API.md."),
        ("Sheet colours", "Navy: instructions and reference. Gold: sheets with inputs to explore. Green: "
                          "accuracy comparisons. Blue: applications."),
        ("Navigation", "Go to, the bar on the left and the links on Home open a sheet by its name. The bar "
                       "and the links work with macros disabled."),
        ("Protection", "The sheets are protected without a password, against accidental edits, and the "
                       "workbook structure is locked so that the links keep working. This is an editing "
                       "aid, not a security boundary."),
    ])
    rows += spacer() + heading("Build")
    rows += table([
        ("Library", f"VBA Probability Distributions, version {version}. The six production modules are "
                    "embedded as the repository holds them."),
        ("Source commit", commit),
        ("Built by", "tools/demobuild/build_demo.py, from source, without Excel. Excel compiles the VBA "
                     "project when the workbook opens."),
    ])
    rows += spacer() + heading("Validation and limits")
    rows += table([
        ("Status", "Development shell: acceptance in desktop Excel is pending. The numerical certification "
                   "of the library is maintained separately, in the repository."),
        ("Layout", "The page, the styles, the navigation bar and the Ribbon structure follow the K-SACCR "
                   "workbook template, under the MPL-2.0."),
    ])
    return rows


def thematic_rows(sheet: Sheet) -> list[Row]:
    rows = band(sheet.title, sheet.tip)
    rows += heading(sheet.title)
    rows += table([
        ("Status", f"Under development. The content is prepared under issue #{sheet.issue}."),
        ("Will show", sheet.tip),
    ])
    return rows


def worksheet(sheet: Sheet, menu: Block, commit: str) -> str:
    links: list[Sheet] = []
    if sheet.code == "shHome":
        rows, links = home_rows(menu)
    elif sheet.code == "shReadme":
        rows = readme_rows(sheet, commit)
    else:
        rows = thematic_rows(sheet)
    link_rows: dict[str, int] = {}
    data = []
    for number, (height, cells) in enumerate(rows, 1):
        written = []
        for column, style, text in cells:
            if text:
                written.append(f'<c r="{column}{number}" s="{style}" t="inlineStr"><is><t>{escape(text)}</t></is></c>')
                if style in S_LINK:
                    link_rows[text] = number
            else:
                written.append(f'<c r="{column}{number}" s="{style}"/>')
        data.append(f'<row r="{number}" ht="{height}" customHeight="1">{"".join(written)}</row>')
    hyperlinks = ""
    if links:
        hyperlinks = "<hyperlinks>" + "".join(
            f'<hyperlink ref="C{link_rows[target.tab]}" location="{target.name}" display={quoteattr(target.tab)}/>'
            for target in links) + "</hyperlinks>"
    selected = ' tabSelected="1"' if sheet.code == "shHome" else ""
    return (f'{XML_HEADER}<worksheet xmlns="{MAIN}" xmlns:r="{RELATIONSHIPS}">'
            f'<sheetPr codeName="{sheet.code}"><tabColor rgb="FF{sheet.color}"/><pageSetUpPr fitToPage="1"/></sheetPr>'
            f'<dimension ref="B1:D{len(rows)}"/>'
            f'<sheetViews><sheetView showGridLines="0"{selected} zoomScaleNormal="100" workbookViewId="0">'
            f'<pane xSplit="1" ySplit="{TOP_ROWS}" topLeftCell="B{TOP_ROWS + 1}" activePane="bottomRight" state="frozen"/>'
            '<selection pane="topRight" activeCell="B1" sqref="B1"/>'
            f'<selection pane="bottomLeft" activeCell="A{TOP_ROWS + 1}" sqref="A{TOP_ROWS + 1}"/>'
            '<selection pane="bottomRight" activeCell="C6" sqref="C6"/></sheetView></sheetViews>'
            '<sheetFormatPr defaultColWidth="0" defaultRowHeight="15"/>'
            f'{COLS}<sheetData>{"".join(data)}</sheetData>'
            '<sheetProtection sheet="1" objects="1" scenarios="1"/>'
            f'{hyperlinks}'
            '<pageMargins left="0.4" right="0.4" top="0.5" bottom="0.5" header="0.3" footer="0.3"/>'
            '<pageSetup paperSize="9" fitToHeight="0" orientation="landscape"/>'
            '<drawing r:id="rIdNavigation"/></worksheet>')


# ---------------------------------------------------------- navigation bar ----
def shape(shape_id: int, name: str, x: int, y: int, w: int, h: int, fill: str | None, text: str = "",
          color: str = "", bold: bool = False, size: int = 900, link: str | None = None, tip: str = "",
          anchor: str = "ctr", inset: int = 0, spacing: int = 0) -> str:
    click = ""
    if link:
        click = f'<a:hlinkClick xmlns:r="{RELATIONSHIPS}" r:id="{link}" tooltip={quoteattr(tip)}/>'
    fill_xml = f'<a:solidFill><a:srgbClr val="{fill}"/></a:solidFill>' if fill else "<a:noFill/>"
    body = ""
    if text:
        run = (f'<a:rPr lang="en-GB" sz="{size}" b="{1 if bold else 0}"{f" spc={quoteattr(str(spacing))}" if spacing else ""}>'
               f'<a:solidFill><a:srgbClr val="{color}"/></a:solidFill>'
               f'<a:latin typeface="{FONT}"/><a:cs typeface="{FONT}"/></a:rPr>')
        body = ('<xdr:txBody><a:bodyPr vertOverflow="clip" horzOverflow="clip" wrap="none" '
                f'lIns="{inset * PX}" tIns="0" rIns="0" bIns="{2 * PX if anchor == "b" else 0}" rtlCol="0" '
                f'anchor="{anchor}"/><a:lstStyle/><a:p><a:pPr algn="l"/>'
                f'<a:r>{run}<a:t>{escape(text)}</a:t></a:r></a:p></xdr:txBody>')
    return (f'<xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="{shape_id}" name="{name}">{click}</xdr:cNvPr>'
            f'<xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{w}" cy="{h}"/></a:xfrm>'
            f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>{fill_xml}<a:ln><a:noFill/></a:ln></xdr:spPr>{body}</xdr:sp>')


def block_height(block: Block) -> int:
    """Height of one block of the bar, in pixels."""
    height = PAD_TOP
    for index, (_, items) in enumerate(block):
        height += (SECTION_GAP if index else 0) + HEAD_H + len(items) * (BTN_H + BTN_GAP) - BTN_GAP
    return height + PAD_BOTTOM


def strip_colour(rgb: str) -> str:
    red, green, blue = (int(rgb[i:i + 2], 16) for i in (0, 2, 4))
    return DARK_STRIP if 0.2126 * red + 0.7152 * green + 0.0722 * blue < 70 else rgb


def group(name: str, block: Block, current: Sheet, first_id: int, links: dict[str, str],
          row: int, y0: int) -> tuple[str, int]:
    """One block of the bar, anchored at the top of `row` (0-based), `y0` EMU down the sheet."""
    x0, children, y, next_id = BAR_X * PX, [], PAD_TOP, first_id
    for index, (title, items) in enumerate(block):
        if index:
            y += SECTION_GAP
        children.append(shape(next_id, "Nav_Section_" + re.sub(r"\W", "", title.title()), x0, y0 + y * PX,
                              BAR_W * PX, HEAD_H * PX, None, title.upper(), HEAD_COLOR, True, 700,
                              anchor="b", inset=4, spacing=60))
        next_id += 1
        y += HEAD_H
        for target in items:
            here = target.code == current.code
            children.append(shape(next_id, "Nav_" + target.tab, x0, y0 + y * PX, BAR_W * PX, BTN_H * PX,
                                  CUR_FILL if here else BTN_FILL, target.tab, CUR_TEXT if here else BTN_TEXT,
                                  here, 900, None if here else links[target.code], target.tip, inset=12))
            children.append(shape(next_id + 1, "Nav_" + target.tab + "_Mark", x0, y0 + y * PX, STRIP_W * PX,
                                  BTN_H * PX, CUR_STRIP if here else strip_colour(target.color)))
            next_id += 2
            y += BTN_H + BTN_GAP
    height = (y - BTN_GAP + PAD_BOTTOM) * PX
    xml = (f'<xdr:grpSp><xdr:nvGrpSpPr><xdr:cNvPr id="{next_id}" name="{name}" descr="Sheet navigation"/>'
           f'<xdr:cNvGrpSpPr/></xdr:nvGrpSpPr><xdr:grpSpPr><a:xfrm><a:off x="{x0}" y="{y0}"/>'
           f'<a:ext cx="{BAR_W * PX}" cy="{height}"/><a:chOff x="{x0}" y="{y0}"/>'
           f'<a:chExt cx="{BAR_W * PX}" cy="{height}"/></a:xfrm></xdr:grpSpPr>{"".join(children)}</xdr:grpSp>')
    return (f'<xdr:oneCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:colOff>{x0}</xdr:colOff>'
            f'<xdr:row>{row}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>'
            f'<xdr:ext cx="{BAR_W * PX}" cy="{height}"/>{xml}'
            '<xdr:clientData fPrintsWithSheet="0"/></xdr:oneCellAnchor>'), next_id + 1


def navigation(current: Sheet, menu: Block) -> tuple[str, str]:
    """The drawing of one sheet and its relationships: the bar's two blocks."""
    sheets = [sheet for _, items in menu for sheet in items]
    links = {sheet.code: "rId%d" % number for number, sheet in enumerate(sheets, 1)}
    top, rest = blocks(menu)
    frozen = float(ROW_PROJECT) + float(ROW_TITLE) + float(ROW_SUBTITLE)
    if block_height(top) * 0.75 > frozen:
        raise ValueError(f"the section {top[0][0]} does not fit the frozen rows: three sheets at most")
    first, next_id = group(GROUP_TOP, top, current, 2, links, 0, 0)
    second, _ = group(GROUP_REST, rest, current, next_id, links, TOP_ROWS, int(round(frozen * 12700)))
    drawing = f'{XML_HEADER}<xdr:wsDr xmlns:xdr="{DRAWING_NS}" xmlns:a="{ART_NS}">{first}{second}</xdr:wsDr>'
    relationships = XML_HEADER + f'<Relationships xmlns="{PACKAGE_RELATIONSHIPS}">' + "".join(
        f'<Relationship Id="{links[sheet.code]}" Type="{RELATIONSHIPS}/hyperlink" Target="#{sheet.name}"/>'
        for sheet in sheets if sheet.code != current.code) + "</Relationships>"
    return drawing, relationships


# ----------------------------------------------------------------- package ----
def workbook_module() -> Module:
    """ThisWorkbook: the save events that keep the saved file with the bar shown."""
    empty = document_module("ThisWorkbook", True)
    events = THIS_WORKBOOK_EVENTS.replace("\r\n", "\n").replace("\n", "\r\n")
    return Module("ThisWorkbook", empty.source + events, DOCUMENT)


def build() -> tuple[bytes, dict[str, str]]:
    menu = sections()
    sheets = [sheet for _, items in menu for sheet in items]
    commit = git(ROOT, "rev-parse", "HEAD")
    sources = sorted((ROOT / "src").glob("*.bas")) + [DEMO / "M_STATS_PROBDIST_DEMO.bas"]
    modules = [workbook_module()]
    modules += [document_module(sheet.code, False) for sheet in sheets]
    modules += [Module(path.stem, module_text(path.read_bytes())) for path in sources]

    parts: dict[str, bytes] = {}

    def put(name: str, text: str) -> None:
        parts[name] = text.encode("utf-8")

    overrides = ['<Override PartName="/xl/workbook.xml" '
                 'ContentType="application/vnd.ms-excel.sheet.macroEnabled.main+xml"/>',
                 '<Override PartName="/xl/styles.xml" '
                 'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>',
                 '<Override PartName="/xl/theme/theme1.xml" '
                 'ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>']
    tabs, relationships, names = [], [], []
    for number, sheet in enumerate(sheets, 1):
        overrides.append(f'<Override PartName="/xl/worksheets/sheet{number}.xml" ContentType='
                         '"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>')
        overrides.append(f'<Override PartName="/xl/drawings/drawing{number}.xml" '
                         'ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>')
        tabs.append(f'<sheet name={quoteattr(sheet.tab)} sheetId="{number}" r:id="rIdSheet{number}"/>')
        relationships.append(f'<Relationship Id="rIdSheet{number}" Type="{RELATIONSHIPS}/worksheet" '
                             f'Target="worksheets/sheet{number}.xml"/>')
        names.append(f"<definedName name=\"{sheet.name}\">{escape(chr(39) + sheet.tab + chr(39))}!$A$1</definedName>")
        put(f"xl/worksheets/sheet{number}.xml", worksheet(sheet, menu, commit))
        put(f"xl/worksheets/_rels/sheet{number}.xml.rels",
            f'{XML_HEADER}<Relationships xmlns="{PACKAGE_RELATIONSHIPS}"><Relationship Id="rIdNavigation" '
            f'Type="{RELATIONSHIPS}/drawing" Target="../drawings/drawing{number}.xml"/></Relationships>')
        drawing, drawing_relationships = navigation(sheet, menu)
        put(f"xl/drawings/drawing{number}.xml", drawing)
        put(f"xl/drawings/_rels/drawing{number}.xml.rels", drawing_relationships)

    put("[Content_Types].xml",
        f'{XML_HEADER}<Types xmlns="{CONTENT_TYPES}">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/>'
        + "".join(overrides) + "</Types>")
    put("_rels/.rels",
        f'{XML_HEADER}<Relationships xmlns="{PACKAGE_RELATIONSHIPS}">'
        f'<Relationship Id="rId1" Type="{RELATIONSHIPS}/officeDocument" Target="xl/workbook.xml"/>'
        f'<Relationship Id="rIdRibbon" Type="{RIBBON_TYPE}" Target="customUI/customUI14.xml"/></Relationships>')
    put("xl/workbook.xml",
        f'{XML_HEADER}<workbook xmlns="{MAIN}" xmlns:r="{RELATIONSHIPS}">'
        '<workbookPr codeName="ThisWorkbook" defaultThemeVersion="124226"/>'
        '<workbookProtection lockStructure="1"/>'
        '<bookViews><workbookView xWindow="-120" yWindow="-120" windowWidth="29040" windowHeight="17520" '
        'firstSheet="0" activeTab="0"/></bookViews>'
        f'<sheets>{"".join(tabs)}</sheets><definedNames>{"".join(sorted(names))}</definedNames>'
        '<calcPr calcId="191029"/></workbook>')
    put("xl/_rels/workbook.xml.rels",
        f'{XML_HEADER}<Relationships xmlns="{PACKAGE_RELATIONSHIPS}">{"".join(relationships)}'
        f'<Relationship Id="rIdStyles" Type="{RELATIONSHIPS}/styles" Target="styles.xml"/>'
        f'<Relationship Id="rIdTheme" Type="{RELATIONSHIPS}/theme" Target="theme/theme1.xml"/>'
        f'<Relationship Id="rIdVBA" Type="{VBA_TYPE}" Target="vbaProject.bin"/></Relationships>')
    put("customUI/customUI14.xml", ribbon(menu))
    parts["xl/styles.xml"] = (DEMO / "reference" / "styles.xml").read_bytes()
    parts["xl/theme/theme1.xml"] = (DEMO / "reference" / "theme.xml").read_bytes()
    parts["xl/vbaProject.bin"] = build_vba_project(modules, project_id(commit))

    out = io.BytesIO()
    with zipfile.ZipFile(out, "w") as archive:
        for name in ["[Content_Types].xml", "_rels/.rels"] + sorted(set(parts) - {"[Content_Types].xml", "_rels/.rels"}):
            info = zipfile.ZipInfo(name, ZIP_TIME)
            info.compress_type = zipfile.ZIP_STORED
            info.create_system = 0
            info.external_attr = 0
            archive.writestr(info, parts[name])
    inputs = sources + sorted((DEMO / "reference").glob("*.xml")) + [
        DEMO / "sheets.json", Path(__file__), ROOT / "tools/devbuild/vba_project.py",
        ROOT / "tools/devbuild/build_workbook.py", ROOT / "VERSION"]
    hashes = {str(path.relative_to(ROOT)): hashlib.sha256(path.read_bytes()).hexdigest() for path in inputs}
    return out.getvalue(), hashes


def main() -> int:
    output = ROOT / "test-results" / "demo"
    output.mkdir(parents=True, exist_ok=True)
    data, hashes = build()
    (output / WORKBOOK_NAME).write_bytes(data)
    (output / "customUI14.xml").write_text(ribbon(sections()), encoding="utf-8")
    manifest = {"commit": git(ROOT, "rev-parse", "HEAD"), "status": "UNVERIFIED IN EXCEL",
                "dirty": bool(git(ROOT, "status", "--porcelain")),
                "workbook_sha256": hashlib.sha256(data).hexdigest(), "inputs": hashes}
    (output / "manifest.json").write_text(json.dumps(manifest, indent=2) + "\n", encoding="utf-8")
    print(output)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
