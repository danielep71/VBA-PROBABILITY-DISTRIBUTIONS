#!/usr/bin/env python3
"""Build the development demo shell, without Excel or external packages.

Run from the repository: python tools/demobuild/build_demo.py
Output is ignored under test-results/demo/. This is not a certified release.
"""
from __future__ import annotations

import hashlib
import io
import json
from pathlib import Path
import sys
import zipfile
from xml.etree import ElementTree as ET
from xml.sax.saxutils import escape, quoteattr

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT / 'tools' / 'devbuild'))
from build_workbook import (  # noqa: E402
    MAIN, PACKAGE_RELATIONSHIPS, RELATIONSHIPS, XML_HEADER, ZIP_TIME,
    git, module_text, package, project_id,
)
from vba_project import Module, build_vba_project, document_module  # noqa: E402

DEMO = ROOT / 'examples' / 'demo'


def sheet_map() -> list[dict[str, str | int]]:
    data: list[dict[str, str | int]] = json.loads((DEMO / 'sheets.json').read_text())
    for field in ('tab', 'code'):
        values = [str(item[field]) for item in data]
        if len(values) != len(set(values)):
            raise ValueError(f'duplicate {field}')
    return data


def ribbon(sheets: list[dict[str, str | int]]) -> str:
    buttons = ''.join(
        f'<button id="go{item["code"]}" label={quoteattr(str(item["title"]))} '
        f'tag="{item["code"]}" onAction="DemoGoTo"/>' for item in sheets)
    return (XML_HEADER + '<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui">'
            '<ribbon><tabs><tab id="pdDemo" label="Probability Distributions">'
            '<group id="pdNavigation" label="Navigate"><menu id="pdGo" label="Go to">'
            + buttons + '</menu></group><group id="pdActions" label="Demo">'
            '<button id="pdRecalculate" label="Recalculate" onAction="DemoRecalculate"/>'
            '<button id="pdReset" label="Reset inputs" enabled="false" '
            'screentip="Available when the thematic demos provide editable inputs"/>'
            '<button id="pdHelp" label="Help" tag="shHome" onAction="DemoGoTo"/>'
            '</group></tab></tabs></ribbon></customUI>')


def styles() -> str:
    fonts = ''.join(f'<font><sz val="{size}"/><color rgb="{color}"/>'
                    '<name val="Aptos"/>' + ('<b/>' if bold else '') + '</font>'
                    for size, color, bold in [(11,'FF253858',False), (18,'FFFFFFFF',True),
                                              (11,'FFFFFFFF',True), (14,'FF265A99',True)])
    fills = '<fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>'
    fills += ''.join(f'<fill><patternFill patternType="solid"><fgColor rgb="{color}"/>'
                     '<bgColor indexed="64"/></patternFill></fill>'
                     for color in ['FF265A99','FFEAF1F8','FF00838F','FFFFF2CC'])
    xfs = ''.join(f'<xf numFmtId="0" fontId="{font}" fillId="{fill}" borderId="0" xfId="0" '
                  'applyAlignment="1"><alignment vertical="center" wrapText="1"/></xf>'
                  for font, fill in [(0,0),(1,2),(2,2),(0,3),(3,0),(2,4),(0,5)])
    return (XML_HEADER + f'<styleSheet xmlns="{MAIN}"><fonts count="4">{fonts}</fonts>'
            f'<fills count="6">{fills}</fills><borders count="1"><border/></borders>'
            '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
            f'<cellXfs count="7">{xfs}</cellXfs><cellStyles count="1">'
            '<cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>')


def worksheet(item: dict[str, str | int], sheets: list[dict[str, str | int]], commit: str) -> str:
    cells: dict[int, list[str]] = {}
    def put(row: int, col: str, value: str, style: int = 0) -> None:
        cells.setdefault(row, []).append(f'<c r="{col}{row}" s="{style}" t="inlineStr">'
                                         f'<is><t>{escape(value)}</t></is></c>')
    put(1,'C','Probability Distributions',1)
    put(2,'C',str(item['title']),4)
    for row, entry in enumerate(sheets, 5):
        put(row,'A',str(entry['title']),5 if entry == item else 3)
    put(4,'C','Start here' if item['code']=='shHome' else 'Worksheet under development',2)
    if item['code']=='shHome':
        lines = [
            ('Purpose','Explore probability distributions in Excel. Demonstrations are educational examples, not numerical certification.'),
            ('Installation','Open in desktop Excel for Windows. Enable macros only after reviewing the source and trusting this workbook.'),
            ('Navigate','Use the left navigation, worksheet tabs or the Probability Distributions Ribbon. Help returns here.'),
            ('Macros disabled','Navigation links and tabs remain usable. VBA functions and Ribbon actions require macros.'),
            ('Inputs','Yellow cells will hold editable example inputs. Blue sections identify formulas and results; references are labelled separately.'),
            ('Reset','Reset is disabled in this shell. Each thematic demo will define its inputs and safe reset behavior.'),
            ('Availability','The eight thematic worksheets are reserved for the linked issues; no calculations are delivered in this shell.'),
            ('Workbook version','v1.0.1 development shell; not a published release.'),
            ('Library version',(ROOT/'VERSION').read_text().strip()),
            ('Source commit',commit),
            ('Validation','Excel open/save/reopen, Ribbon and navigation acceptance is pending. See examples/demo/README.md.'),
        ]
    else:
        lines = [('Scope',str(item['title'])),('Implementation',f'Content tracked in issue #{item["issue"]}.'),
                 ('Inputs','Reserved for editable example inputs.'),('Results','Reserved for visible formulas and charts.'),
                 ('References','Reserved for independently sourced reference values.'),
                 ('Navigation','Use the left navigation or Help on the Ribbon to return to Start Here.')]
    for row,(label,value) in enumerate(lines,6):
        put(row,'C',label,3)
        put(row,'D',value,6 if label=='Inputs' else 0)
    rows = ''.join(f'<row r="{r}" ht="{42 if r>3 else 34}" customHeight="1">{"".join(cells[r])}</row>' for r in sorted(cells))
    links = ''.join(f'<hyperlink ref="A{r}" location="nav_{entry["code"]}" display={quoteattr(str(entry["title"]))}/>'
                    for r,entry in enumerate(sheets,5))
    return (XML_HEADER+f'<worksheet xmlns="{MAIN}"><sheetPr codeName="{item["code"]}"><pageSetUpPr fitToPage="1"/></sheetPr>'
            '<sheetViews><sheetView showGridLines="0" workbookViewId="0"><pane xSplit="2" ySplit="3" '
            'topLeftCell="C4" activePane="bottomRight" state="frozen"/></sheetView></sheetViews>'
            '<sheetFormatPr defaultRowHeight="24"/><cols><col min="1" max="1" width="29" customWidth="1"/>'
            '<col min="2" max="2" width="4" customWidth="1"/><col min="3" max="3" width="26" customWidth="1"/>'
            '<col min="4" max="4" width="88" customWidth="1"/></cols>'
            f'<sheetData>{rows}</sheetData><sheetProtection sheet="1" objects="1" scenarios="1" '
            'selectLockedCells="0" selectUnlockedCells="0"/>'
            '<mergeCells count="3"><mergeCell ref="C1:D1"/><mergeCell ref="C2:D2"/><mergeCell ref="C4:D4"/></mergeCells>'
            f'<hyperlinks>{links}</hyperlinks><pageMargins left="0.25" right="0.25" top="0.4" bottom="0.4" header="0.2" footer="0.2"/>'
            '<pageSetup orientation="landscape" paperSize="9" fitToWidth="1" fitToHeight="1"/></worksheet>')


def build() -> tuple[bytes, dict[str, str]]:
    sheets = sheet_map()
    commit = git(ROOT,'rev-parse','HEAD')
    paths = sorted((ROOT/'src').glob('*.bas')) + [DEMO/'M_STATS_PROBDIST_DEMO.bas']
    modules = [document_module('ThisWorkbook',True)]
    modules += [document_module(str(s['code']),False) for s in sheets]
    modules += [Module(p.stem,module_text(p.read_bytes())) for p in paths]
    project = build_vba_project(modules,project_id(commit))
    with zipfile.ZipFile(io.BytesIO(package({},project))) as source:
        parts = {name: source.read(name) for name in source.namelist()}
    def setpart(name: str, text: str) -> None:
        parts[name] = text.encode('utf-8')
    names = ''.join(f'<definedName name="nav_{s["code"]}">{escape(chr(39)+str(s["tab"])+chr(39)+"!$C$1")}</definedName>' for s in sheets)
    tabs = ''.join(f'<sheet name={quoteattr(str(s["tab"]))} sheetId="{i}" r:id="sheet{i}"/>' for i,s in enumerate(sheets,1))
    setpart('xl/workbook.xml',XML_HEADER+f'<workbook xmlns="{MAIN}" xmlns:r="{RELATIONSHIPS}"><workbookPr codeName="ThisWorkbook"/>'
            f'<bookViews><workbookView/></bookViews><sheets>{tabs}</sheets><definedNames>{names}</definedNames></workbook>')
    rels = ''.join(f'<Relationship Id="sheet{i}" Type="{RELATIONSHIPS}/worksheet" Target="worksheets/sheet{i}.xml"/>' for i in range(1,len(sheets)+1))
    setpart('xl/_rels/workbook.xml.rels',XML_HEADER+f'<Relationships xmlns="{PACKAGE_RELATIONSHIPS}">{rels}'
            f'<Relationship Id="styles" Type="{RELATIONSHIPS}/styles" Target="styles.xml"/>'
            '<Relationship Id="vba" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="vbaProject.bin"/></Relationships>')
    ct = ET.fromstring(parts['[Content_Types].xml'])
    for i in range(2,len(sheets)+1):
        ET.SubElement(ct,'{http://schemas.openxmlformats.org/package/2006/content-types}Override',PartName=f'/xl/worksheets/sheet{i}.xml',ContentType='application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml')
    parts['[Content_Types].xml'] = ET.tostring(ct,encoding='utf-8',xml_declaration=True)
    rel = ET.fromstring(parts['_rels/.rels'])
    ET.SubElement(rel,f'{{{PACKAGE_RELATIONSHIPS}}}Relationship',Id='ribbon',Type='http://schemas.microsoft.com/office/2007/relationships/ui/extensibility',Target='customUI/customUI14.xml')
    parts['_rels/.rels'] = ET.tostring(rel,encoding='utf-8',xml_declaration=True)
    setpart('customUI/customUI14.xml',ribbon(sheets))
    setpart('xl/styles.xml',styles())
    for i,s in enumerate(sheets,1):
        setpart(f'xl/worksheets/sheet{i}.xml',worksheet(s,sheets,commit))
    out = io.BytesIO()
    with zipfile.ZipFile(out,'w') as archive:
        for name,data in sorted(parts.items()):
            info = zipfile.ZipInfo(name,ZIP_TIME)
            info.create_system = 0
            archive.writestr(info,data)
    inputs = paths + [DEMO/'sheets.json', Path(__file__), ROOT/'tools/devbuild/vba_project.py', ROOT/'tools/devbuild/build_workbook.py', ROOT/'VERSION']
    hashes = {str(p.relative_to(ROOT)):hashlib.sha256(p.read_bytes()).hexdigest() for p in inputs}
    return out.getvalue(), hashes


if __name__ == '__main__':
    output = ROOT/'test-results/demo'
    output.mkdir(parents=True,exist_ok=True)
    data, hashes = build()
    (output/'Probability-Distributions-Demo.xlsm').write_bytes(data)
    (output/'customUI14.xml').write_text(ribbon(sheet_map()),encoding='utf-8')
    manifest = {'commit':git(ROOT,'rev-parse','HEAD'), 'status':'UNVERIFIED IN EXCEL',
                'dirty':bool(git(ROOT,'status','--porcelain')),
                'workbook_sha256':hashlib.sha256(data).hexdigest(),'inputs':hashes}
    (output/'manifest.json').write_text(json.dumps(manifest,indent=2)+'\n',encoding='utf-8')
    print(output)
