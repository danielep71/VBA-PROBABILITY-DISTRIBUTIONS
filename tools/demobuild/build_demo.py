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
            '<group id="pdNavigation" label="Navigate"><menu id="pdGo" label="Go to" size="large" imageMso="GoTo">'
            + buttons + '</menu></group><group id="pdActions" label="Demo">'
            '<button id="pdRecalculate" label="Recalculate" size="large" imageMso="CalculateNow" onAction="DemoRecalculate"/>'
            '<button id="pdReset" label="Reset inputs" enabled="false" '
            'screentip="Available when the thematic demos provide editable inputs"/>'
            '<button id="pdHelp" label="Help" imageMso="Help" tag="shHome" onAction="DemoGoTo"/>'
            '</group></tab></tabs></ribbon></customUI>')


def styles() -> str:
    return (DEMO/'reference/styles.xml').read_text(encoding='utf-8')


def worksheet(item: dict[str, str | int], sheets: list[dict[str, str | int]], commit: str) -> str:
    ns = {'s': MAIN}
    sheet = ET.parse(DEMO/'reference/home-layout.xml').getroot()
    prop = sheet.find('s:sheetPr',ns)
    assert prop is not None
    prop.set('codeName',str(item['code']))
    view = sheet.find('s:sheetViews/s:sheetView',ns)
    assert view is not None
    view.set('tabSelected', '1' if item['code']=='shHome' else '0')
    data = sheet.find('s:sheetData',ns)
    assert data is not None
    values = {
        'B1':'Probability Distributions', 'B2':str(item['title']),
        'C3':'Explore the library: select a worksheet on the left or use Go to on the Probability Distributions Ribbon.',
        'C5':'About this workbook',
        'C6':'Purpose', 'D6':'Explore distributions, probabilities and quantiles with the K_STATS_ library in Excel.',
        'C7':'Development version', 'D7':'This workbook provides the navigation shell. The thematic examples are under development.',
        'C8':'Library', 'D8':'Version '+(ROOT/'VERSION').read_text().strip(),
        'C10':'Getting started',
        'C11':'1. Choose a worksheet', 'D11':'Use the navigation bar or the sheet links below. The selected worksheet is highlighted on the left.',
        'C12':'2. Enable macros', 'D12':'Ribbon commands and library functions require macros. Worksheet links remain available without macros.',
        'C13':'3. Read the example', 'D13':'Thematic examples will distinguish inputs, formulas, results and independent references.',
        'C14':'4. Recalculate', 'D14':'Recalculate on the Ribbon updates this workbook only.',
        'C15':'Reset inputs', 'D15':'Disabled until the thematic examples define their editable inputs.',
        'C16':'Build', 'D16':'v1.0.1 development shell. Source: '+commit,
        'C17':'Validation', 'D17':'Educational workbook. Desktop Excel acceptance is pending; numerical certification is maintained separately.',
        'C19':'Worksheets',
    }
    if item['code'] != 'shHome':
        values.update({'C5':str(item['title']), 'D6':'Thematic worksheet: '+str(item['title'])+'.',
                       'D7':'Content is being prepared under issue #'+str(item['issue'])+'.'})
    for r,entry in enumerate(sheets[1:],20):
        values['C'+str(r)] = str(entry['title'])
        values['D'+str(r)] = 'Example under development — issue #'+str(entry['issue'])+'.'
    for row in list(data):
        if int(row.get('r','0'))>27:
            data.remove(row)
            continue
        for c in row:
            value=values.get(c.get('r',''))
            if value is not None:
                c.set('t','inlineStr')
                ET.SubElement(ET.SubElement(c,'{'+MAIN+'}is'),'{'+MAIN+'}t').text=value
    dimension=sheet.find('s:dimension',ns)
    assert dimension is not None
    dimension.set('ref','A1:S27')
    links=ET.Element('{'+MAIN+'}hyperlinks')
    for r,entry in enumerate(sheets[1:],20):
        ET.SubElement(links,'{'+MAIN+'}hyperlink',ref='C'+str(r),location='nav_'+str(entry['code']))
    # Keep the template schema ordering: hyperlinks precede page margins.
    margins=sheet.find('s:pageMargins',ns)
    assert margins is not None
    sheet.insert(list(sheet).index(margins),links)
    ET.SubElement(sheet,'{'+MAIN+'}drawing',{'{'+RELATIONSHIPS+'}id':'demoNavigation'})
    return ET.tostring(sheet,encoding='unicode')


def navigation(item: dict[str, str | int], sheets: list[dict[str, str | int]]) -> str:
    from copy import deepcopy
    xdr='http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing'
    art='http://schemas.openxmlformats.org/drawingml/2006/main'
    ns={'x':xdr,'a':art}
    source=ET.parse(DEMO/'reference/navigation.xml').getroot()
    button=next(sp for sp in source.findall('.//x:sp',ns)
                if (pr := sp.find('x:nvSpPr/x:cNvPr',ns)) is not None and pr.get('name')=='Nav_Params')
    drawing=ET.Element('{'+xdr+'}wsDr')
    labels=['Start Here','01  Explorer','02  Probabilities','03  Parameters','04  Comparison',
            '05  Extreme tails','06  Discrete','07  Risk','08  Diagnostics']
    layout = ET.parse(DEMO/'reference/home-layout.xml').getroot()
    heights = {int(row.get('r','0'))-1: float(row.get('ht','15'))*12700 for row in layout.findall('{'+MAIN+'}sheetData/{'+MAIN+'}row')}
    for i,entry in enumerate(sheets):
        anchor=ET.SubElement(drawing,'{'+xdr+'}oneCellAnchor')
        pos=ET.SubElement(anchor,'{'+xdr+'}from')
        row_index = 3
        offset = 127000 + i * 300000
        while offset >= heights.get(row_index,190500):
            offset -= int(heights.get(row_index,190500))
            row_index += 1
        for tag,val in [('col',0),('colOff',76200),('row',row_index),('rowOff',offset)]:
            ET.SubElement(pos,'{'+xdr+'}'+tag).text=str(val)
        ET.SubElement(anchor,'{'+xdr+'}ext',cx='1343025',cy='254000')
        sp=deepcopy(button)
        props=sp.find('x:nvSpPr/x:cNvPr',ns)
        assert props is not None
        props.set('id',str(i+2));props.set('name','Nav_'+str(entry['code']))
        link=props.find('a:hlinkClick',ns)
        assert link is not None
        link.set('{'+RELATIONSHIPS+'}id','nav'+str(i));link.set('tooltip',str(entry['title']))
        text=sp.find('.//a:t',ns)
        assert text is not None
        text.text=labels[i]
        fill=sp.find('x:spPr/a:solidFill/a:srgbClr',ns)
        assert fill is not None
        fill.set('val','7F6000' if entry==item else '262626')
        anchor.append(sp)
        ET.SubElement(anchor,'{'+xdr+'}clientData')
    return ET.tostring(drawing,encoding='unicode')


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
            f'<Relationship Id="theme" Type="{RELATIONSHIPS}/theme" Target="theme/theme1.xml"/>'
            '<Relationship Id="vba" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="vbaProject.bin"/></Relationships>')
    ct = ET.fromstring(parts['[Content_Types].xml'])
    ET.SubElement(ct,'{http://schemas.openxmlformats.org/package/2006/content-types}Override',PartName='/xl/theme/theme1.xml',ContentType='application/vnd.openxmlformats-officedocument.theme+xml')
    for i in range(1,len(sheets)+1):
        ET.SubElement(ct,'{http://schemas.openxmlformats.org/package/2006/content-types}Override',PartName=f'/xl/drawings/drawing{i}.xml',ContentType='application/vnd.openxmlformats-officedocument.drawing+xml')
    for i in range(2,len(sheets)+1):
        ET.SubElement(ct,'{http://schemas.openxmlformats.org/package/2006/content-types}Override',PartName=f'/xl/worksheets/sheet{i}.xml',ContentType='application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml')
    parts['[Content_Types].xml'] = ET.tostring(ct,encoding='utf-8',xml_declaration=True)
    rel = ET.fromstring(parts['_rels/.rels'])
    ET.SubElement(rel,f'{{{PACKAGE_RELATIONSHIPS}}}Relationship',Id='ribbon',Type='http://schemas.microsoft.com/office/2007/relationships/ui/extensibility',Target='customUI/customUI14.xml')
    parts['_rels/.rels'] = ET.tostring(rel,encoding='utf-8',xml_declaration=True)
    setpart('customUI/customUI14.xml',ribbon(sheets))
    setpart('xl/styles.xml',styles())
    parts['xl/theme/theme1.xml']=(DEMO/'reference/theme.xml').read_bytes()
    for i,s in enumerate(sheets,1):
        setpart(f'xl/worksheets/sheet{i}.xml',worksheet(s,sheets,commit))
        setpart(f'xl/drawings/drawing{i}.xml',navigation(s,sheets))
        setpart(f'xl/worksheets/_rels/sheet{i}.xml.rels',f'<Relationships xmlns="{PACKAGE_RELATIONSHIPS}"><Relationship Id="demoNavigation" Type="{RELATIONSHIPS}/drawing" Target="../drawings/drawing{i}.xml"/></Relationships>')
        links=''.join(f'<Relationship Id="nav{j}" Type="{RELATIONSHIPS}/hyperlink" Target="#nav_{e["code"]}"/>' for j,e in enumerate(sheets))
        setpart(f'xl/drawings/_rels/drawing{i}.xml.rels',f'<Relationships xmlns="{PACKAGE_RELATIONSHIPS}">{links}</Relationships>')
    out = io.BytesIO()
    with zipfile.ZipFile(out,'w') as archive:
        for name,data in sorted(parts.items()):
            info = zipfile.ZipInfo(name,ZIP_TIME)
            info.create_system = 0
            archive.writestr(info,data)
    inputs = paths + sorted((DEMO/'reference').glob('*.xml')) + [DEMO/'sheets.json', Path(__file__), ROOT/'tools/devbuild/vba_project.py', ROOT/'tools/devbuild/build_workbook.py', ROOT/'VERSION']
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
