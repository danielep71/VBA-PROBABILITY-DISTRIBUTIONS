"""Package-level demo controls; real Excel acceptance remains separate."""
import importlib.util
import io
from pathlib import Path
import unittest
import zipfile
from xml.etree import ElementTree as ET

ROOT = Path(__file__).resolve().parents[1]
SPEC = importlib.util.spec_from_file_location('demo_builder',ROOT/'tools/demobuild/build_demo.py')
assert SPEC and SPEC.loader
DEMO = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(DEMO)


class DemoShellTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.data, cls.hashes = DEMO.build()

    def test_demo_source_placement_is_narrow(self):
        import sys
        sys.path.insert(0, str(ROOT / "tools"))
        from check_vba_source import placement
        self.assertEqual(placement("examples/demo/M_STATS_PROBDIST_DEMO.bas"), [])
        self.assertTrue(placement("examples/other/M_STATS_PROBDIST_DEMO.bas"))
        self.assertTrue(placement("examples/demo/Arbitrary.bas"))

    def test_rebuild_is_byte_identical(self):
        self.assertEqual(self.data, DEMO.build()[0])

    def test_destinations_and_package_relationships(self):
        with zipfile.ZipFile(io.BytesIO(self.data)) as package:
            names = package.namelist()
            self.assertEqual(len(names), len(set(names)))
            for name in names:
                if name.endswith(('.xml','.rels')):
                    ET.fromstring(package.read(name))
            ns = {'s':DEMO.MAIN}
            workbook = ET.fromstring(package.read('xl/workbook.xml'))
            sheets = DEMO.sheet_map()
            self.assertEqual([x.attrib['name'] for x in workbook.findall('s:sheets/s:sheet',ns)],
                             [x['tab'] for x in sheets])
            targets = {x.attrib['name'] for x in workbook.findall('s:definedNames/s:definedName',ns)}
            for i,item in enumerate(sheets,1):
                sheet = ET.fromstring(package.read(f'xl/worksheets/sheet{i}.xml'))
                self.assertEqual(sheet.find('s:sheetPr',ns).attrib['codeName'],item['code'])
                self.assertEqual({x.attrib['location'] for x in sheet.findall('s:hyperlinks/s:hyperlink',ns)},targets - {'nav_shHome'})
            self.assertEqual(package.read('xl/styles.xml'), (ROOT/'examples/demo/reference/styles.xml').read_bytes())
            self.assertEqual(package.read('xl/theme/theme1.xml'), (ROOT/'examples/demo/reference/theme.xml').read_bytes())
            layout = ET.parse(ROOT/'examples/demo/reference/home-layout.xml').getroot()
            for i in range(1, len(sheets)+1):
                page = ET.fromstring(package.read(f'xl/worksheets/sheet{i}.xml'))
                for tag in ['cols', 'sheetFormatPr', 'pageMargins', 'pageSetup', 'sheetViews']:
                    self.assertEqual(ET.tostring(page.find('s:'+tag,ns)), ET.tostring(layout.find('s:'+tag,ns)))
            ui = ET.fromstring(package.read('customUI/customUI14.xml'))
            buttons = list(ui.iter('{http://schemas.microsoft.com/office/2009/07/customui}button'))
            self.assertEqual({b.attrib['tag'] for b in buttons if 'tag' in b.attrib},
                             {s['code'] for s in sheets})
            self.assertEqual(next(b for b in buttons if b.attrib['id']=='pdReset').attrib['enabled'],'false')
            source = (ROOT/'examples/demo/M_STATS_PROBDIST_DEMO.bas').read_text()
            for button in buttons:
                if 'onAction' in button.attrib:
                    self.assertIn('Public Sub '+button.attrib['onAction']+'(',source)
            self.assertTrue(package.read('xl/vbaProject.bin').startswith(bytes.fromhex('d0cf11e0a1b11ae1')))


if __name__ == '__main__':
    unittest.main()
