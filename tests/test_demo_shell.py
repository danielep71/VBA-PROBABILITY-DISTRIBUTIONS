"""Package-level demo controls; real Excel acceptance remains separate."""
import importlib.util
import io
import re
import sys
import unittest
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

ROOT = Path(__file__).resolve().parents[1]
SPEC = importlib.util.spec_from_file_location("demo_builder", ROOT / "tools/demobuild/build_demo.py")
assert SPEC and SPEC.loader
DEMO = importlib.util.module_from_spec(SPEC)
sys.modules["demo_builder"] = DEMO
SPEC.loader.exec_module(DEMO)

UI = "{http://schemas.microsoft.com/office/2009/07/customui}"
XDR = "{http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing}"
ART = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
REL = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
PKG = "{http://schemas.openxmlformats.org/package/2006/relationships}"
NS = {"s": DEMO.MAIN}
MODULE = ROOT / "src/workbook/M_STATS_PROBDIST_DEMO.bas"


class DemoShellTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.data, cls.hashes = DEMO.build()
        cls.package = zipfile.ZipFile(io.BytesIO(cls.data))
        cls.menu = DEMO.sections()
        cls.sheets = DEMO.sheet_map()
        cls.source = MODULE.read_text(encoding="ascii")

    @classmethod
    def tearDownClass(cls):
        cls.package.close()

    def part(self, name):
        return ET.fromstring(self.package.read(name))

    def test_demo_source_placement_is_narrow(self):
        sys.path.insert(0, str(ROOT / "tools"))
        from check_vba_source import placement
        self.assertEqual(placement("src/workbook/M_STATS_PROBDIST_DEMO.bas"), [])
        self.assertEqual(placement("src/workbook/ThisWorkbook.cls"), [])
        self.assertEqual(placement("src/workbook/shExplorer.cls"), [])
        self.assertTrue(placement("examples/demo/M_STATS_PROBDIST_DEMO.bas"))
        self.assertTrue(placement("src/workbook/sub/M_STATS_PROBDIST_DEMO.bas"))
        self.assertTrue(placement("src/workbook/Arbitrary.bas"))
        self.assertTrue(placement("src/workbook/ThisWorkbook.bas"))
        self.assertTrue(placement("src/ThisWorkbook.cls"))
        self.assertTrue(placement("tests/shExplorer.cls"))

    def test_rebuild_is_byte_identical(self):
        self.assertEqual(self.data, DEMO.build()[0])

    def test_every_part_is_well_formed_and_named_once(self):
        names = self.package.namelist()
        self.assertEqual(len(names), len(set(names)))
        self.assertEqual(names[0], "[Content_Types].xml")
        for name in names:
            if name.endswith((".xml", ".rels")):
                ET.fromstring(self.package.read(name))
        self.assertTrue(self.package.read("xl/vbaProject.bin").startswith(bytes.fromhex("d0cf11e0a1b11ae1")))

    def test_no_dangling_compatibility_prefix(self):
        # ElementTree accepts dangling prefixes inside attribute VALUES; Excel's
        # markup-compatibility reader does not. Check each scope.
        compatibility = "{http://schemas.openxmlformats.org/markup-compatibility/2006}"
        for name in self.package.namelist():
            if not name.endswith(".xml"):
                continue
            scopes, pending = [{}], {}
            events = ET.iterparse(io.BytesIO(self.package.read(name)), events=("start-ns", "start", "end"))
            for event, node in events:
                if event == "start-ns":
                    pending[node[0]] = node[1]
                elif event == "start":
                    scope = dict(scopes[-1])
                    scope.update(pending)
                    pending.clear()
                    scopes.append(scope)
                    for attribute in ("Ignorable", "MustUnderstand"):
                        for prefix in node.get(compatibility + attribute, "").split():
                            self.assertIn(prefix, scope, name + " has undeclared " + prefix)
                else:
                    scopes.pop()

    def test_template_parts_are_written_unchanged(self):
        reference = ROOT / "examples/demo/reference"
        self.assertEqual(self.package.read("xl/styles.xml"), (reference / "styles.xml").read_bytes())
        self.assertEqual(self.package.read("xl/theme/theme1.xml"), (reference / "theme.xml").read_bytes())
        styles = self.part("xl/styles.xml")
        count = len(styles.findall("s:cellXfs/s:xf", NS))
        used = set()
        for number in range(1, len(self.sheets) + 1):
            sheet = self.part(f"xl/worksheets/sheet{number}.xml")
            used |= {int(cell.get("s")) for cell in sheet.iter("{%s}c" % DEMO.MAIN)}
            used |= {int(col.get("style")) for col in sheet.iter("{%s}col" % DEMO.MAIN) if col.get("style")}
        self.assertLess(max(used), count)

    def test_workbook_lists_the_sheets_in_menu_order(self):
        workbook = self.part("xl/workbook.xml")
        self.assertEqual([sheet.get("name") for sheet in workbook.findall("s:sheets/s:sheet", NS)],
                         [sheet.tab for sheet in self.sheets])
        self.assertEqual(self.sheets[0].code, "shHome")
        self.assertEqual(workbook.find("s:workbookProtection", NS).get("lockStructure"), "1")
        names = {name.get("name"): name.text for name in workbook.findall("s:definedNames/s:definedName", NS)}
        self.assertEqual(names, {sheet.name: f"'{sheet.tab}'!$A$1" for sheet in self.sheets})

    def test_every_sheet_has_the_template_page(self):
        for number, sheet in enumerate(self.sheets, 1):
            page = self.part(f"xl/worksheets/sheet{number}.xml")
            properties = page.find("s:sheetPr", NS)
            self.assertEqual(properties.get("codeName"), sheet.code)
            self.assertEqual(properties.find("s:tabColor", NS).get("rgb"), "FF" + sheet.color)
            view = page.find("s:sheetViews/s:sheetView", NS)
            self.assertEqual(view.get("showGridLines"), "0")
            pane = view.find("s:pane", NS)
            self.assertEqual((pane.get("xSplit"), pane.get("ySplit"), pane.get("state")), ("1", "3", "frozen"))
            self.assertIn(DEMO.COLS, self.package.read(f"xl/worksheets/sheet{number}.xml").decode("utf-8"))
            self.assertEqual(page.find("s:sheetProtection", NS).attrib,
                             {"sheet": "1", "objects": "1", "scenarios": "1"})
            self.assertEqual(page.find("s:pageSetup", NS).get("orientation"), "landscape")
            cells = {cell.get("r"): cell for cell in page.iter("{%s}c" % DEMO.MAIN)}
            self.assertEqual("".join(cells["B1"].itertext()), DEMO.PROJECT)
            self.assertEqual("".join(cells["B2"].itertext()), sheet.title)
            self.assertEqual(cells["B1"].get("s"), str(DEMO.S_PROJECT[0]))
            self.assertEqual(cells["B2"].get("s"), str(DEMO.S_TITLE[0]))

    def test_home_links_every_other_sheet(self):
        home = self.part("xl/worksheets/sheet1.xml")
        links = home.findall("s:hyperlinks/s:hyperlink", NS)
        self.assertEqual([link.get("location") for link in links], [sheet.name for sheet in self.sheets[1:]])
        cells = {cell.get("r"): "".join(cell.itertext()) for cell in home.iter("{%s}c" % DEMO.MAIN)}
        self.assertEqual([cells[link.get("ref")] for link in links], [sheet.tab for sheet in self.sheets[1:]])
        for number in range(2, len(self.sheets) + 1):
            self.assertIsNone(self.part(f"xl/worksheets/sheet{number}.xml").find("s:hyperlinks", NS))

    def test_navigation_bar_lists_every_sheet_on_every_sheet(self):
        top, rest = DEMO.blocks(self.menu)
        expected = {DEMO.GROUP_TOP: top, DEMO.GROUP_REST: rest}
        self.assertEqual([title for title, _ in rest][-1], "Start")
        for number, current in enumerate(self.sheets, 1):
            drawing = self.part(f"xl/drawings/drawing{number}.xml")
            targets = {relationship.get("Id"): relationship.get("Target") for relationship in
                       self.part(f"xl/drawings/_rels/drawing{number}.xml.rels").iter(PKG + "Relationship")}
            anchors = drawing.findall(XDR + "oneCellAnchor")
            self.assertEqual(len(anchors), 2)
            identifiers = [node.get("id") for node in drawing.iter(XDR + "cNvPr")]
            self.assertEqual(len(identifiers), len(set(identifiers)))
            for anchor in anchors:
                self.assertEqual(anchor.find(XDR + "clientData").get("fPrintsWithSheet"), "0")
                group = anchor.find(XDR + "grpSp")
                name = group.find(XDR + "nvGrpSpPr/" + XDR + "cNvPr").get("name")
                buttons = {}
                for shape in group.findall(XDR + "sp"):
                    properties = shape.find(XDR + "nvSpPr/" + XDR + "cNvPr")
                    label = properties.get("name")
                    if label.startswith("Nav_Section_") or label.endswith("_Mark"):
                        continue
                    link = properties.find(ART + "hlinkClick")
                    buttons[label[4:]] = None if link is None else targets[link.get(REL + "id")]
                wanted = {sheet.tab: (None if sheet.code == current.code else "#" + sheet.name)
                          for _, items in expected[name] for sheet in items}
                self.assertEqual(buttons, wanted, f"{current.tab}: {name}")
            row = anchors[1].find(XDR + "from/" + XDR + "row").text
            self.assertEqual((anchors[0].find(XDR + "from/" + XDR + "row").text, row), ("0", "3"))

    def test_ribbon_menu_and_callbacks_agree_with_the_module(self):
        ui = self.part("customUI/customUI14.xml")
        self.assertEqual(ui.get("onLoad"), "RibbonOnLoad")
        menu = next(node for node in ui.iter(UI + "menu") if node.get("id") == DEMO.GO_MENU)
        listed, section = [], None
        for node in menu:
            if node.tag == UI + "menuSeparator":
                section = node.get("title")
            else:
                listed.append((section, node.get("id"), node.get("label"), node.get("onAction")))
        self.assertEqual(listed, [(sheet.section, sheet.control, sheet.tab, "RibbonGoTo")
                                  for sheet in self.sheets])
        identifiers = [node.get("id") for node in ui.iter() if node.get("id")]
        self.assertEqual(len(identifiers), len(set(identifiers)))
        for node in ui.iter():
            for attribute in ("onAction", "getEnabled", "getPressed", "onLoad"):
                if node.get(attribute):
                    self.assertRegex(self.source, r"(?m)^Public Sub %s\(" % node.get(attribute))
        cases = dict(re.findall(r'Case "(pdGo\w*)"\s+RibbonSheet = "(\w+)"', self.source))
        wanted = {sheet.control: sheet.code for sheet in self.sheets}
        wanted["pdGo"] = "shHome"
        self.assertEqual(cases, wanted)
        actions = {node.get("id") for node in ui.iter(UI + "button") if node.get("onAction") == "RibbonAction"}
        self.assertEqual(actions, set(re.findall(r'Case "(pd(?!Go|Navigation)\w+)"', self.source)))
        self.assertIn('ReturnedVal = (Control.Id <> "pdResetInputs")', self.source)
        for name in (DEMO.GROUP_TOP, DEMO.GROUP_REST):
            self.assertIn(f'As String = "{name}"', self.source)

    def test_the_workbook_module_saves_with_the_bar_shown(self):
        module = DEMO.workbook_module()
        for event, call in (("Workbook_BeforeSave", "NavigationBarBeforeSave"),
                            ("Workbook_AfterSave", "NavigationBarAfterSave")):
            self.assertIn("Private Sub " + event, module.source)
            self.assertIn("M_STATS_PROBDIST_DEMO." + call, module.source)
            self.assertRegex(self.source, r"(?m)^Public Sub %s\(\)" % call)
        self.assertNotIn("\n", module.source.replace("\r\n", ""))


if __name__ == "__main__":
    unittest.main()
