"""Independent reference provenance and the optional Explorer package contract."""
import importlib.util
import io
import json
import re
import sys
import unittest
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "tools/demobuild"))
import build_demo


class ExplorerTests(unittest.TestCase):
    def test_capabilities_exist_in_target_public_api(self):
        source = "\n".join(p.read_text() for p in (ROOT / "src").glob("*.bas"))
        public = set(re.findall(r"Public Function (K_STATS_\w+)\(", source))
        families = json.loads((ROOT / "examples/demo/explorer-capabilities.json").read_text())["families"]
        self.assertEqual([f["name"] for f in families], ["Normal", "Exponential", "Poisson"])
        for family in families:
            for operation in ("density", "cdf", "survival"):
                self.assertIn(family[operation], public)

    def test_reference_fixture_reproduces_independently(self):
        path = ROOT / "examples/demo/explorer_references.py"
        spec = importlib.util.spec_from_file_location("explorer_refs", path)
        module = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(module)
        self.assertEqual(module.references(), json.loads(path.with_name("explorer-references.json").read_text()))

    def test_shell_variant_does_not_acquire_explorer_callbacks(self):
        shell = zipfile.ZipFile(io.BytesIO(build_demo.build()[0]))
        explorer, inputs = build_demo.build(explorer=True)
        variant = zipfile.ZipFile(io.BytesIO(explorer))
        self.assertNotIn(b'ExplorerAction', shell.read("customUI/customUI14.xml"))
        self.assertIn(b'ExplorerAction', variant.read("customUI/customUI14.xml"))
        self.assertIn("examples/demo/M_STATS_PROBDIST_EXPLORER.bas", inputs)
        for name in shell.namelist():
            if name not in ("xl/vbaProject.bin", "customUI/customUI14.xml"):
                self.assertEqual(shell.read(name), variant.read(name), name)
