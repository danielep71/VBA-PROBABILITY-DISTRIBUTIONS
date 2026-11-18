"""Bind Function Wizard metadata to the supported signatures, not a copied count."""
import re
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]


class RegistrationTests(unittest.TestCase):
    def test_manifest_covers_public_signatures_and_argument_order(self):
        text = (ROOT / 'src/workbook/M_STATS_PROBDIST_REGISTER.bas').read_text()
        logical = re.sub(r' _\n\s*', ' ', text)
        records = {}
        for line in logical.splitlines():
            if not line.strip().startswith('AddRecord "'):
                continue
            strings = [s.replace('""', '"') for s in re.findall(r'"((?:[^"]|"")*)"', line)]
            name, args, description, *help_text = strings
            self.assertNotIn(name, records)
            records[name] = args.split('|') if args else []
            self.assertEqual(len(help_text), len(records[name]))
            for arg, help_ in zip(records[name], help_text, strict=True):
                self.assertTrue(help_.startswith(arg + ':'), (name, arg))
            for help_ in [description, *help_text]:
                self.assertGreater(len(help_), 10)
                self.assertLessEqual(len(help_), 255)
        expected = {}
        for line in (ROOT / 'docs/PUBLIC_API.txt').read_text().splitlines():
            if line.startswith('#') or not line.strip():
                continue
            _, _, name, declaration = line.split('\t')
            expected[name] = re.findall(r'(?:ByVal|ByRef)\s+(\w+)', declaration)
        self.assertEqual(records, expected)
