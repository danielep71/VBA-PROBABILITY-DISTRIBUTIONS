"""The user manual names only functions the library declares (#263)."""
from __future__ import annotations

import importlib.util
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
MANUAL = ROOT / "docs" / "user-manual" / "v1.0.0"


def load():
    spec = importlib.util.spec_from_file_location("build_manual", MANUAL / "build_manual.py")
    assert spec and spec.loader
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


class Manual(unittest.TestCase):
    def test_every_name_and_call_matches_the_public_api(self) -> None:
        self.assertEqual(load().problems(), [])

    def test_the_api_baseline_is_read(self) -> None:
        self.assertEqual(len(load().declarations()), 112)

    def test_a_wrong_argument_count_is_found(self) -> None:
        manual = load()
        self.assertEqual(list(manual.calls("=K_STATS_Gamma_Mean(3,2) and K_STATS_Gamma_Mean(...)")),
                         [("K_STATS_Gamma_Mean", 2)])
        self.assertEqual(list(manual.calls("K_STATS_F_Density(1,(2+3),4)")), [("K_STATS_F_Density", 3)])

    def test_the_chapters_are_numbered_in_order(self) -> None:
        chapters = load().chapters()
        self.assertEqual(len(chapters), 14)
        for number, path in enumerate(chapters):
            first = path.read_text(encoding="utf-8").splitlines()[0]
            expected = "# About this manual" if number == 0 else f"# {number}. "
            self.assertTrue(first.startswith(expected), f"{path.name}: {first!r}")


if __name__ == "__main__":
    unittest.main()
