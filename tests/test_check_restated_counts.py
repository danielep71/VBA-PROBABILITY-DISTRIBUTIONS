"""Fixtures for tools/check_restated_counts.py (#328)."""

from __future__ import annotations

import json
import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "tools"))

import check_restated_counts as C  # noqa: E402

TABLE = "\n".join([C.TABLE_BEGIN, "| Change | Evidence |", "| --- | --- |",
                   "| docs | static checks |", C.TABLE_END])


def build(root: Path, docs: dict[str, str], tables: tuple[str, str] = (TABLE, TABLE)) -> None:
    (root / ".github").mkdir()
    (root / ".github/excel-evidence-policy.json").write_text(
        json.dumps({"expected_assertions": 5987}), encoding="utf-8")
    (root / ".github/excel-study-policy.json").write_text(
        json.dumps({"expected_assertions": 8901}), encoding="utf-8")
    (root / "benchmark/holdout").mkdir(parents=True)
    (root / "benchmark/probability_accuracy_grid.csv").write_text(
        "h\n" + "r\n" * 2151, encoding="utf-8")
    (root / "benchmark/holdout/holdout_grid.csv").write_text(
        "h\n" + "r\n" * 659, encoding="utf-8")
    (root / "docs").mkdir()
    for name, text in docs.items():
        (root / "docs" / name).write_text(text, encoding="utf-8")
    (root / "CONTRIBUTING.md").write_text("# C\n\n" + tables[0] + "\n", encoding="utf-8")
    (root / "docs/EXCEL_CERTIFICATION.md").write_text("# E\n\n" + tables[1] + "\n",
                                                      encoding="utf-8")


class RestatedCounts(unittest.TestCase):
    def findings(self, docs: dict[str, str], tables: tuple[str, str] = (TABLE, TABLE)) -> list[str]:
        with tempfile.TemporaryDirectory() as d:
            build(Path(d), docs, tables)
            return C.findings(Path(d))

    def test_the_repository_passes(self) -> None:
        self.assertEqual(C.findings(ROOT), [])

    def test_clean_documents_pass(self) -> None:
        self.assertEqual(self.findings({"A.md": "See the README assurance block.\n"}), [])

    def test_current_counts_are_refused(self) -> None:
        for text, label in (("passed 5,987 assertions", "regression assertions"),
                            ("passed 5987 assertions", "regression assertions"),
                            ("held to 8,901", "study assertions"),
                            ("the 2,151 rows", "main-grid rows"),
                            ("holdout of 659 rows", "holdout rows"),
                            ("passed `5,987` assertions", "regression assertions"),
                            ("contains `659` rows", "holdout rows")):
            found = self.findings({"A.md": text + "\n"})
            self.assertTrue(any(label in f for f in found), (text, found))

    def test_stale_copies_are_refused_by_phrase(self) -> None:
        for text in ("currently 955", "lists 8 open blockers", "every blocker (11 open)",
                     "the 2,139-row main grid", "the `2,139-row` grid"):
            self.assertTrue(self.findings({"A.md": text + "\n"}), text)

    def test_references_and_other_numbers_pass(self) -> None:
        text = ("Issue #659, commit `5987abc`, commit `abc5987`, threshold 5E-15, "
                "run 42659870, 1.659, 659.5, 5,987,000 and 26590, currently 5E-15.\n")
        self.assertEqual(self.findings({"A.md": text}), [])

    def test_research_records_are_out_of_scope(self) -> None:
        with tempfile.TemporaryDirectory() as d:
            build(Path(d), {"A.md": "clean\n"})
            (Path(d) / "docs/research").mkdir()
            (Path(d) / "docs/research/R.md").write_text("passed 5,987\n", encoding="utf-8")
            self.assertEqual(C.findings(Path(d)), [])

    def test_tables_must_agree_and_exist(self) -> None:
        other = TABLE.replace("static checks", "an Excel wave")
        self.assertTrue(any("differ" in f for f in self.findings({}, (TABLE, other))))
        self.assertTrue(any("markers missing" in f
                            for f in self.findings({}, (TABLE, "no table"))))


if __name__ == "__main__":
    unittest.main()
