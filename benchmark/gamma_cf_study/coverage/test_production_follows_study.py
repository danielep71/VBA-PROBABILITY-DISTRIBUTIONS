"""The one check of the Gamma study that still follows the working tree (#338).

The replay fixtures read frozen_tree.FROZEN and no longer notice production
changes. That is right for history, but it would also hide a change to the
code this study put into production: the generated public Gamma block of
src/M_STATS_PROBDIST_SPECIALFUNCS.bas.

While PRODUCTION_FOLLOWS_STUDY is True, that block in the working tree must
hold the same code as at the pin; comments and layout may differ
(vba_code_identity). A change that deliberately moves production on, such as
the fast Gamma route of #314, sets it to False and states why in MOVED_ON.
The study's evidence then describes the pinned block only.

Run: python3 test_production_follows_study.py   (exit 0 = pass)
"""
import sys
import unittest
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE.parents[1]))
from frozen_tree import FROZEN  # noqa: E402
from public_integration import BEGIN, END, ROOT, SRC  # noqa: E402
import vba_code_identity as V  # noqa: E402

PRODUCTION_FOLLOWS_STUDY = True
MOVED_ON = ""


def block(text):
    if text.count(BEGIN) != 1 or text.count(END) != 1:
        raise ValueError("the generated public Gamma block is missing or duplicated")
    return text[text.index(BEGIN):text.index(END) + len(END)]


class ProductionFollowsStudy(unittest.TestCase):
    def test_declaration_is_coherent(self):
        self.assertEqual(PRODUCTION_FOLLOWS_STUDY, not MOVED_ON.strip(),
                         "state why production moved on, or keep following the study")

    def test_generated_block_holds_the_studied_code(self):
        if not PRODUCTION_FOLLOWS_STUDY:
            self.skipTest("production moved on: " + MOVED_ON)
        studied = block((FROZEN / SRC).read_text()).encode("ascii")
        live = block((ROOT / SRC).read_text()).encode("ascii")
        self.assertIsNone(V.first_difference(studied, live),
                          "the generated public Gamma block changed in code; if that is "
                          "deliberate, set PRODUCTION_FOLLOWS_STUDY = False and fill MOVED_ON")

    def test_a_code_change_in_the_block_is_detected(self):
        studied = block((FROZEN / SRC).read_text())
        first = next(line for line in studied.split("\n")
                     if line.strip() and not line.strip().startswith("'"))
        changed = studied.replace(first, first + "\nPublic Const PROB_FIXTURE_338 As Long = 1", 1)
        self.assertIsNotNone(V.first_difference(studied.encode("ascii"), changed.encode("ascii")))
        commented = studied.replace(BEGIN, BEGIN + "\n' an added comment", 1)
        self.assertIsNone(V.first_difference(studied.encode("ascii"), commented.encode("ascii")))


if __name__ == "__main__":
    unittest.main()
