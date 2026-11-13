"""The production sources as the Gamma study's replay fixtures last saw them (#338).

The replay fixtures of this study prove historical facts: at the study's
commits, the generated port was the only difference from the earlier sources.
They used to read src/ and tests/ from the working tree and undo every later
change with public_integration.restore(), so each production change, a
comment included, needed a new inverse in this study's tooling before CI
passed.

They now read FROZEN, the tree of benchmark/_frozen_sources.GAMMA_STUDY_PIN.
The working tree no longer enters these fixtures, so later production changes
leave them green, and restore() needs no new inverse.

One narrow check still follows the working tree, in
test_production_follows_study.py: the Gamma block this study generated is
still in production unchanged, unless the study declares that production has
moved on.
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[2]))
from _frozen_sources import tree  # noqa: E402

PATHS = ("src", "tests/M_STATS_PROBDIST_TEST.bas",
         ".github/excel-evidence-policy.json", ".github/excel-study-policy.json")
FROZEN = tree(PATHS)
