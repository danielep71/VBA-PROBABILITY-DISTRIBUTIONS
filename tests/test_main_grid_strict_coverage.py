"""The committed main grid is fully claimed, and stays so only with its contracts (#22).

The coverage check runs in strict mode since the transition fingerprint was
deleted. These tests read the committed grid and registry: every main-grid
row has a disposition, no fingerprint is left, and removing any one of the
contracts that closed #22 leaves rows without a disposition, which strict
mode refuses.
"""
from __future__ import annotations

import csv
import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
BENCH = ROOT / "benchmark"
sys.path.insert(0, str(BENCH))
import check_grid_coverage as C  # noqa: E402

GRID = BENCH / "probability_accuracy_grid.csv"
CONTRACTS = BENCH / "accuracy_contracts.csv"
GOVERNING = (
    "F_InverseCumulative.envelope_domain.quantile_rel",
    "ChiSquare_InverseCumulative.envelope_domain.quantile_rel",
    "StudentT_InverseCumulative.exact_median.output_abs",
)


def rows(path: Path) -> list[dict[str, str]]:
    with open(path, encoding="utf-8-sig", newline="") as handle:
        return list(csv.DictReader(handle))


class StrictCoverage(unittest.TestCase):
    def test_no_transition_fingerprint_is_committed(self) -> None:
        self.assertFalse((BENCH / "coverage_debt_v1_0_0.json").exists())

    def test_every_main_grid_row_has_a_disposition(self) -> None:
        result = C.evaluate_coverage(rows(GRID), rows(CONTRACTS), [])
        self.assertEqual(C.validate_strict(result), [])

    def test_removing_a_governing_contract_fails_strict_mode(self) -> None:
        contracts = rows(CONTRACTS)
        self.assertTrue(set(GOVERNING) <= {c["contract_id"] for c in contracts})
        for contract_id in GOVERNING:
            with self.subTest(contract=contract_id):
                kept = [c for c in contracts if c["contract_id"] != contract_id]
                result = C.evaluate_coverage(rows(GRID), kept, [])
                self.assertTrue(C.validate_strict(result),
                                f"strict mode accepted the grid without {contract_id}")


if __name__ == "__main__":
    unittest.main()
