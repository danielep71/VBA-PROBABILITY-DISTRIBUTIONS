"""Validate a machine-readable Excel certification record against Git evidence."""
import argparse
import json
import os
import sys

from excel_certification import CertificationError, git_head, validate_record


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--record", required=True,
                    help="path to excel-certification.json or retained excel_regression_record.json")
    ap.add_argument("--candidate-sha", default=None,
                    help="full SHA the record must certify; defaults to the record's own candidate")
    ap.add_argument("--evidence-ref", default=None,
                    help="Git ref whose source/grid bytes must still match the record; defaults to HEAD")
    ap.add_argument("--policy-ref", default=None,
                    help="Git ref from which to read .github/excel-evidence-policy.json; defaults to evidence ref")
    ap.add_argument("--require-pass", action="store_true")
    ap.add_argument("--require-grid", default=None)
    ap.add_argument("--log-directory", default=None,
                    help="when supplied, verify the retained log file and digest")
    a = ap.parse_args()

    root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    try:
        with open(a.record, encoding="utf-8-sig") as f:
            record = json.load(f)
        evidence_ref = a.evidence_ref or git_head(root)
        policy_ref = a.policy_ref or evidence_ref
        validate_record(record, root, policy_ref=policy_ref,
                        evidence_ref=evidence_ref,
                        expected_candidate_sha=a.candidate_sha,
                        require_pass=a.require_pass,
                        require_grid=a.require_grid,
                        log_directory=a.log_directory)
    except (OSError, ValueError, CertificationError) as exc:
        print("FAIL: Excel certification record")
        print("  - " + str(exc))
        return 1

    grids = [path for path, entry in record["grids"].items() if entry["exported"]]
    print("PASS: Excel certification record")
    print(f"  candidate: {record['candidate_sha']}")
    print(f"  regression: {record['harness']['passed']}/{record['harness']['assertions']} PASS")
    print("  exported grids: " + (", ".join(grids) if grids else "none (regression-only evidence)"))
    return 0


if __name__ == "__main__":
    sys.exit(main())
