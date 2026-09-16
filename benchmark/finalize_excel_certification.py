"""Add truthful fresh-export claims to an exact-SHA Excel regression record.

This tool does not run Excel. It is deliberately interlocked: use it only
immediately after the selected grids were re-exported from the exact candidate
and Export_ExcelEnvironment refreshed benchmark/excel_environment.json.

The selected grid bytes may legitimately differ from HEAD at this point: they
are the fresh, not-yet-committed Excel output. The finalizer hashes those
working-tree bytes. The per-commit manifest-provenance guard later verifies that
the retained record matches the bytes actually committed with the manifest.
"""
import argparse
import datetime
import json
import os
import sys

from excel_certification import (CertificationError, git_bytes, git_head,
                                 load_policy, row_count_file, sha256_bytes,
                                 sha256_file, validate_record)

ENV_PATH = "benchmark/excel_environment.json"
DEFAULT_OUT = "benchmark/excel_regression_record.json"


def normalize_bitness(value):
    text = str(value).strip().lower().replace(" ", "")
    if text in ("64", "64bit", "64-bit", "x64"):
        return "64-bit"
    if text in ("32", "32bit", "32-bit", "x86"):
        return "32-bit"
    return str(value).strip()


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--record", required=True,
                    help="green regression record produced by the exact candidate")
    ap.add_argument("--main", action="store_true", help="bind the main accuracy grid")
    ap.add_argument("--holdout", action="store_true", help="bind the independent holdout grid")
    ap.add_argument("--from-fresh-export", action="store_true",
                    help="assert that the selected grids were just exported in Excel")
    ap.add_argument("--out", default=DEFAULT_OUT)
    a = ap.parse_args()
    if not (a.main or a.holdout):
        raise SystemExit("select --main and/or --holdout")
    if not a.from_fresh_export:
        raise SystemExit("refusing to add export claims without --from-fresh-export")

    root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    head = git_head(root)
    try:
        with open(a.record, encoding="utf-8-sig") as f:
            record = json.load(f)

        # The regression artifact must still bind the exact checked-out
        # candidate and its committed regression sources. At this point the
        # selected CSV may already contain fresh uncommitted Excel output, but
        # the record has no export claims yet, so evidence_ref=HEAD is correct.
        validate_record(record, root, policy_ref=head, evidence_ref=head,
                        expected_candidate_sha=head, require_pass=True,
                        log_directory=os.path.dirname(os.path.abspath(a.record)))
        policy = load_policy(root, head)

        with open(os.path.join(root, ENV_PATH), encoding="utf-8-sig") as f:
            env = json.load(f)
        expected_env = record["environment"]
        comparisons = {
            "excel_version": str(env.get("excel_version", "")).strip(),
            "excel_build": str(env.get("excel_build", "")).strip(),
            "office_bitness": normalize_bitness(env.get("office_bitness", "")),
        }
        for key, observed in comparisons.items():
            expected = (normalize_bitness(expected_env[key]) if key == "office_bitness"
                        else str(expected_env[key]).strip())
            if not observed or observed != expected:
                raise CertificationError(
                    f"{ENV_PATH} {key}={observed!r} differs from regression record {expected!r}")

        selected = []
        if a.main:
            selected.append("benchmark/probability_accuracy_grid.csv")
        if a.holdout:
            selected.append("benchmark/holdout/holdout_grid.csv")
        now = datetime.datetime.now(datetime.timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")

        source_map = {item["path"]: item["sha256"] for item in record["sources"]}
        for grid in selected:
            exporter = policy["grid_exporters"][grid]
            source_map[exporter] = sha256_bytes(git_bytes(root, head, exporter))
            grid_abs = os.path.join(root, grid)
            if not os.path.isfile(grid_abs):
                raise CertificationError(f"fresh-export target is missing: {grid}")
            record["grids"][grid] = {
                "exported": True,
                "sha256": sha256_file(grid_abs),
                "row_count": row_count_file(grid_abs),
                "exported_utc": now,
            }
        record["sources"] = [
            {"path": path, "sha256": source_map[path]} for path in sorted(source_map)
        ]

        # Do not compare a freshly exported working-tree CSV with pre-export
        # HEAD. validate_record still verifies the candidate, policy and exact
        # exporter/source blobs. The commit-time manifest guard supplies
        # evidence_ref=<commit> and therefore verifies these recorded grid
        # hashes/row counts against the bytes that were actually committed.
        for grid in selected:
            validate_record(record, root, policy_ref=head, evidence_ref=None,
                            expected_candidate_sha=head, require_pass=True,
                            require_grid=grid)

        out = os.path.join(root, a.out)
        os.makedirs(os.path.dirname(out), exist_ok=True)
        with open(out, "w", encoding="utf-8", newline="\n") as f:
            json.dump(record, f, indent=2)
            f.write("\n")
    except (OSError, ValueError, KeyError, CertificationError) as exc:
        print("FAIL: finalize Excel certification")
        print("  - " + str(exc))
        return 1

    print(f"wrote {a.out}: candidate {head}; fresh export claims: {', '.join(selected)}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
