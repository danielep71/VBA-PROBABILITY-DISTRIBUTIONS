"""
Write benchmark/observation_manifest.json, binding the current committed
observations to the source that produced them.

Run this AT EXPORT TIME - immediately after re-exporting observations from the
current source in Excel - and commit the manifest together with the grid. The
strict accuracy gate then refuses to certify the summary if the checked-out
.bas files no longer match the recorded hashes (i.e. source drifted from the
evidence).

Environment fields (Excel version/build, Office bitness) come from the machine
that ran the export; pass them so the manifest records the full provenance:

    python write_manifest.py --excel-version 2408 --excel-build 16.0.17928.20216 \
                             --office-bitness 64 --commit-sha <sha>

Any omitted field is recorded as "unrecorded" rather than guessed.
"""
import argparse
import datetime
import json
import os

from _manifest import (build_manifest, repo_root, MANIFEST_NAME)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--excel-version", default="unrecorded")
    ap.add_argument("--excel-build", default="unrecorded")
    ap.add_argument("--office-bitness", default="unrecorded")
    ap.add_argument("--commit-sha", default="unrecorded")
    ap.add_argument("--notes", default="")
    a = ap.parse_args()

    root = repo_root()
    grid = os.path.join(root, "benchmark", "probability_accuracy_grid.csv")
    contracts = os.path.join(root, "benchmark", "accuracy_contracts.csv")
    now = datetime.datetime.now(datetime.timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")

    manifest = build_manifest(
        root, grid, contracts, generated_utc=now, commit_sha=a.commit_sha,
        environment={"excel_version": a.excel_version, "excel_build": a.excel_build,
                     "office_bitness": a.office_bitness},
        notes=a.notes)

    out = os.path.join(root, "benchmark", MANIFEST_NAME)
    with open(out, "w", encoding="utf-8", newline="\n") as f:
        json.dump(manifest, f, indent=2, ensure_ascii=False)
        f.write("\n")
    n = len(manifest["source_binding"])
    print(f"wrote {out}: {n} .bas modules bound, schema {manifest['grid_schema_version']}")


if __name__ == "__main__":
    main()
