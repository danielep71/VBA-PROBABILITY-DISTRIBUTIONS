"""
The documented re-reference command of the tchi_large_df grid works on the
committed grid (#214 added the #210 seam rows, which the generator did not
produce, and the merge then refused the whole grid).

  1. `generate_tchi_large_df.py --merge-references` (report only) exits 0 and
     reports that no reference would change: the generator's row set and
     references agree with the committed grid, seam rows included.
  2. The report leaves the committed grid byte-for-byte unchanged.

Run: python3 test_tchi_merge.py   (exit 0 = pass)
"""
import hashlib
import os
import subprocess
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
GRID = os.path.join(HERE, "tchi_large_df_grid.csv")


def digest():
    with open(GRID, "rb") as f:
        return hashlib.sha256(f.read()).hexdigest()


def main():
    before = digest()
    run = subprocess.run([sys.executable, os.path.join(HERE, "generate_tchi_large_df.py"),
                          "--merge-references"], capture_output=True, text=True, cwd=HERE)
    fails = []
    if run.returncode != 0:
        tail = (run.stderr or run.stdout).strip().splitlines()[-1:]
        fails.append(f"--merge-references exited {run.returncode}: {tail}")
    elif "0 reference(s) would change, 0 held" not in run.stdout:
        fails.append("--merge-references reports reference changes on the committed grid")
    if digest() != before:
        fails.append("the report-only merge changed the committed grid")
    if fails:
        print("FAIL: tchi_large_df re-reference merge")
        for f in fails:
            print("  - " + f)
        return 1
    print("PASS: tchi_large_df --merge-references covers the whole committed grid, seam rows "
          "included; no reference would change and the grid is untouched")
    return 0


if __name__ == "__main__":
    sys.exit(main())
