"""
Fail when a study binds the whole live grid instead of the rows it uses (#302).

A study script or test that hashes or byte-compares a live grid file goes
stale whenever an unrelated row is appended; #23, #289 and #327 each paid for
that with a metadata-only regeneration. Rows are bound with
_row_binding.row_digest, and a comparison with a historical export reads that
export from the commit that holds it.

The check is static and deliberately coarse. A Python file in a study
directory (any directory under benchmark/) is FLAGGED when it both

  * names a live grid: probability_accuracy_grid.csv or holdout_grid.csv, and
  * uses a whole-file primitive: hashlib, a *sha256* helper, read_bytes,
    filecmp or normalized_hash.

Every flagged file must be listed in DECLARED with the reason it is safe:
what it hashes, or that it binds rows through _row_binding. An undeclared
flagged file fails, and so does a declaration that no longer matches a flagged
file, so the list cannot rot.

    python check_study_grid_binding.py     exit 0 = pass
"""
import os
import re
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
LIVE_GRIDS = ("probability_accuracy_grid.csv", "holdout_grid.csv")
PRIMITIVE = re.compile(r"hashlib|sha256|read_bytes|filecmp|normalized_hash")

# Path relative to benchmark/ -> why the whole-file primitive does not bind a live grid.
DECLARED = {
    "normal_tail_study/result_binding.py":
        "binds the three grids by row through _row_binding.row_digest (Amendment 4); "
        "file hashes cover code, design and non-grid inputs only",
    "normal_tail_study/test_result_binding.py":
        "fixtures of that binding; proves an appended unrelated row leaves it current",
    "normal_tail_study/validate_kernel.py":
        "reads grid rows by function; hashes only the coefficient file",
    "normal_tail_study/test_kernel_model.py":
        "looks rows up in results.json by grid name; hashes only the coefficient file",
    "normal_tail_study/check_excel_wave.py":
        "reads the grids from named commits through git show; hashes only its predictions file",
    "student_t_large_df_study/model_fidelity.py":
        "reads the origin grids from named commits through git, never the live file; "
        "hashes VBA procedures, the model and its own observation file",
    "student_t_large_df_study/predict_production.py":
        "looks rows up by function and argument; hashes only the coefficient file",
}


def flagged(root=HERE):
    out = []
    for entry in sorted(os.listdir(root)):
        study = os.path.join(root, entry)
        if not os.path.isdir(study) or entry.startswith((".", "__")):
            continue
        for folder, dirs, files in os.walk(study):
            dirs[:] = sorted(d for d in dirs if not d.startswith((".", "__")))
            for name in sorted(files):
                if not name.endswith(".py"):
                    continue
                path = os.path.join(folder, name)
                with open(path, encoding="utf-8") as f:
                    text = f.read()
                if any(g in text for g in LIVE_GRIDS) and PRIMITIVE.search(text):
                    out.append(os.path.relpath(path, root).replace(os.sep, "/"))
    return out


def problems(root=HERE, declared=None):
    declared = DECLARED if declared is None else declared
    found = flagged(root)
    out = [f"{p}: names a live grid and uses a whole-file hash or comparison; bind rows with "
           "_row_binding.row_digest, or declare in check_study_grid_binding.DECLARED why it is safe"
           for p in found if p not in declared]
    out += [f"{p}: declared but no longer flagged; remove the declaration"
            for p in sorted(declared) if p not in found]
    out += [f"{p}: declaration has no reason" for p in sorted(declared)
            if p in found and not declared[p].strip()]
    return out


def main():
    bad = problems()
    for p in bad:
        print("FAIL: " + p)
    if bad:
        return 1
    print(f"PASS: no study binds a whole live grid ({len(DECLARED)} declared file(s) use a "
          "whole-file primitive for something else)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
