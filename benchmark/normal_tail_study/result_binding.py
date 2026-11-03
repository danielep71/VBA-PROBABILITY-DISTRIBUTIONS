"""
Result binding of the #72 normal-tail study (#79, PREREGISTRATION.md
Amendment 3).

results.json records, in its `provenance` block, the SHA-256 of every file
that produces or shapes its measurements: the generator and its convergence
record, the models, the validator, the design and the inputs. The candidate,
kernel_coefficients.json, is bound by the top-level `coefficients_sha256`,
which `holdout_exposure.candidate` repeats and holdout_record.check_record
cross-checks; it is not listed a second time here.

`stale(result)` compares those identities with the files on disk and returns
every difference: a changed, missing or unlisted file, or a missing block.
An empty list means the record was made from exactly these files. It reads
and hashes files only, so it takes well under a second; the 35 s replay
(validate_kernel.py --verify) is what shows the values themselves reproduce.

A legitimate change to any bound file makes the record stale until
`validate_kernel.py --regenerate --reason TEXT` rewrites it. The regeneration
entry lists the identities that changed. The holdout is not evaluated, and a
changed candidate or frozen preregistration text is still refused by the
holdout record.

Standard library only, so its fixtures run without mpmath.
"""
from __future__ import annotations

import hashlib
import os
from typing import Any

HERE = os.path.dirname(os.path.abspath(__file__))
BENCH = os.path.dirname(HERE)

CANDIDATE = "normal_tail_study/kernel_coefficients.json"

# Path relative to benchmark/, and its role in producing results.json.
BOUND: tuple[tuple[str, str], ...] = (
    ("normal_tail_study/derive_kernel.py", "generator: the fit and its convergence decision"),
    ("normal_tail_study/fit_convergence.json", "generator output: the fit convergence record"),
    ("normal_tail_study/kernel_model.py", "model: the binary64 candidate and public paths"),
    ("student_t_large_df_study/binary64_route_model.py", "model: today's kernel and the #34 route"),
    ("student_t_large_df_study/coefficients.json", "model input: the #34 route coefficients"),
    ("normal_tail_study/validate_kernel.py", "validator: sets, oracles, checks and predictions"),
    ("normal_tail_study/holdout_record.py", "validator: the exposure record and the modes"),
    ("normal_tail_study/result_binding.py", "validator: this binding"),
    ("normal_tail_study/PREREGISTRATION.md", "design and configuration, with its amendments"),
    ("probability_accuracy_grid.csv", "input: §7 rows and committed observations"),
    ("survival_boundary/survival_boundary_grid.csv", "input: §7 rows and committed observations"),
    ("holdout_older/holdout_older_grid.csv", "input: §7 rows and committed observations"),
    ("accuracy_contracts.csv", "input: the P4 contracts"),
    ("student_t_large_df_study/results.json", "input: the A7 route points"),
)

RULE = ("SHA-256 of every file that produces or shapes this record, by path relative to benchmark/. "
        "The candidate is bound by coefficients_sha256. A change to any of them makes the record stale "
        "until validate_kernel.py --regenerate --reason TEXT rewrites it (PREREGISTRATION.md Amendment 3).")


def file_sha256(path: str) -> str:
    with open(path, "rb") as f:
        return hashlib.sha256(f.read()).hexdigest()


def identities(root: str = BENCH) -> dict[str, Any]:
    """The provenance block for the files under root as they are now."""
    return {"rule": RULE,
            "files": {name: {"role": role, "sha256": file_sha256(os.path.join(root, name))}
                      for name, role in BOUND}}


def _current(root: str, name: str) -> str | None:
    try:
        return file_sha256(os.path.join(root, name))
    except OSError:
        return None


def stale(result: Any, root: str = BENCH) -> list[str]:
    """Every way the record's identities differ from the files under root; [] if none."""
    if not isinstance(result, dict):
        return ["the record is not a JSON object"]
    out = []
    now = _current(root, CANDIDATE)
    if result.get("coefficients_sha256") != now:
        out.append(f"{CANDIDATE} (candidate): recorded {result.get('coefficients_sha256')}, now {now}")
    block = result.get("provenance")
    files = block.get("files") if isinstance(block, dict) else None
    if not isinstance(files, dict):
        return out + ["results.json has no provenance block: it does not record which generator, model, "
                      "validator, design and inputs produced it"]
    roles = dict(BOUND)
    for name in [n for n, _ in BOUND] + sorted(n for n in files if n not in roles):
        entry = files.get(name)
        recorded = entry.get("sha256") if isinstance(entry, dict) else None
        if name not in roles:
            out.append(f"{name}: recorded but no longer bound")
            continue
        now = _current(root, name)
        if entry is None:
            out.append(f"{name} ({roles[name]}): bound but not recorded")
        elif now is None:
            out.append(f"{name} ({roles[name]}): recorded {recorded}, but the file is missing")
        elif recorded != now:
            out.append(f"{name} ({roles[name]}): recorded {recorded}, now {now}")
    return out


def remedy(problems: list[str]) -> str:
    return ("results.json no longer matches the files that produced it:\n  - " + "\n  - ".join(problems)
            + "\n  If the change is intended, run validate_kernel.py --regenerate --reason TEXT, then --verify.")
