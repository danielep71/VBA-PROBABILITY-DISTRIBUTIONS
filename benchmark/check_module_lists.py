#!/usr/bin/env python3
"""Fail when the Excel evidence policy's module list differs from the tracked modules.

`.github/excel-evidence-policy.json` names, in `regression_sources`, every VBA
module the Excel regression run imports and every module whose canonical Git
bytes an exact-SHA certification record must hash. It is read by PowerShell on
the Excel runner and by the Python certification tooling, so it stays an
explicit list; the runner imports the modules in list order, and the
certification tooling sorts the paths before hashing them. This check binds that list to the tracked
`.bas` files under src/ and tests/ instead:

* every tracked module under src/ or tests/ must be listed;
* every listed module must be a tracked `.bas` file under src/ or tests/;
* a `.bas` file present under src/ or tests/ but not tracked is reported, since
  it would escape both the public-API freeze and the Excel evidence binding.

The public-API freeze (check_public_api.py) derives its module list from the
same tracked files (_tracked_modules.py), so it needs no list of its own.
Comparison is order-insensitive; duplicates are rejected.

Run: python3 check_module_lists.py   (exit 0 = lists agree, 1 = drift or error)
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from _tracked_modules import (  # noqa: E402
    REGRESSION_DIRS,
    TrackedFilesError,
    present_modules,
    regression_modules,
)

ROOT = Path(__file__).resolve().parents[1]
POLICY_PATH = ".github/excel-evidence-policy.json"
SCOPE = " or ".join(directory + "/" for directory in REGRESSION_DIRS)


class ModuleListError(ValueError):
    pass


def policy_sources(policy_file: Path) -> list[str]:
    try:
        policy = json.loads(policy_file.read_text(encoding="utf-8-sig"))
    except (OSError, UnicodeError, ValueError) as exc:
        raise ModuleListError(f"cannot read {policy_file}: {exc}") from exc
    if not isinstance(policy, dict):
        raise ModuleListError(f"{policy_file} must hold a JSON object")
    sources = policy.get("regression_sources")
    if not isinstance(sources, list) or not sources or not all(
        isinstance(item, str) and item for item in sources
    ):
        raise ModuleListError("regression_sources must be a non-empty array of paths")
    return sources


def findings(root: Path, sources: list[str]) -> list[str]:
    """Every difference between `sources` and the tracked modules, as messages."""
    try:
        tracked = regression_modules(root)
    except TrackedFilesError as exc:
        raise ModuleListError(f"cannot list the tracked modules: {exc}") from exc
    present = set(present_modules(root, REGRESSION_DIRS))
    listed = set(sources)
    result: list[str] = []
    for path in sorted({p for p in sources if sources.count(p) > 1}):
        result.append(f"regression_sources lists {path} more than once")
    for path in sorted(set(tracked) - listed):
        result.append(f"tracked module {path} is missing from regression_sources")
    for path in sorted(listed - set(tracked)):
        reason = "present but not tracked" if (root / path).is_file() else "not a tracked file"
        result.append(
            f"regression_sources lists {path}, which is not a tracked .bas module "
            f"under {SCOPE} ({reason})"
        )
    for path in sorted(present - set(tracked) - listed):
        result.append(
            f"{path} is present but not tracked: commit it and list it, or remove it"
        )
    return result


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__.split("\n", 1)[0])
    parser.add_argument("--root", type=Path, default=ROOT)
    parser.add_argument("--policy", type=Path, default=None,
                        help=f"policy file; defaults to <root>/{POLICY_PATH}")
    args = parser.parse_args(argv)
    root = args.root.resolve()
    policy_file = args.policy or (root / POLICY_PATH)
    try:
        sources = policy_sources(policy_file)
        problems = findings(root, sources)
    except ModuleListError as exc:
        print("FAIL: module lists")
        print("  - " + str(exc))
        return 1
    if problems:
        print(f"FAIL: {POLICY_PATH} regression_sources differ from the tracked modules")
        for problem in problems:
            print("  - " + problem)
        return 1
    print(f"PASS: module lists ({len(sources)} regression_sources match the tracked "
          f".bas modules under {SCOPE}; the public-API module list is derived)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
