"""Tracked VBA module inventory shared by the module-list checks.

The production modules live in src/ and the regression harness in tests/. Any
check that needs "the production modules" or "the regression sources" derives
them from the files Git tracks there, so a new module cannot escape the
public-API freeze or the Excel evidence binding by being missing from a
hand-maintained list.

Git is the source of truth on purpose: CI, the Excel runner and certification
records all see the committed tree, not the working directory. A root that is
not the top level of a Git work tree therefore fails closed instead of falling
back to a directory scan, which could not tell a committed module from a local
scratch file.
"""
from __future__ import annotations

import subprocess
from pathlib import Path

PRODUCTION_DIR = "src"
REGRESSION_DIRS = ("src", "tests")
MODULE_SUFFIX = ".bas"


class TrackedFilesError(ValueError):
    """Git could not report the tracked files of the given root."""


def _git(root: Path, *args: str) -> bytes:
    try:
        completed = subprocess.run(
            ["git", "-C", str(root), *args], capture_output=True, check=False
        )
    except OSError as exc:
        raise TrackedFilesError(f"cannot run git: {exc}") from exc
    if completed.returncode:
        detail = completed.stderr.decode("utf-8", errors="replace").strip()
        raise TrackedFilesError(
            f"{root} is not a readable Git work tree"
            + (f": {detail}" if detail else "")
        )
    return completed.stdout


def _require_top_level(root: Path) -> None:
    top = _git(root, "rev-parse", "--show-toplevel").decode("utf-8", errors="surrogateescape").strip()
    if not top or Path(top).resolve() != root.resolve():
        raise TrackedFilesError(
            f"{root} is not the top level of its Git work tree"
            + (f" (top level: {top})" if top else "")
        )


def is_module(path: str) -> bool:
    return Path(path).suffix.lower() == MODULE_SUFFIX


def tracked_modules(root: Path, directories: tuple[str, ...]) -> list[str]:
    """Tracked `.bas` files at any depth under `directories`, sorted, "/"-separated."""
    _require_top_level(root)
    raw = _git(root, "ls-files", "-z", "--", *directories)
    names = (item.decode("utf-8", errors="surrogateescape") for item in raw.split(b"\0") if item)
    return sorted(name for name in names if is_module(name))


def present_modules(root: Path, directories: tuple[str, ...]) -> list[str]:
    """`.bas` files on disk at any depth under `directories`, sorted, "/"-separated."""
    found: list[str] = []
    for directory in directories:
        base = root / directory
        if not base.is_dir():
            continue
        found.extend(
            path.relative_to(root).as_posix()
            for path in base.rglob("*")
            if path.is_file() and is_module(path.name)
        )
    return sorted(found)


def production_modules(root: Path) -> list[str]:
    """The tracked production modules that make up the public API."""
    return tracked_modules(root, (PRODUCTION_DIR,))


def regression_modules(root: Path) -> list[str]:
    """The tracked modules the Excel regression run must import and hash."""
    return tracked_modules(root, REGRESSION_DIRS)
