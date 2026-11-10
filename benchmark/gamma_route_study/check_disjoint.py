"""Prove the #307 routing holdout shapes are new to the repository.

Every number in every tracked CSV, JSON, Markdown and VBA file is parsed as a
binary64 value (decimal text, and hi;lo pairs as their two parts). A holdout
shape A, or its Chi-square doubling 2A, must not occur anywhere. Run before
the holdout is evaluated; the design records its output.

Run: python3 check_disjoint.py   (exit 0 = disjoint)
"""
import re
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
HERE = Path(__file__).resolve().parent
NUMBER = re.compile(r"(?<![\w.])[-+]?(?:\d+\.\d*|\.\d+|\d+)(?:[eE][-+]?\d+)?#?(?![\w.])")
HEX = re.compile(r"0x[0-9a-fA-F]\.[0-9a-fA-F]+p[-+]?\d+|0x0\.[0-9a-fA-F]+p-1022")

sys.path.insert(0, str(HERE))
from holdout_design import SHAPES  # noqa: E402


def values(text):
    out = set()
    for m in NUMBER.finditer(text):
        try:
            out.add(float(m.group(0).rstrip("#")))
        except ValueError:
            pass
    for m in HEX.finditer(text):
        try:
            out.add(float.fromhex(m.group(0)))
        except ValueError:
            pass
    return out


def main():
    files = subprocess.check_output(["git", "ls-files", "*.csv", "*.json", "*.md", "*.bas", "*.py"],
                                    cwd=ROOT, text=True).split()
    own = {p.relative_to(ROOT).as_posix() for p in HERE.iterdir() if p.name.startswith(("holdout_", "HOLDOUT_", "check_disjoint"))}
    seen = {}
    for name in files:
        if name in own:
            continue
        for v in values((ROOT / name).read_text(encoding="utf-8", errors="replace")):
            seen.setdefault(v, name)
    clash = [(a, k, seen[k]) for a in SHAPES for k in (a, 2 * a) if k in seen]
    if clash:
        for a, k, where in clash:
            print(f"CLASH: shape {a!r}: {k!r} already in {where}")
        return 1
    print(f"PASS: {len(SHAPES)} holdout shapes and their doublings occur in none of "
          f"{len(files) - len(own & set(files))} tracked files")
    return 0


if __name__ == "__main__":
    sys.exit(main())
