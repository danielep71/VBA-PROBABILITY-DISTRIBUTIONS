"""
Structural guard for the repository root README.

WHY THIS EXISTS

The root README was silently replaced by an unrelated 88-line study README
and no automated check noticed. The Accuracy Gate's README step runs
`render_contract_table.py --write` with working directory `benchmark`, so
`git diff --exit-code README.md` there resolves to `benchmark/README.md`.
The root file - the public face of the project - had no guard at all.

WHY NOT A WHOLE-FILE HASH

The root README is meant to change. #28 will regenerate the assurance
figures into it, and prose is edited routinely. A content hash would fail
on every legitimate edit, and a guard that cries wolf gets deleted.

WHAT IS ACTUALLY GUARDED

Identity and structure, which wholesale replacement destroys and ordinary
editing preserves:

  * the required top-level sections are all present, in order;
  * the document is not implausibly short;
  * the document does not appear to be some OTHER document that happens to
    live at this path - the failure that actually occurred.

Editing the assurance numbers, rewording prose, or adding a section all
pass. Overwriting the file with a study README, truncating it, or dropping
a major section all fail.

Run: python3 check_root_readme.py   (exit 0 = pass, nonzero = fail)
"""
import os
import sys

from _manifest import repo_root

README_NAME = "README.md"

# Structural anchors. Section headings, deliberately NOT content: a heading
# survives any rewrite of the text beneath it, so #28 regenerating the
# assurance table cannot trip this. Order is enforced too, because a
# document that has the right headings in the wrong order is not this
# document.
REQUIRED_ANCHORS = (
    "# 📊 VBA Probability Distributions",
    "### 📐 Assurance at a glance",
    "## ✨ What this project is",
    "# 🧩 Distribution catalogue",
    "# ⚡ Quick start",
    "# 📁 Repository structure",
    "# 🔍 Scope and validation boundary",
    "# ✅ Release checklist",
    "# 📜 Citation",
    "# 📄 License",
    "# 👤 Maintainer",
)

# Well below the current length; this catches truncation and wholesale
# replacement without failing on ordinary trimming.
MIN_LINES = 900

# Markers of documents that must never occupy the root README path. The
# study README is listed because that is the substitution that actually
# happened.
FOREIGN_TITLE_MARKERS = (
    "Chi-square inverse reference study",
    "Accuracy summary",
    "independent holdout",
)


def check(text, line_count):
    """Return a list of structural problems. Empty list means the file is
    recognisably the project README."""
    problems = []

    if line_count < MIN_LINES:
        problems.append(
            f"root README is {line_count} lines, below the {MIN_LINES}-line "
            "floor: truncated or replaced by a different document")

    # A foreign document in the first heading is the replacement failure.
    first_heading = ""
    for line in text.splitlines():
        if line.startswith("#"):
            first_heading = line.strip()
            break
    for marker in FOREIGN_TITLE_MARKERS:
        if marker.lower() in first_heading.lower():
            problems.append(
                f"root README's first heading is {first_heading!r}, which "
                f"belongs to another document ({marker!r}); the project "
                "README appears to have been overwritten")

    positions = []
    for anchor in REQUIRED_ANCHORS:
        idx = text.find(anchor)
        if idx < 0:
            problems.append(f"required section missing from root README: {anchor!r}")
        else:
            positions.append((anchor, idx))

    if len(positions) == len(REQUIRED_ANCHORS):
        for (a_prev, i_prev), (a_next, i_next) in zip(positions, positions[1:]):
            if i_next < i_prev:
                problems.append(
                    f"root README sections are out of order: {a_next!r} "
                    f"appears before {a_prev!r}")
                break

    return problems


def main():
    root = repo_root()
    path = os.path.join(root, README_NAME)
    if not os.path.exists(path):
        print(f"FAIL: root {README_NAME} is missing entirely")
        return 1
    with open(path, encoding="utf-8") as f:
        text = f.read()
    problems = check(text, text.count("\n") + 1)
    if problems:
        print("FAIL: root README structural guard")
        for p in problems:
            print("  - " + p)
        return 1
    print(f"PASS: root README intact ({len(REQUIRED_ANCHORS)} sections, "
          f"{text.count(chr(10)) + 1} lines)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
