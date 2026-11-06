"""The label guide, the label catalogue, the issue forms and Dependabot agree (#203).

docs/repo_governance/ISSUE_LABELS.md documents .github/labels.json. These tests
fail when its catalogue table, its usage tables or a stated label count drift
from the manifest, and when the issue forms and .github/dependabot.yml name
different assignees or Dependabot uses a label outside the catalogue.
"""

from __future__ import annotations

import json
import re
import subprocess
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
GUIDE = ROOT / "docs/repo_governance/ISSUE_LABELS.md"
MANIFEST = ROOT / ".github/labels.json"
FORMS = ("bug.yml", "documentation.yml", "feature.yml", "security-contact.yml")
WORDS = {"one": 1, "two": 2, "three": 3, "four": 4, "five": 5, "six": 6, "seven": 7,
         "eight": 8, "nine": 9, "ten": 10, "eleven": 11, "twelve": 12}
COUNT = r"(\d+|" + "|".join(WORDS) + r")"


def number(text: str) -> int:
    return int(text) if text.isdigit() else WORDS[text.lower()]


def labels() -> list[dict[str, str]]:
    manifest: list[dict[str, str]] = json.loads(MANIFEST.read_text(encoding="utf-8"))["labels"]
    return manifest


def guide() -> str:
    return GUIDE.read_text(encoding="utf-8")


def section(text: str, heading: str) -> str:
    """The body of the level-2 or level-3 section whose heading ends with ``heading``."""
    match = re.search(rf"^(#{{2,3}}) [^\n]*{re.escape(heading)}\n(.*?)(?=^#{{2,3}} |\Z)",
                      text, re.M | re.S)
    assert match, f"ISSUE_LABELS.md has no section ending with {heading!r}"
    level = len(match.group(1))
    if level == 2:
        # A level-2 section runs to the next level-2 heading, through its subsections.
        whole = re.search(rf"^## [^\n]*{re.escape(heading)}\n(.*?)(?=^## |\Z)", text, re.M | re.S)
        assert whole
        return whole.group(1)
    return match.group(2)


def table_names(text: str) -> list[str]:
    return re.findall(r"^\| `([^`]+)` \|", text, re.M)


def tracked_markdown() -> list[Path]:
    names = subprocess.run(["git", "-C", str(ROOT), "ls-files", "-z", "--", "*.md"],
                           capture_output=True, text=True, encoding="utf-8",
                           check=True).stdout.split("\0")
    return [ROOT / name for name in names if name and (ROOT / name).is_file()]


def yaml_values(text: str, key: str) -> list[list[str]]:
    """Every list under ``key:`` in a simple YAML file, flow or block style."""
    found: list[list[str]] = []
    lines = text.splitlines()
    for index, line in enumerate(lines):
        match = re.match(rf"^(\s*)(?:- )?{re.escape(key)}:\s*(.*)$", line)
        if not match:
            continue
        indent, rest = len(match.group(1)), match.group(2).strip()
        if rest.startswith("["):
            found.append([item.strip().strip("\"'") for item in rest.strip("[]").split(",")
                          if item.strip()])
            continue
        items = []
        for following in lines[index + 1:]:
            stripped = following.strip()
            if stripped.startswith("- ") and len(following) - len(following.lstrip()) > indent:
                items.append(stripped[2:].strip().strip("\"'"))
            elif stripped and not stripped.startswith("#"):
                break
        found.append(items)
    return found


class CatalogueTableTests(unittest.TestCase):
    def test_catalogue_table_matches_the_manifest(self) -> None:
        rows = re.findall(r"^\| `([^`]+)` \| `#([0-9A-F]{6})` \| (.+) \|$",
                          section(guide(), "Catalogue"), re.M)
        expected = [(label["name"], label["color"], label["description"]) for label in labels()]
        self.assertEqual(rows, expected)


class UsageTableTests(unittest.TestCase):
    def test_usage_tables_cover_every_label_once(self) -> None:
        names = table_names(section(guide(), "Using the labels"))
        self.assertEqual(sorted(names, key=str.lower),
                         sorted((label["name"] for label in labels()), key=str.lower))

    def test_every_repository_specific_label_has_a_rule(self) -> None:
        areas = section(guide(), "Areas")
        rows = re.findall(r"^\| `([^`]+)` \| (.+) \|$", areas, re.M)
        self.assertTrue(rows, "the Areas table is empty")
        for name, rule in rows:
            with self.subTest(label=name):
                self.assertGreater(len(rule.split()), 5, "the rule says too little")


class LabelCountTests(unittest.TestCase):
    def test_stated_counts_match_the_catalogue(self) -> None:
        text = guide()
        total = len(labels())
        specific = len(table_names(section(text, "Areas")))
        stated_total = re.findall(rf"\b{COUNT}\s+issue\s+labels\b", text, re.I)
        stated_total += re.findall(r"labels-(\d+)-", text)
        self.assertTrue(stated_total, "ISSUE_LABELS.md states no total label count")
        self.assertEqual({number(item) for item in stated_total}, {total})
        common = re.findall(rf"\b{COUNT}\s+common\s+labels\b", text, re.I)
        own = re.findall(rf"\b{COUNT}\s+repository-specific\s+labels\b", text, re.I)
        self.assertTrue(common and own, "ISSUE_LABELS.md does not state the split")
        self.assertEqual({number(item) for item in own}, {specific})
        self.assertEqual({number(item) for item in common}, {total - specific})

    def test_no_other_document_states_a_different_count(self) -> None:
        total = len(labels())
        pattern = re.compile(rf"\b{COUNT}\s+issue\s+labels\b|labels-(\d+)-", re.I)
        wrong = []
        for path in tracked_markdown():
            for prose, badge in pattern.findall(path.read_text(encoding="utf-8")):
                if number(prose or badge) != total:
                    wrong.append((path.relative_to(ROOT).as_posix(), prose or badge))
        self.assertEqual(wrong, [])


class ClosingIssuesTests(unittest.TestCase):
    def test_closing_section_requires_evidence_and_a_resolution(self) -> None:
        text = section(guide(), "Closing issues")
        for phrase in ("evidence", "Resolution", "not validated"):
            with self.subTest(phrase=phrase):
                self.assertIn(phrase, text)


class AssigneeTests(unittest.TestCase):
    def test_issue_forms_and_dependabot_name_the_same_assignee(self) -> None:
        forms = {}
        for name in FORMS:
            text = (ROOT / ".github/ISSUE_TEMPLATE" / name).read_text(encoding="utf-8")
            values = yaml_values(text, "assignees")
            self.assertEqual(len(values), 1, f"{name} must declare assignees once")
            forms[name] = values[0]
        dependabot = yaml_values((ROOT / ".github/dependabot.yml").read_text(encoding="utf-8"),
                                 "assignees")
        self.assertTrue(dependabot, "dependabot.yml assigns no one")
        everyone = list(forms.values()) + dependabot
        self.assertEqual(len({tuple(value) for value in everyone}), 1, everyone)
        self.assertEqual(len(everyone[0]), 1, "exactly one maintainer is assigned")

    def test_dependabot_labels_are_in_the_catalogue(self) -> None:
        catalogue = {label["name"] for label in labels()}
        used = yaml_values((ROOT / ".github/dependabot.yml").read_text(encoding="utf-8"), "labels")
        self.assertTrue(used)
        self.assertEqual([name for group in used for name in group if name not in catalogue], [])


if __name__ == "__main__":
    unittest.main()
