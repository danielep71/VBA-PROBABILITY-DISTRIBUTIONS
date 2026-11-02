"""Tests for the Markdown link checker (tools/check_links.py).

Each fixture is a small Markdown tree written to a temporary directory. The
live check of this repository's own Markdown is the markdown-links check of
tools/run_static_checks.py.
"""

from __future__ import annotations

import contextlib
import io
import subprocess
import sys
import tempfile
import unittest
import unittest.mock
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "tools"))

import check_links  # noqa: E402

REPOSITORY = "owner/repo"


def run_check(files: dict[str, str]) -> tuple[list[str], set[str], int]:
    with tempfile.TemporaryDirectory() as directory:
        root = Path(directory)
        for name, text in files.items():
            path = root / name
            path.parent.mkdir(parents=True, exist_ok=True)
            path.write_text(text, encoding="utf-8")
        # Not a Git checkout: the checker falls back to every *.md file.
        return check_links.check(root)


class SlugTests(unittest.TestCase):
    def test_github_slugs(self) -> None:
        slug = check_links.github_slug
        self.assertEqual(slug("Assurance at a glance"), "assurance-at-a-glance")
        self.assertEqual(slug("🧭 Maintenance policy"), "-maintenance-policy")
        # U+FE0F after an emoji is a combining mark that GitHub keeps; "①" is not a word character.
        self.assertEqual(slug("▶️ Launch sequence"), "️-launch-sequence")
        self.assertEqual(slug("① Preview the changes"), "-preview-the-changes")
        self.assertEqual(slug("Café 2026_v1"), "café-2026_v1")
        self.assertEqual(slug("Run `check.py` [now](x.md)"), "run-checkpy-now")
        self.assertEqual(slug("Rock &amp; Roll"), "rock--roll")
        self.assertEqual(slug("K_STATS_* API: v1.0.0 (draft)?"), "k_stats_-api-v100-draft")
        self.assertEqual(slug("[Unreleased]"), "unreleased")
        # GitHub slugs the rendered text: "_" emphasis delimiters go, "_" in words stays.
        self.assertEqual(slug("__API__"), "api")
        self.assertEqual(slug("__API__ and _x_"), "api-and-x")
        self.assertEqual(slug("foo_bar_ baz"), "foo_bar_-baz")
        self.assertEqual(slug("`__init__` and __init__"), "__init__-and-init")
        self.assertEqual(slug("__API__", emphasis=False), "__api__")

    def test_duplicate_and_setext_headings(self) -> None:
        _, anchors = check_links.markdown_links("# Notes\n\n## Notes\n\nTitle\n=====\n")
        self.assertEqual(anchors, {"notes", "notes-1", "title"})
        _, anchors = check_links.markdown_links("# Foo\n# Foo-1\n# Foo\n")
        self.assertEqual(anchors, {"foo", "foo-1", "foo-2"})

    def test_code_is_ignored_and_explicit_anchors_count(self) -> None:
        text = ('<a id="custom"></a>\n<a name=\'legacy\'></a>\n\n```sh\n[x](missing.md)\n'
                "# not a heading\n```\n~~~\n<a href=\"gone.md\">x</a>\n~~~\n"
                "`[y](also-missing.md)` [z](real.md#part)\n")
        links, anchors = check_links.markdown_links(text)
        self.assertEqual(links, ["real.md#part"])
        self.assertEqual(anchors, {"custom", "legacy"})


class CheckTests(unittest.TestCase):
    def test_valid_links_pass_and_external_links_are_collected(self) -> None:
        errors, external, count = run_check({
            "README.md": "[Guide](docs/guide.md#-install) · [site](https://example.com/page#top)\n"
                         "[self](#readme) <https://example.org>\n\n# README\n"
                         "[wiki](https://github.com/owner/repo/wiki/Home)\n",
            "docs/guide.md": "# 🚀 Install\n\n[Back](../README.md)\n",
        })
        self.assertEqual(errors, [])
        self.assertEqual(external, {"https://example.com/page", "https://example.org",
                                    "https://github.com/owner/repo/wiki/Home"})
        self.assertEqual(count, 6)

    def test_github_slug_rules_resolve_duplicates_and_punctuation(self) -> None:
        errors, _, _ = run_check({
            "README.md": "[a](docs/notes.md#usage) [b](docs/notes.md#usage-1) "
                         "[c](docs/notes.md#usage-2) [d](docs/notes.md#whats-new-v100)\n"
                         "[e](docs/notes.md#-maintenance-policy) [f](docs/notes.md#custom-anchor)\n"
                         "[g](docs/notes.md#unreleased)\n",
            "docs/notes.md": "# Usage\n## Usage\n### Usage\n## What's new (v1.0.0)?\n"
                             "## 🧭 Maintenance policy\n<a id=\"custom-anchor\"></a>\n"
                             "## [Unreleased]\n",
        })
        self.assertEqual(errors, [])

    def test_percent_encoded_paths_and_html_attributes_are_checked(self) -> None:
        errors, _, _ = run_check({
            "README.md": "[ok](docs/my%20notes.md#notes) <img src=\"docs/logo.svg\">\n"
                         "<a href=\"docs/my%20notes.md#nowhere\">x</a> "
                         "<img src=\"docs/missing.svg\">\n",
            "docs/my notes.md": "# Notes\n",
            "docs/logo.svg": "<svg/>\n",
        })
        self.assertEqual(errors, ["README.md: missing anchor: docs/my%20notes.md#nowhere",
                                  "README.md: missing target: docs/missing.svg"])

    def test_link_inside_code_is_ignored(self) -> None:
        errors, _, count = run_check({
            "README.md": "```markdown\n[broken](missing.md)\n```\n\n"
                         "Write `[text](missing.md#anchor)` for a link.\n",
        })
        self.assertEqual((errors, count), ([], 0))

    def test_code_span_across_lines_is_ignored(self) -> None:
        # A code span may continue over line breaks within its paragraph.
        errors, _, count = run_check({
            "README.md": "Use this literal example: `\n[example](missing.md)\n`\n\n"
                         "Or ``this `one`\n[x](missing.md#a) ``.\n",
        })
        self.assertEqual((errors, count), ([], 0))

    def test_unclosed_backticks_do_not_hide_links(self) -> None:
        # An unmatched run is literal, and a span never crosses a blank line.
        errors, _, _ = run_check({
            "README.md": "A lone ` backtick\n[a](missing-a.md)\n\n"
                         "Open `here\n\n[b](missing-b.md) and `\n\n"
                         "Escaped \\`[c](missing-c.md)\\`\n",
        })
        self.assertEqual(errors, ["README.md: missing target: missing-a.md",
                                  "README.md: missing target: missing-b.md",
                                  "README.md: missing target: missing-c.md"])

    def test_indented_code_blocks_are_ignored(self) -> None:
        errors, _, count = run_check({
            "README.md": "Text\n\n    [a](missing.md)\n\n    [b](missing.md)\n\n"
                         "\t[c](missing.md)\n\n# Heading\n    [d](missing.md)\n\n"
                         "- item\n\n      [e](missing.md)\n\n"
                         "1. step\n\n    ```text\n    [f](missing.md)\n    ```\n",
        })
        self.assertEqual((errors, count), ([], 0))

    def test_indented_lines_that_are_not_code_are_checked(self) -> None:
        # Indented code cannot interrupt a paragraph, and list content is
        # indented relative to its item, not to the margin.
        errors, _, _ = run_check({
            "README.md": "Paragraph\n    [a](missing-a.md)\n\n"
                         "- item\n\n    [b](missing-b.md)\n\n"
                         "  - nested\n\n    [c](missing-c.md)\n\n"
                         "1. step\n   continued [d](missing-d.md)\n",
        })
        self.assertEqual(errors, [f"README.md: missing target: missing-{name}.md"
                                  for name in "abcd"])

    def test_character_references_in_destinations_are_decoded(self) -> None:
        errors, external, _ = run_check({
            "README.md": '<a href="notes&amp;examples.md#top">x</a> '
                         "[y](notes&amp;examples.md) [z](notes\\&examples.md)\n"
                         '<img src="a&#32;b.svg"> <a href="a&ampb.md">literal</a>\n'
                         '<a href="missing&amp;file.md">w</a>\n'
                         "[q](https://example.com/?a=1&amp;b=2)\n",
            "notes&examples.md": "# Top\n",
            "a b.svg": "<svg/>\n",
            "a&ampb.md": "# Literal\n",
        })
        self.assertEqual(errors, ["README.md: missing target: missing&file.md"])
        self.assertEqual(external, {"https://example.com/?a=1&b=2"})

    def test_escaped_parentheses_in_destinations(self) -> None:
        errors, _, count = run_check({
            "README.md": "[ok](notes\\).md) [ok](notes\\(draft.md) [ok](a(b).md#top)\n"
                         "[bad](gone\\).md)\n",
            "notes).md": "# Notes\n",
            "notes(draft.md": "# Draft\n",
            "a(b).md": "# Top\n",
        })
        self.assertEqual(count, 4)
        self.assertEqual(errors, ["README.md: missing target: gone).md"])

    def test_heading_emphasis_is_not_part_of_the_slug(self) -> None:
        errors, _, _ = run_check({
            "README.md": "[a](guide.md#api) [b](guide.md#snake_case-names) "
                         "[c](guide.md#__init__-method) [d](guide.md#old-new-bold-text)\n"
                         "[e](guide.md#_private-helper) [f](guide.md#api-1)\n"
                         "[g](guide.md#nowhere) [h](guide.md#api-2)\n",
            "guide.md": "# __API__\n## snake_case names\n## `__init__` method\n"
                        "## ~~old~~ *new* **bold** _text_\n## _private helper\n## API\n",
        })
        self.assertEqual(errors, ["README.md: missing anchor: guide.md#nowhere",
                                  "README.md: missing anchor: guide.md#api-2"])

    def test_missing_files_anchors_and_paths_are_errors(self) -> None:
        errors, _, _ = run_check({
            "README.md": "[a](docs/missing.md) [b](docs/guide.md#nowhere) [c](/abs.md)\n"
                         "[d](../outside.md) [e](empty.md) [f](ftp://example.com)\n"
                         "[g](#guide-1)\n",
            "docs/guide.md": "# Guide\n",
            "empty.md": "",
        })
        self.assertEqual(errors, [
            "README.md: missing target: docs/missing.md",
            "README.md: missing anchor: docs/guide.md#nowhere",
            "README.md: root-relative link: /abs.md",
            "README.md: path outside the repository: ../outside.md",
            "README.md: empty target: empty.md",
            "README.md: unsupported URL scheme: ftp://example.com",
            "README.md: missing anchor: #guide-1",
        ])

    def test_tracked_file_missing_from_the_working_tree_is_reported(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            (root / "README.md").write_text("[guide](guide.md)\n", encoding="utf-8")
            (root / "guide.md").write_text("# Guide\n", encoding="utf-8")
            subprocess.run(["git", "init", "-q", str(root)], check=True)
            subprocess.run(["git", "-C", str(root), "add", "."], check=True)
            (root / "guide.md").unlink()
            errors, _, _ = check_links.check(root)
        self.assertEqual(errors, ["guide.md: tracked but missing from the working tree",
                                  "README.md: missing target: guide.md"])

    def test_main_exit_codes(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            (root / "README.md").write_text("# Top\n[ok](#top)\n", encoding="utf-8")
            with contextlib.redirect_stdout(io.StringIO()) as out:
                self.assertEqual(check_links.main(["--root", str(root)]), 0)
            self.assertIn("Checked 1 link(s) in Markdown; 0 error(s).", out.getvalue())
            (root / "README.md").write_text("[bad](#nowhere)\n", encoding="utf-8")
            with contextlib.redirect_stdout(io.StringIO()) as out:
                self.assertEqual(check_links.main(["--root", str(root)]), 1)
            self.assertIn("ERROR: README.md: missing anchor: #nowhere", out.getvalue())


class ExternalTests(unittest.TestCase):
    def test_wiki_pages_are_requested_as_raw_markdown(self) -> None:
        request = check_links.request_url
        self.assertEqual(request("https://github.com/owner/repo/wiki/API-Reference", REPOSITORY),
                         "https://raw.githubusercontent.com/wiki/owner/repo/API-Reference.md")
        self.assertEqual(request("https://github.com/Owner/Repo/wiki", REPOSITORY),
                         "https://raw.githubusercontent.com/wiki/owner/repo/Home.md")
        for url in ("https://github.com/owner/repo/issues/11",
                    "https://github.com/owner/repo-other/wiki/Page",
                    "https://example.com/wiki/Page"):
            self.assertEqual(request(url, REPOSITORY), url)
        self.assertEqual(request("https://github.com/owner/repo/wiki/Page", ""),
                         "https://github.com/owner/repo/wiki/Page")

    def test_only_listed_pages_are_unchecked(self) -> None:
        reason = check_links.unchecked_reason
        self.assertIsNotNone(reason("https://github.com/owner/repo/security/code-scanning/1",
                                    REPOSITORY))
        self.assertIsNone(reason("https://github.com/owner/repo/issues/11", REPOSITORY))
        self.assertIsNone(reason("https://github.com/other/repo/security/code-scanning/1",
                                 REPOSITORY))

    def test_only_errors_fail_and_unchecked_pages_are_not_requested(self) -> None:
        wiki = "https://raw.githubusercontent.com/wiki/owner/repo/Gone.md"
        results = {"https://a.example/": ("PASS", "https://a.example/"),
                   "https://b.example/": ("WARN", "https://b.example/: unavailable"),
                   "https://c.example/": ("ERROR", "https://c.example/: HTTP 404"),
                   "https://github.com/owner/repo/issues/1": ("PASS", "issue"),
                   wiki: ("ERROR", "wiki page: HTTP 404")}
        with unittest.mock.patch.object(
                check_links, "check_external",
                side_effect=lambda url, requested=None: results[requested or url]) as fake, \
                contextlib.redirect_stdout(io.StringIO()) as out:
            errors = check_links.external_errors(
                ["https://a.example/", "https://b.example/", "https://c.example/",
                 "https://github.com/owner/repo/issues/1",
                 "https://github.com/owner/repo/wiki/Gone",
                 "https://github.com/owner/repo/security/code-scanning/3"], REPOSITORY)
        self.assertEqual(errors, ["https://c.example/: HTTP 404", "wiki page: HTTP 404"])
        self.assertEqual(sorted(call.args[1] for call in fake.call_args_list), sorted(results))
        self.assertIn("SKIP: https://github.com/owner/repo/security/code-scanning/3:",
                      out.getvalue())

    def test_http_codes_are_classified(self) -> None:
        def curl(code: str, returncode: int = 0) -> unittest.mock.Mock:
            return unittest.mock.Mock(stdout=code, returncode=returncode)

        for code, returncode, status in (("200", 0, "PASS"), ("404", 0, "ERROR"),
                                         ("403", 0, "WARN"), ("429", 0, "WARN"),
                                         ("503", 0, "WARN"), ("", 6, "WARN")):
            with self.subTest(code=code), unittest.mock.patch(
                    "check_links.subprocess.run", return_value=curl(code, returncode)):
                self.assertEqual(check_links.check_external("https://x.example/")[0], status)

    def test_requested_url_is_reported(self) -> None:
        with unittest.mock.patch("check_links.subprocess.run",
                                 return_value=unittest.mock.Mock(stdout="404",
                                                                 returncode=0)) as run:
            status, message = check_links.check_external("https://a.example/w", "https://b.example/w.md")
        self.assertEqual(status, "ERROR")
        self.assertEqual(message, "https://a.example/w (requested as https://b.example/w.md): HTTP 404")
        self.assertEqual(run.call_args.args[0][-1], "https://b.example/w.md")


if __name__ == "__main__":
    unittest.main()
