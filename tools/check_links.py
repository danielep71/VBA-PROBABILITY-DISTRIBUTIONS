#!/usr/bin/env python3
"""Check the links in the repository's tracked Markdown files.

Offline (the default, run by tools/run_static_checks.py on every change), every
relative link, in Markdown syntax or in an HTML href or src attribute, must name
an existing directory or non-empty file inside the repository, after URL
decoding (%20 is a space). Every fragment must name a heading anchor, generated
with GitHub's slug rules, or an explicit <a id> or <a name> anchor of the
Markdown file it points to. Links inside fenced or inline code are ignored.
External links, including the wiki, issues and other pages of this repository
on github.com, are only collected.

With --external (the weekly Link check workflow), each external http(s) link is
also requested once, with timeouts and retries:

- a page of this repository's wiki is requested as its raw Markdown, which
  answers 404 when the page does not exist;
- links listed in UNCHECKED are not requested; each entry gives its reason;
- HTTP 403, 429 and 5xx answers and network failures are warnings, because
  they say nothing about the link (rate limits, bot protection, outages);
  any other non-2xx answer is an error.
"""

from __future__ import annotations

import argparse
import html
import os
import re
import subprocess
import sys
import unicodedata
from collections.abc import Iterable, Sequence
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path
from urllib.parse import unquote, urlsplit

ROOT = Path(__file__).resolve().parents[1]
FENCE = re.compile(r" {0,3}(`{3,}|~{3,})")
ATTRIBUTE = r"""\s*=\s*(?:"([^"]*)"|'([^']*)'|([^\s"'=<>`]+))"""

# Pages of this repository that --external deliberately does not request:
# (path prefix after https://github.com/<owner>/<name>, reason).
UNCHECKED: tuple[tuple[str, str], ...] = (
    ("/security/code-scanning/",
     "code-scanning alerts are visible only to signed-in users with write access"),
)


# ---------------------------------------------------------------------------
# Markdown parsing.

def github_slug(heading: str) -> str:
    """Anchor GitHub generates for a Markdown heading, before duplicate suffixes."""
    text = re.sub(r"!?\[([^\]]*)\]\([^)]*\)", r"\1", heading)
    # Decode only complete character references (ending in ';'), as Markdown does.
    text = re.sub(r"&(?:#[0-9]{1,7}|#[xX][0-9a-fA-F]{1,6}|[A-Za-z][A-Za-z0-9]{1,31});",
                  lambda match: html.unescape(match.group(0)), re.sub(r"<[^>]+>", "", text))
    return "".join(char for char in text.lower() if slug_character(char)).replace(" ", "-")


def slug_character(char: str) -> bool:
    """GitHub keeps word characters (letters, marks, decimal digits and connector
    punctuation, such as the variation selector U+FE0F after many emoji), "-" and spaces."""
    return char in "- " or unicodedata.category(char)[:1] in ("L", "M") \
        or unicodedata.category(char) in ("Nd", "Pc")


def heading_anchor(slug: str, used: dict[str, int]) -> str:
    """Make a heading slug unique the way GitHub does: -1, -2... until unused."""
    anchor = slug
    while anchor in used:
        used[slug] += 1
        anchor = f"{slug}-{used[slug]}"
    used[anchor] = 0
    return anchor


def inline_destinations(line: str) -> list[str]:
    """Destinations of [text](destination "title"), allowing <...> and balanced parentheses."""
    found = []
    for match in re.finditer(r"\]\(", line):
        rest = line[match.end():].lstrip()
        if rest.startswith("<"):
            end = rest.find(">")
            if end > 1:
                found.append(rest[1:end])
            continue
        depth, destination = 0, ""
        for char in rest:
            if char.isspace() or (char == ")" and depth == 0):
                break
            depth += {"(": 1, ")": -1}.get(char, 0)
            destination += char
        if destination:
            found.append(destination)
    return found


def ends_paragraph(line: str, heading: bool) -> bool:
    return (heading or not line.strip()
            or bool(re.match(r" {0,3}([-*_])(\s*\1){2,}\s*$", line))
            or bool(re.match(r" {4}| {0,3}([-*+]\s|\d+[.)]\s|[>|<])", line)))


def line_links(line: str) -> tuple[list[str], set[str]]:
    """Links and explicit anchors on one line outside code."""
    line = re.sub(r"(`+).*?\1", "", line)
    anchors = {"".join(value) for value in
               re.findall(r"<a\s[^>]*?(?<![\w-])(?:id|name)" + ATTRIBUTE, line)}
    links = inline_destinations(line)
    links += ["".join(value) for value in
              re.findall(r"^ {0,3}\[(?!\^)[^\]]+\]:\s*(?:<([^>]*)>|(\S+))", line)]
    links += re.findall(r"<(https?://[^>\s]+)>", line)
    links += ["".join(value) for value in re.findall(r"(?<![\w-])(?:href|src)" + ATTRIBUTE, line)]
    return links, anchors


def markdown_links(text: str) -> tuple[list[str], set[str]]:
    """Return (links, anchors) of a Markdown document, ignoring fenced and inline code."""
    links: list[str] = []
    anchors: set[str] = set()
    used: dict[str, int] = {}
    fence: str | None = None
    paragraph: list[str] = []
    for line in text.splitlines():
        match = FENCE.match(line)
        if fence:
            if match and match.group(1)[0] == fence[0] and len(match.group(1)) >= len(fence):
                fence = None
            continue
        if match:
            fence, paragraph = match.group(1), []
            continue
        if paragraph and re.match(r" {0,3}(=+|-+)\s*$", line):
            # Setext heading: the paragraph above an === or --- underline.
            anchors.add(heading_anchor(github_slug(" ".join(paragraph)), used))
            paragraph = []
            continue
        heading = re.match(r" {0,3}#{1,6}\s+(.*?)(?:\s+#+)?\s*$", line)
        if heading:
            anchors.add(heading_anchor(github_slug(heading.group(1)), used))
        paragraph = [] if ends_paragraph(line, bool(heading)) else [*paragraph, line.strip()]
        found, explicit = line_links(line)
        links += found
        anchors |= explicit
    return links, anchors


# ---------------------------------------------------------------------------
# Offline check.

def tracked_markdown(root: Path) -> list[Path]:
    try:
        names = subprocess.run(["git", "-C", str(root), "ls-files", "-z", "--", "*.md"],
                               capture_output=True, text=True, encoding="utf-8",
                               check=True).stdout.split("\0")
        return [root / name for name in names if name]
    except (OSError, subprocess.CalledProcessError):
        return [path for path in root.rglob("*.md") if ".git" not in path.relative_to(root).parts]


def local_error(root: Path, source: Path, link: str, anchors: dict[Path, set[str]]) -> str | None:
    url = urlsplit(link)
    if url.path.startswith("/"):
        return f"root-relative link: {link}"
    target = (source.parent / unquote(url.path)).resolve() if url.path else source.resolve()
    if not target.is_relative_to(root.resolve()):
        return f"path outside the repository: {link}"
    if not target.exists():
        return f"missing target: {link}"
    if target.is_file() and target.stat().st_size == 0:
        return f"empty target: {link}"
    if url.fragment and target in anchors and unquote(url.fragment) not in anchors[target]:
        return f"missing anchor: {link}"
    return None


def check(root: Path = ROOT) -> tuple[list[str], set[str], int]:
    """Return (errors, external URLs, number of links checked) for the Markdown files."""
    errors: list[str] = []
    documents: dict[Path, tuple[list[str], set[str]]] = {}
    for path in tracked_markdown(root):
        if not path.is_file():
            # Tracked by Git but deleted from the working tree without git rm.
            errors.append(f"{path.relative_to(root).as_posix()}: tracked but missing from the working tree")
            continue
        documents[path.resolve()] = markdown_links(path.read_text(encoding="utf-8"))
    anchors = {path: found for path, (_, found) in documents.items()}
    external: set[str] = set()
    count = 0
    for source, (links, _) in sorted(documents.items()):
        for link in links:
            count += 1
            url = urlsplit(link)
            if url.scheme in ("http", "https") or (not url.scheme and url.netloc):
                external.add(("https:" + link if link.startswith("//") else link).split("#")[0])
                continue
            if url.scheme in ("mailto", "tel", "data"):
                continue
            problem = (f"unsupported URL scheme: {link}" if url.scheme
                       else local_error(root, source, link, anchors))
            if problem:
                errors.append(f"{source.relative_to(root.resolve()).as_posix()}: {problem}")
    return errors, external, count


# ---------------------------------------------------------------------------
# External check.

def own_path(url: str, repository: str) -> str | None:
    """The path after https://github.com/<repository> ("" for the repository's
    own page), or None when the URL does not point into this repository on github.com."""
    if not repository:
        return None
    prefix = f"https://github.com/{repository}".lower()
    lowered = url.lower()
    if lowered != prefix and not lowered.startswith(prefix + "/"):
        return None
    return url[len(prefix):]


def unchecked_reason(url: str, repository: str) -> str | None:
    """Why a link is deliberately not requested, or None when it is requested."""
    path = own_path(url, repository)
    if path is None:
        return None
    return next((reason for prefix, reason in UNCHECKED if path.startswith(prefix)), None)


def request_url(url: str, repository: str) -> str:
    """The URL to request for a link: a page of this repository's wiki is read as
    its raw Markdown, which answers 404 when the page does not exist."""
    path = own_path(url, repository)
    if path is None:
        return url
    page = re.fullmatch(r"/wiki/?([^/?#]*)", path.split("?")[0])
    if page is None:
        return url
    return f"https://raw.githubusercontent.com/wiki/{repository}/{page.group(1) or 'Home'}.md"


def check_external(url: str, requested: str | None = None) -> tuple[str, str]:
    target = requested or url
    label = url if target == url else f"{url} (requested as {target})"
    # GET follows redirects; HEAD is rejected by some otherwise working sites.
    result = subprocess.run([
        "curl", "--silent", "--show-error", "--location", "--output", os.devnull,
        "--write-out", "%{http_code}", "--connect-timeout", "10", "--max-time", "25",
        "--retry", "2", "--retry-delay", "2", "--proto", "=https,http",
        "--proto-redir", "=https,http", target,
    ], capture_output=True, text=True, timeout=95, check=False)
    code = result.stdout.strip()
    if result.returncode or code in ("403", "429") or code.startswith("5"):
        return "WARN", f"{label}: unavailable for verification (HTTP {code or 'unknown'})"
    if not code.startswith("2"):
        return "ERROR", f"{label}: HTTP {code}"
    return "PASS", label


def external_errors(urls: Iterable[str], repository: str) -> list[str]:
    requested: list[str] = []
    for url in sorted(urls):
        reason = unchecked_reason(url, repository)
        if reason:
            print(f"SKIP: {url}: {reason}")
        else:
            requested.append(url)
    errors = []
    with ThreadPoolExecutor(max_workers=3) as pool:
        for status, message in pool.map(
                lambda url: check_external(url, request_url(url, repository)), requested):
            print(f"{status}: {message}")
            if status == "ERROR":
                errors.append(message)
    return errors


def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=(__doc__ or "").splitlines()[0])
    parser.add_argument("--root", type=Path, default=ROOT)
    parser.add_argument("--external", action="store_true",
                        help="also request every external http(s) link")
    parser.add_argument("--repository", default=os.environ.get("GITHUB_REPOSITORY", ""),
                        help="owner/name of this repository, whose wiki pages and unchecked "
                             "pages are recognised (default: GITHUB_REPOSITORY)")
    args = parser.parse_args(argv)
    root = args.root.resolve()
    errors, external, count = check(root)
    if args.external:
        errors += external_errors(external, args.repository)
    for error in errors:
        print(f"ERROR: {error}")
    print(f"Checked {count} link(s) in Markdown; {len(errors)} error(s).")
    return 1 if errors else 0


if __name__ == "__main__":
    sys.exit(main())
