#!/usr/bin/env python3
"""Check the links in the repository's tracked Markdown files.

Offline (the default, run by tools/run_static_checks.py on every change), every
relative link, in Markdown syntax or in an HTML href or src attribute, must name
an existing directory or non-empty file inside the repository. The path is read
as GitHub reads it: Markdown backslash escapes removed, character references
decoded (&amp; is &), then URL-decoded (%20 is a space). Every fragment must name a heading
anchor, generated with GitHub's slug rules, or an explicit <a id> or <a name>
anchor of the Markdown file it points to. Links inside fenced code blocks,
indented code blocks, code spans (which may continue over line breaks) and
HTML comments are ignored, and so is text that only looks like a link, such as
\[text](path). Blockquote and list-item markers are read before fences and
headings. External links, including the wiki, issues and other pages of this
repository on github.com, are only collected.

The checker approximates GitHub's Markdown renderer; it is not a full CommonMark
parser. A link it reports wrongly is listed, with the reason, in
tools/link_check_exclusions.json; an entry that no longer matches a report is
itself reported, so the list cannot go stale.

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
import json
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
EXCLUSIONS = "tools/link_check_exclusions.json"
ATTRIBUTE = r"""\s*=\s*(?:"([^"]*)"|'([^']*)'|([^\s"'=<>`]+))"""

# Pages of this repository that --external deliberately does not request:
# (path prefix after https://github.com/<owner>/<name>, reason).
UNCHECKED: tuple[tuple[str, str], ...] = (
    ("/security/code-scanning/",
     "code-scanning alerts are visible only to signed-in users with write access"),
)


# ---------------------------------------------------------------------------
# Markdown parsing.

REFERENCE = re.compile(r"&(?:#[0-9]{1,7}|#[xX][0-9a-fA-F]{1,6}|[A-Za-z][A-Za-z0-9]{1,31});")
ESCAPABLE = re.compile(r"[!-/:-@\[-`{-~]")  # ASCII punctuation, escapable with a backslash
LIST_ITEM = re.compile(r"( *)([-*+]|\d{1,9}[.)])( {1,4}|$)")
QUOTE = re.compile(r" {0,3}>[ \t]?")
COMMENT = re.compile(r"<!--.*?-->", re.DOTALL)


def decode_references(text: str) -> str:
    """Decode complete character references (ending in ';'), as Markdown and HTML
    attributes do: notes&amp;examples.md names notes&examples.md."""
    return REFERENCE.sub(lambda match: html.unescape(match.group(0)), text)


def decode_destination(text: str) -> str:
    """A Markdown link destination as the renderer reads it: backslash escapes of
    ASCII punctuation removed, then character references decoded."""
    return decode_references(re.sub(r"\\(" + ESCAPABLE.pattern + ")", r"\1", text))


def github_slug(heading: str, emphasis: bool = True) -> str:
    """Anchor GitHub generates for a Markdown heading, before duplicate suffixes.

    GitHub slugs the rendered text. "*" and "~" are punctuation, which the slug
    drops anyway; "_" is kept, so with emphasis (the default) the "_" runs that
    delimit emphasis are removed first: "__API__" gives "api", while
    "snake_case" and "`__init__`" keep their underscores.
    """
    text = re.sub(r"!?\[([^\]]*)\]\([^)]*\)", r"\1", heading)
    if emphasis:
        text = strip_underscore_emphasis(text)
    text = decode_references(re.sub(r"<[^>]+>", "", text))
    return "".join(char for char in text.lower() if slug_character(char)).replace(" ", "-")


def strip_underscore_emphasis(text: str) -> str:
    """Remove the "_" runs that delimit emphasis, outside code spans.

    A run can open when it is followed by a non-space and not preceded by a
    letter or digit, and close when it is preceded by a non-space and not
    followed by a letter or digit (CommonMark's rule for "_": never inside a
    word). An opener pairs with the next closer of the same length; unpaired
    runs are literal. heading_anchors also accepts the slug without this step,
    so a case that GitHub renders differently cannot cause a false error.
    """
    parts = re.split(r"((`+).+?\2)", text)
    # re.split gives [text, code span, its backticks, text, ...]; only text changes.
    for index in range(0, len(parts), 3):
        parts[index] = strip_segment_emphasis(parts[index])
    return "".join(part for index, part in enumerate(parts) if index % 3 != 2)


def strip_segment_emphasis(segment: str) -> str:
    def side(position: int) -> str:
        return segment[position] if 0 <= position < len(segment) else " "

    runs = list(re.finditer(r"_+", segment))
    removed: set[int] = set()
    for number, opener in enumerate(runs):
        if number in removed or side(opener.end()).isspace() \
                or side(opener.start() - 1).isalnum():
            continue
        for later in range(number + 1, len(runs)):
            closer = runs[later]
            if (len(closer.group()) == len(opener.group())
                    and not side(closer.start() - 1).isspace()
                    and not side(closer.end()).isalnum()):
                removed |= {number, later}
                break
    kept, position = [], 0
    for number, run in enumerate(runs):
        kept.append(segment[position:run.start()] + ("" if number in removed else run.group()))
        position = run.end()
    return "".join(kept) + segment[position:]


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


def heading_anchors(heading: str, used: dict[str, int]) -> set[str]:
    """The anchor of a heading, plus its slug with "_" emphasis left in place, which
    is accepted too: where GitHub's rendering is ambiguous, no error is reported."""
    return {heading_anchor(github_slug(heading), used), github_slug(heading, emphasis=False)}


def inline_destinations(text: str) -> list[str]:
    """Destinations of [text](destination "title") and ![alt](destination), allowing
    <...> and balanced parentheses. A "](" counts only after an unescaped "[" in the
    same block of text, so \\[text](path) is plain text."""
    found = []
    openers = 0
    index = 0
    while index < len(text):
        char = text[index]
        if char == "\\" and ESCAPABLE.fullmatch(text[index + 1:index + 2]):
            index += 2
            continue
        if char == "[":
            openers += 1
        elif char == "]" and openers:
            openers -= 1
            if text.startswith("(", index + 1):
                destination = link_destination(text[index + 2:])
                if destination:
                    found.append(destination)
        index += 1
    return found


def link_destination(rest: str) -> str:
    """The destination at the start of the text after "](", or "" if there is none."""
    rest = rest.lstrip(" \t")
    if rest.startswith("<"):
        end = rest.find(">")
        return rest[1:end] if end > 1 and "\n" not in rest[:end] else ""
    depth, destination, index = 0, "", 0
    while index < len(rest):
        char = rest[index]
        if char == "\\" and ESCAPABLE.fullmatch(rest[index + 1:index + 2]):
            # An escaped character, such as \\) or \\(, never delimits.
            destination += rest[index:index + 2]
            index += 2
            continue
        if char.isspace() or (char == ")" and depth == 0):
            break
        depth += {"(": 1, ")": -1}.get(char, 0)
        destination += char
        index += 1
    return destination


def ends_paragraph(line: str, heading: bool) -> bool:
    return (heading or not line.strip()
            or bool(re.match(r" {0,3}([-*_])(\s*\1){2,}\s*$", line))
            or bool(re.match(r" {4}| {0,3}([-*+]\s|\d+[.)]\s|[>|<])", line)))


def strip_code_spans(text: str) -> str:
    """Remove the code spans of one block of text, which may span several lines.

    As in CommonMark, a backtick run opens a span only if an equally long run
    follows in the same block; otherwise it is literal. The newlines inside a
    removed span are kept, so the text keeps its lines.
    """
    runs = list(re.finditer(r"`+", text))
    kept, position, index = [], 0, 0
    while index < len(runs):
        opener = runs[index]
        escaped = opener.start() > 0 and text[opener.start() - 1] == "\\"
        closer = None if escaped else next(
            (later for later in range(index + 1, len(runs))
             if len(runs[later].group()) == len(opener.group())), None)
        if closer is None:
            index += 1
            continue
        kept += [text[position:opener.start()],
                 "\n" * text.count("\n", opener.start(), runs[closer].end())]
        position, index = runs[closer].end(), closer + 1
    return "".join(kept) + text[position:]


def line_links(line: str) -> tuple[list[str], set[str]]:
    """Reference definitions, autolinks, HTML links and explicit anchors on one line
    whose code spans and comments are already removed. Destinations are decoded the
    way GitHub reads them."""
    anchors = {decode_references("".join(value)) for value in
               re.findall(r"<a\s[^>]*?(?<![\w-])(?:id|name)" + ATTRIBUTE, line)}
    links = [decode_destination("".join(value)) for value in
              re.findall(r"^ {0,3}\[(?!\^)[^\]]+\]:\s*(?:<([^>]*)>|(\S+))", line)]
    links += re.findall(r"<(https?://[^>\s]+)>", line)
    links += [decode_references("".join(value)) for value in
              re.findall(r"(?<![\w-])(?:href|src)" + ATTRIBUTE, line)]
    return links, anchors


def block_links(lines: Sequence[str]) -> tuple[list[str], set[str]]:
    """Links and explicit anchors of one block of text lines, outside its code spans
    and HTML comments."""
    text = strip_code_spans("\n".join(lines))
    text = COMMENT.sub(lambda match: "\n" * match.group().count("\n"), text)
    links = [decode_destination(link) for link in inline_destinations(text)]
    anchors: set[str] = set()
    for line in text.split("\n"):
        found, explicit = line_links(line)
        links += found
        anchors |= explicit
    return links, anchors


def indentation(line: str) -> int:
    """Leading columns of a line, with tabs to the next multiple of 4."""
    expanded = line.expandtabs(4)
    return len(expanded) - len(expanded.lstrip(" "))


def strip_quotes(line: str) -> tuple[int, str]:
    """Blockquote depth of a line, and the line without its ">" markers."""
    depth = 0
    while match := QUOTE.match(line):
        depth, line = depth + 1, line[match.end():]
    return depth, line


def markdown_links(text: str) -> tuple[list[str], set[str]]:
    """Return (links, anchors) of a Markdown document, ignoring fenced and indented
    code blocks, code spans and HTML comments.

    Blockquote markers are removed, and list-item markers read, before fences and
    headings are recognised. Indented code follows CommonMark conservatively: a
    line indented four columns beyond the content of the open list item (or
    beyond the margin, outside lists) is code unless it continues a paragraph,
    which code cannot interrupt.
    """
    links: list[str] = []
    anchors: set[str] = set()
    used: dict[str, int] = {}
    fence: tuple[str, int, int] | None = None  # (opening run, content column, quote depth)
    comment = False                        # inside an HTML comment block
    quote = 0                              # blockquote depth of the previous line
    paragraph: list[str] = []              # candidate setext heading text
    block: list[str] = []                  # lines whose code spans and links are read together
    items: list[int] = []                  # content columns of the open list items
    previous = "blank"                     # blank, text, code or closed (heading, fence, break)

    def flush() -> None:
        found, explicit = block_links(block)
        links.extend(found)
        anchors.update(explicit)
        block.clear()

    for raw in text.splitlines():
        depth, line = strip_quotes(raw)
        if depth < quote and previous == "text" and line.strip() and not fence:
            depth = quote   # a lazy continuation line of the quoted paragraph
        if depth != quote:
            # A new or closed blockquote: its blocks, lists and fences end here.
            flush()
            if fence and depth < fence[2]:
                fence = None
            items, paragraph, previous, quote = [], [], "blank", depth
        expanded = line.expandtabs(4)
        indent = indentation(line)
        if fence:
            relative = expanded[min(indent, fence[1]):]
            match = FENCE.match(relative)
            if (match and match.group(1)[0] == fence[0][0]
                    and len(match.group(1)) >= len(fence[0])
                    and not relative[match.end():].strip()):
                fence, previous = None, "closed"
            continue
        if comment:
            # An HTML comment block ends on the line that contains "-->".
            if "-->" in line:
                comment, previous = False, "closed"
            continue
        if not line.strip():
            flush()
            paragraph, previous = [], "blank"
            continue
        base = items[-1] if items else 0
        thematic = re.match(r" {0,3}([-*_])(\s*\1){2,}\s*$", expanded[min(indent, base):])
        item = None if thematic or indent >= base + 4 else LIST_ITEM.match(expanded)
        if item or previous != "text":
            # A list item ends where a line starts left of its content.
            while items and indent < items[-1]:
                items.pop()
            base = items[-1] if items else 0
        if not item and indent >= base + 4 and previous != "text":
            flush()
            paragraph, previous = [], "code"
            continue
        if item:
            flush()
            # The item's content is read as a block of its own: a fence or a heading
            # may start on the marker's line.
            base = len(item.group(0)) + (0 if item.group(3) else 1)
            items.append(base)
            relative = expanded[base:]
        else:
            relative = expanded[min(indent, base):]
        match = FENCE.match(relative)
        if match and not (match.group(1)[0] == "`" and "`" in relative[match.end():]):
            flush()
            fence, paragraph = (match.group(1), base, quote), []
            continue
        if re.match(r" {0,3}<!--", relative):
            flush()
            comment = "-->" not in relative[relative.index("<!--") + 4:]
            paragraph, previous = [], "closed"
            continue
        if paragraph and not item and re.match(r" {0,3}(=+|-+)\s*$", relative):
            # Setext heading: the paragraph above an === or --- underline.
            anchors |= heading_anchors(" ".join(paragraph), used)
            flush()
            paragraph, previous = [], "closed"
            continue
        heading = re.match(r" {0,3}#{1,6}\s+(.*?)(?:\s+#+)?\s*$", relative)
        if heading or thematic:
            flush()
        if heading:
            anchors |= heading_anchors(heading.group(1), used)
        paragraph = [] if ends_paragraph(line, bool(heading)) else [*paragraph, line.strip()]
        block.append(relative if item else line)
        if heading or thematic:
            flush()
            previous = "closed"
        else:
            previous = "text"
    flush()
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


def load_exclusions(root: Path) -> tuple[list[dict[str, str]], list[str]]:
    """The reviewed exclusions of root, and the problems with the file itself.

    Each entry names a Markdown file ("path", relative to the root), a link as
    the report prints it ("link") and why it is not an error ("reason").
    """
    path = root / EXCLUSIONS
    if not path.is_file():
        return [], []
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
        entries = data["exclusions"]
        if data.get("schema_version") != 1 or not isinstance(entries, list):
            raise ValueError("schema_version must be 1 and exclusions a list")
    except (OSError, ValueError, KeyError, TypeError) as error:
        return [], [f"{EXCLUSIONS}: unreadable: {error}"]
    valid, problems = [], []
    for number, entry in enumerate(entries):
        if (not isinstance(entry, dict) or set(entry) != {"path", "link", "reason"}
                or not all(isinstance(value, str) and value.strip() for value in entry.values())):
            problems.append(f"{EXCLUSIONS}: entry {number} needs exactly a non-empty "
                            "path, link and reason")
        else:
            valid.append(entry)
    return valid, problems


def check(root: Path = ROOT) -> tuple[list[str], set[str], int]:
    """Return (errors, external URLs, number of links checked) for the Markdown files.

    A reported link listed in the exclusions file is not an error; an exclusion
    that matches no reported link is.
    """
    root = root.resolve()
    errors: list[str] = []
    documents: dict[Path, tuple[list[str], set[str]]] = {}
    for path in tracked_markdown(root):
        if not path.is_file():
            # Tracked by Git but deleted from the working tree without git rm.
            errors.append(f"{path.relative_to(root).as_posix()}: tracked but missing from the working tree")
            continue
        documents[path.resolve()] = markdown_links(path.read_text(encoding="utf-8"))
    anchors = {path: found for path, (_, found) in documents.items()}
    exclusions, problems = load_exclusions(root)
    excluded = {(entry["path"], entry["link"]) for entry in exclusions}
    unmatched = set(excluded)
    external: set[str] = set()
    count = 0
    for source, (links, _) in sorted(documents.items()):
        name = source.relative_to(root).as_posix()
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
            if problem and (name, link) in excluded:
                unmatched.discard((name, link))
            elif problem:
                errors.append(f"{name}: {problem}")
    errors += problems
    errors += [f"{EXCLUSIONS}: stale exclusion, {path} no longer reports {link}"
               for path, link in sorted(unmatched)]
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
