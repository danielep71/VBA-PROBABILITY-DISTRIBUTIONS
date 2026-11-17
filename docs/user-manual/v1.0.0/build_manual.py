"""Build the user manual PDF from the Markdown chapters of this directory.

    python3 build_manual.py            writes build/manual.html and the PDF
    python3 build_manual.py --check    also fails if a K_STATS_ name or the
                                       argument count of a call is not
                                       what docs/PUBLIC_API.txt declares

Tools: pandoc (Markdown to HTML with MathML and a table of contents) and a
Chromium or Chrome binary, version 131 or later (HTML to PDF; the running
footer and page numbers come from manual.css). Both versions are printed, and
recorded next to the PDF.
"""
import argparse
import hashlib
import json
import os
import pathlib
import re
import shutil
import subprocess
import sys

HERE = pathlib.Path(__file__).resolve().parent
ROOT = HERE.parents[2]
VERSION = HERE.name                                   # the directory is the version
TITLE = "VBA Probability Distributions — User Manual"
PDF_NAME = f"VBA-Probability-Distributions-User-Manual-{VERSION}.pdf"
BUILD = HERE / "build"


def chapters():
    return sorted(HERE.glob("[0-9][0-9]-*.md"))


def declarations():
    """{name: (required, total)} argument counts from docs/PUBLIC_API.txt."""
    out = {}
    for line in (ROOT / "docs" / "PUBLIC_API.txt").read_text(encoding="utf-8").splitlines():
        m = re.search(r"Function (K_STATS_\w+)\((.*)\)", line)
        if not m:
            continue
        params = [p for p in m.group(2).split(",") if p.strip()]
        out[m.group(1)] = (sum("Optional" not in p for p in params), len(params))
    return out


def calls(text):
    """(name, argument count) of every K_STATS_ call written with parentheses."""
    for m in re.finditer(r"(K_STATS_\w+)\(", text):
        depth, count, seen, i = 1, 0, False, m.end()
        while i < len(text) and depth:
            c = text[i]
            if c == "(":
                depth += 1
            elif c == ")":
                depth -= 1
            elif c == "," and depth == 1:
                count += 1
            if depth and not c.isspace():
                seen = True
            i += 1
        if depth == 0 and "..." not in text[m.end():i]:     # an elided call is prose
            yield m.group(1), (count + 1 if seen else 0)


def problems():
    known = declarations()
    found = []
    for path in chapters():
        text = path.read_text(encoding="utf-8")
        for name in sorted(set(re.findall(r"K_STATS_\w+", text)) - set(known)):
            found.append(f"{name} ({path.name}) is not in docs/PUBLIC_API.txt")
        for name, count in calls(text):
            if name in known and not known[name][0] <= count <= known[name][1]:
                found.append(f"{name} ({path.name}) is called with {count} argument(s); "
                             f"it takes {known[name][0]} to {known[name][1]}")
    return found


def html():
    BUILD.mkdir(exist_ok=True)
    out = BUILD / "manual.html"
    front, body = chapters()[0], chapters()[1:]
    command = ["pandoc", "--from", "gfm+tex_math_dollars", "--to", "html5", "--standalone",
               "--mathml", "--toc", "--toc-depth=2",
               "--metadata", f"title={TITLE}", "--metadata", f"subtitle=Version {VERSION}",
               "--metadata", "lang=en", "--css", "manual.css", "--embed-resources",
               "--resource-path", str(HERE), "--output", str(out),
               str(front), *map(str, body)]
    subprocess.run(command, check=True, cwd=HERE)
    return out


def browser():
    """A Chromium or Chrome binary: $CHROME, or the first one found."""
    candidates = [os.environ.get("CHROME", ""), "chromium", "chromium-browser", "google-chrome",
                  "chrome"]
    for name in candidates:
        path = shutil.which(name) if name else None
        if path:
            return path
    raise SystemExit("no Chromium or Chrome found; set CHROME to its path")


def pdf(source):
    target = BUILD / PDF_NAME
    binary = browser()
    subprocess.run([binary, "--headless", "--no-sandbox", "--disable-gpu", "--no-pdf-header-footer",
                    f"--print-to-pdf={target}", source.as_uri()],
                   check=True, capture_output=True)
    version = subprocess.run([binary, "--version"], capture_output=True, text=True).stdout.strip()
    return target, version


def main():
    ap = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    ap.add_argument("--check", action="store_true")
    a = ap.parse_args()
    if a.check:
        found = problems()
        for finding in found:
            print(f"FAIL: {finding}")
        if found:
            return 1
        print("PASS: every K_STATS_ name in the manual is in docs/PUBLIC_API.txt, "
              "and every call has an argument count its declaration accepts")
    source = html()
    target, chromium = pdf(source)
    pandoc = subprocess.run(["pandoc", "--version"], capture_output=True, text=True).stdout.splitlines()[0]
    digest = hashlib.sha256(target.read_bytes()).hexdigest()
    record = {"pdf": target.name, "sha256": digest, "version": VERSION,
              "chapters": [c.name for c in chapters()], "pandoc": pandoc, "browser": chromium}
    (BUILD / "build_record.json").write_text(json.dumps(record, indent=1) + "\n", encoding="utf-8")
    print(f"{target.name}: {target.stat().st_size} bytes, sha256 {digest}")
    print(f"built with {pandoc} and {chromium}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
