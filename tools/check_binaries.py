#!/usr/bin/env python3
"""Allow tracked binary files only by approved path or by reviewed SHA-256.

The check reads every regular file in the Git index (in CI, the checked-out
commit) and the allowlist from the index too, so stage a change to check it.
A file is binary when .gitattributes unsets its text attribute (the
office-doc, opaque and vba-form-bin macros, or an explicit -text), or when its
content contains a NUL byte. A binary file passes only if it is:

- an image directly under assets/ (.png, .jpg or .jpeg) whose bytes start with
  the signature of that format; or
- listed in .github/binary-allowlist.json with the SHA-256 of its content and
  a reason.

Every other binary fails, and so does an allowlisted file whose content no
longer has the recorded SHA-256, an allowlist entry for a file that is not
tracked, and a NUL-containing file that .gitattributes does not declare binary
(Git could rewrite its line endings, so it must be declared binary before it can
be approved). The SHA-256 is that of the blob Git stores; Git converts nothing
in a file declared binary, so it equals `sha256sum` of the checked-out file.

The check never changes a file.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import re
import subprocess
import sys
from collections.abc import Sequence
from dataclasses import dataclass
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parents[1]
ALLOWLIST = ".github/binary-allowlist.json"
# Images the README and social preview use: one directory, no subdirectories.
APPROVED_IMAGE = re.compile(r"assets/[^/]+\.(png|jpg|jpeg)")
SIGNATURES = {"png": b"\x89PNG\r\n\x1a\n", "jpg": b"\xff\xd8\xff", "jpeg": b"\xff\xd8\xff"}
SHA256 = re.compile(r"[0-9a-f]{64}")
REGULAR_MODES = {"100644", "100755"}


@dataclass(frozen=True)
class Binary:
    """A tracked binary file and how it was allowed."""

    path: str
    sha256: str
    allowed_by: str


def git(root: Path, *args: str, stdin: bytes | None = None) -> bytes:
    return subprocess.run(["git", *args], cwd=root, input=stdin, capture_output=True,
                          check=True).stdout


def tracked_blobs(root: Path) -> dict[str, str]:
    """Map the path of every regular file in the index to its blob id."""
    blobs: dict[str, str] = {}
    for record in git(root, "ls-files", "-s", "-z").split(b"\0"):
        if not record:
            continue
        info, _, name = record.decode("utf-8", "surrogateescape").partition("\t")
        mode, oid, _stage = info.split()
        if mode in REGULAR_MODES:
            blobs[name] = oid
    return blobs


def blob_contents(root: Path, oids: Sequence[str]) -> list[bytes]:
    """The bytes of each blob, in order, read with one git cat-file process."""
    if not oids:
        return []
    output = git(root, "cat-file", "--batch", stdin="".join(f"{oid}\n" for oid in oids).encode())
    contents: list[bytes] = []
    position = 0
    for _ in oids:
        header_end = output.index(b"\n", position)
        size = int(output[position:header_end].split()[2])
        start = header_end + 1
        contents.append(output[start:start + size])
        position = start + size + 1
    return contents


def text_attributes(root: Path, names: Sequence[str]) -> dict[str, str]:
    """The text attribute of each path: set, unset, auto, unspecified or a value."""
    if not names:
        return {}
    stdin = "".join(f"{name}\0" for name in names).encode("utf-8", "surrogateescape")
    fields = git(root, "check-attr", "-z", "--stdin", "text", stdin=stdin) \
        .decode("utf-8", "surrogateescape").split("\0")
    return {fields[index]: fields[index + 2] for index in range(0, len(fields) - 2, 3)}


def load_allowlist(content: bytes | None, findings: list[str]) -> dict[str, str]:
    """Map each allowlisted path to its SHA-256; report a malformed allowlist."""
    if content is None:
        return {}
    try:
        data: Any = json.loads(content.decode("utf-8"))
    except (UnicodeDecodeError, ValueError) as error:
        findings.append(f"{ALLOWLIST}: not valid UTF-8 JSON ({error})")
        return {}
    files = data.get("files") if isinstance(data, dict) else None
    if not isinstance(data, dict) or data.get("schema_version") != 1 \
            or not isinstance(files, dict):
        findings.append(f"{ALLOWLIST}: needs schema_version 1 and a files object")
        return {}
    allowed: dict[str, str] = {}
    for name, entry in files.items():
        digest = entry.get("sha256") if isinstance(entry, dict) else None
        reason = entry.get("reason") if isinstance(entry, dict) else None
        if not isinstance(digest, str) or not SHA256.fullmatch(digest) \
                or not isinstance(reason, str) or not reason.strip():
            findings.append(f"{ALLOWLIST}: {name} needs a sha256 of 64 lowercase hex digits "
                            "and a non-empty reason")
            continue
        allowed[name] = digest
    return allowed


def approved_image(name: str, content: bytes) -> bool:
    match = APPROVED_IMAGE.fullmatch(name)
    return match is not None and content.startswith(SIGNATURES[match.group(1)])


def check(root: Path) -> tuple[list[str], list[Binary]]:
    """Return the findings and every binary file that passed."""
    findings: list[str] = []
    blobs = tracked_blobs(root)
    names = sorted(blobs)
    attributes = text_attributes(root, names)
    contents = dict(zip(names, blob_contents(root, [blobs[name] for name in names])))
    # The allowlist is read from the index too, like every file it approves.
    allowed = load_allowlist(contents.get(ALLOWLIST), findings)
    passed: list[Binary] = []
    for name in names:
        content = contents[name]
        attribute = attributes.get(name, "unspecified")
        declared = attribute == "unset"
        has_nul = b"\0" in content
        digest = hashlib.sha256(content).hexdigest()
        if has_nul and not declared:
            findings.append(
                f"{name}: contains a NUL byte but .gitattributes treats it as text "
                f"(text: {attribute}); add a binary rule for it to .gitattributes, "
                "then approve it by path or in the allowlist")
            continue
        # An allowlisted file is checked against its digest whatever it holds now,
        # so a replacement never passes by no longer looking binary.
        if name in allowed:
            if digest == allowed[name]:
                passed.append(Binary(name, digest, f"allowlisted SHA-256 in {ALLOWLIST}"))
            else:
                findings.append(
                    f"{name}: SHA-256 {digest} differs from the reviewed {allowed[name]} "
                    f"in {ALLOWLIST}; review the new content, then set its sha256 to "
                    f"{digest} and update its reason in the same pull request")
            continue
        if not declared and not has_nul:
            continue
        if approved_image(name, content):
            passed.append(Binary(name, digest, "approved image path assets/*.png|jpg|jpeg"))
            continue
        hint = " (approved path, but its bytes are not that image format)" \
            if APPROVED_IMAGE.fullmatch(name) else ""
        findings.append(
            f"{name}: binary file outside the approved paths{hint}; remove it, or review it "
            f'and add "{name}": {{"sha256": "{digest}", "reason": "<why it is tracked>"}} '
            f"to the files of {ALLOWLIST}")
    for name in sorted(set(allowed) - set(blobs)):
        findings.append(f"{ALLOWLIST}: {name} is not a tracked file; remove its entry")
    return findings, passed


def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=(__doc__ or "").splitlines()[0])
    parser.add_argument("--root", type=Path, default=ROOT)
    args = parser.parse_args(argv)
    findings, passed = check(args.root.resolve())
    for binary in passed:
        print(f"ok: {binary.path} (sha256 {binary.sha256}; {binary.allowed_by})")
    for finding in findings:
        print(f"FAIL: {finding}")
    status = "FAIL" if findings else "PASS"
    print(f"binaries: {status} ({len(passed)} allowed binary file(s), "
          f"{len(findings)} finding(s))")
    return 1 if findings else 0


if __name__ == "__main__":
    sys.exit(main())
