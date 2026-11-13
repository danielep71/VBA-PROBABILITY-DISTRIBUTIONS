"""
Carried certification: prove that a VBA module differs from the bytes Excel
executed in comments and whitespace only (#301).

The Excel record and the observation manifests bind the exact bytes of every
.bas module. A change to a comment therefore made the evidence stale and cost
a full Excel wave, although Excel would execute the same code. Decision of
2026-11-13 (docs/EXCEL_CERTIFICATION.md, "Carried certification"): a module
whose code is identical to the executed one keeps its certification. The
binding is then CARRIED from the executed commit, and says so wherever it is
reported. A release is never certified on a carried binding.

"Identical code" is decided by the comment-free statement stream of
tools/vba_lexer.py, the lexer the VBA source checks already use:

  * comments are removed: from an apostrophe outside a string and outside
    [...], or from Rem at the start of a statement, including a comment
    continued with " _";
  * continuation lines are joined into their logical line;
  * outside string literals and [...] expressions, runs of whitespace become
    one space and leading and trailing whitespace is dropped; inside them
    every character is kept;
  * blank lines are dropped; conditional-compilation directives are kept and
    marked as such.

Whitespace is never inserted or removed between two tokens, so `a=1` and
`a = 1` are different code. Anything else that differs is a code change. The
comparison fails closed: a non-ASCII byte, an unreadable file, a missing
commit or a recorded digest that does not match the executed bytes all refuse
the carry.

    python vba_code_identity.py --from COMMIT [PATH ...]
        Compare each tracked .bas (or the given paths) with COMMIT and report
        IDENTICAL, CARRIED (comments or whitespace only) or CODE CHANGED.
        Exit 1 if any module's code changed.
"""
import argparse
import hashlib
import os
import subprocess
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, os.path.join(ROOT, "tools"))
import vba_lexer  # noqa: E402


class CarryRefused(Exception):
    """The module cannot be shown to hold the executed code."""


def _normalize(code):
    """Collapse whitespace outside string literals and [...] expressions."""
    out = []
    in_string = False
    brackets = 0
    pending_space = False
    index = 0
    while index < len(code):
        char = code[index]
        if in_string:
            out.append(char)
            if char == '"':
                if code[index + 1:index + 2] == '"':
                    out.append('"')
                    index += 1
                else:
                    in_string = False
        elif brackets:
            out.append(char)
            if char == "[":
                brackets += 1
            elif char == "]":
                brackets -= 1
        elif char.isspace():
            pending_space = True
        else:
            if pending_space and out:
                out.append(" ")
            pending_space = False
            out.append(char)
            if char == '"':
                in_string = True
            elif char == "[":
                brackets = 1
        index += 1
    return "".join(out)


def decode(raw):
    """Module text with LF line ends; refuses anything but ASCII."""
    try:
        text = raw.decode("ascii")
    except UnicodeDecodeError as exc:
        raise CarryRefused(f"non-ASCII byte at offset {exc.start}") from None
    return text.replace("\r\n", "\n").replace("\r", "\n")


def code_stream(raw):
    """The comment-free logical lines of a module, as (kind, code) pairs."""
    stream = []
    for _first, _last, kind, code in vba_lexer.logical_units(decode(raw).split("\n")):
        normal = _normalize(code)
        if normal:
            stream.append((kind, normal))
    return stream


def code_digest(raw):
    payload = "\n".join(f"{kind}\t{code}" for kind, code in code_stream(raw))
    return "sha256:" + hashlib.sha256(payload.encode("ascii")).hexdigest()


def first_difference(old_raw, new_raw):
    """None when the code streams are equal, else a short description."""
    old, new = code_stream(old_raw), code_stream(new_raw)
    for index, (a, b) in enumerate(zip(old, new)):
        if a != b:
            return f"statement {index + 1}: executed {a[1][:80]!r}, now {b[1][:80]!r}"
    if len(old) != len(new):
        longer, label = (old, "executed") if len(old) > len(new) else (new, "current")
        return (f"{abs(len(old) - len(new))} statement(s) only in the {label} module, "
                f"first {longer[min(len(old), len(new))][1][:80]!r}")
    return None


def normalized_hash_bytes(raw):
    data = raw.replace(b"\r\n", b"\n").replace(b"\r", b"\n")
    return "sha256:" + hashlib.sha256(data).hexdigest()


def git_show(root, commit, path):
    try:
        proc = subprocess.run(["git", "-C", root, "show", f"{commit}:{path}"],
                              capture_output=True, check=False)
    except OSError as exc:
        raise CarryRefused(f"git is not available: {exc}") from None
    if proc.returncode:
        raise CarryRefused(f"{path} cannot be read at commit {commit}")
    return proc.stdout


def carried(root, path, recorded, commit, digest=normalized_hash_bytes):
    """Prove that `path` in the tree holds the code executed at `commit`.

    `recorded` is the digest the record holds for the executed bytes, in the
    form `digest` produces. Returns the code digest shared by both versions.
    Raises CarryRefused with the reason otherwise.
    """
    if not commit:
        raise CarryRefused("the record names no executed commit")
    executed = git_show(root, commit, path)
    if digest(executed) != recorded:
        raise CarryRefused(f"the bytes at commit {commit} are not the recorded ones")
    try:
        with open(os.path.join(root, *path.split("/")), "rb") as f:
            current = f.read()
    except OSError as exc:
        raise CarryRefused(f"cannot read the module: {exc}") from None
    difference = first_difference(executed, current)
    if difference is not None:
        raise CarryRefused("code changed: " + difference)
    return code_digest(current)


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[1])
    ap.add_argument("--from", dest="commit", required=True)
    ap.add_argument("paths", nargs="*")
    a = ap.parse_args(argv)
    paths = a.paths
    if not paths:
        listed = subprocess.run(["git", "-C", ROOT, "ls-files", "*.bas"], capture_output=True,
                                text=True, check=True).stdout.split()
        paths = sorted(listed)
    changed = 0
    for path in paths:
        try:
            executed = git_show(ROOT, a.commit, path)
            with open(os.path.join(ROOT, *path.split("/")), "rb") as f:
                current = f.read()
            if normalized_hash_bytes(executed) == normalized_hash_bytes(current):
                continue
            difference = first_difference(executed, current)
        except (CarryRefused, OSError) as exc:
            difference = str(exc)
        if difference is None:
            print(f"CARRIED       {path}: comments or whitespace only")
        else:
            changed += 1
            print(f"CODE CHANGED  {path}: {difference}")
    if changed:
        print(f"FAIL: {changed} module(s) changed in code since {a.commit}; an Excel wave is needed")
        return 1
    print(f"PASS: every compared module holds the code of {a.commit}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
