#!/usr/bin/env python3
"""Build the release source bundle and release notes for one tagged version.

Used by .github/workflows/release.yml; runnable locally for a rehearsal. Every
input is read from Git at an explicit revision, never from the working tree, so
checkout settings cannot change the result and the same revision always yields
byte-identical output with the same Python and zlib.

Bundle  dist/VBA-Probability-Distributions-<version>.zip
        <prefix>/src/*.bas       production modules, CRLF (the VBE import form)
        <prefix>/tests/*.bas     regression harness, CRLF
        <prefix>/<documents>     LICENSE, README, INSTALLATION, CHANGELOG,
                                 SECURITY and VERSION, as committed
        <prefix>/MANIFEST.sha256 SHA-256 of every member above
Sums    dist/SHA256SUMS          SHA-256 of the bundle
Notes   the CHANGELOG section for the version, followed by an asset table

Entries are sorted, timestamps fixed at 1980-01-01 and permissions fixed, so
the bundle depends only on the committed bytes.

    python3 .github/scripts/build_release_bundle.py --version 1.0.0 \\
        --ref v1.0.0 --out dist --notes dist/release-notes.md
    python3 .github/scripts/build_release_bundle.py --self-test
"""
import argparse
import hashlib
import io
import os
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile

NAME = "VBA-Probability-Distributions"
DOCUMENTS = ("CHANGELOG.md", "INSTALLATION.md", "LICENSE", "README.md", "SECURITY.md",
             "VERSION")
MODULE_DIRS = ("src", "tests")
FIXED_TIME = (1980, 1, 1, 0, 0, 0)
SEMVER = re.compile(r"(0|[1-9]\d*)\.(0|[1-9]\d*)\.(0|[1-9]\d*)")
HEADING = re.compile(r"^## \[([^\]]+)\] - (\d{4}-\d{2}-\d{2})\s*$")


class BundleError(Exception):
    pass


def git(root, *args):
    proc = subprocess.run(["git", *args], cwd=root, capture_output=True, check=False)
    if proc.returncode:
        raise BundleError(proc.stderr.decode("utf-8", "replace").strip()
                          or f"git {' '.join(args)} failed")
    return proc.stdout


def blob(root, ref, path):
    try:
        return git(root, "show", f"{ref}:{path}")
    except BundleError:
        raise BundleError(f"{path} is missing at {ref}")


def module_paths(root, ref):
    paths = []
    for directory in MODULE_DIRS:
        listing = git(root, "ls-tree", "--name-only", f"{ref}:{directory}").decode()
        paths += [f"{directory}/{name}" for name in listing.split() if name.endswith(".bas")]
    if not any(p.startswith("src/") for p in paths):
        raise BundleError(f"no production modules under src/ at {ref}")
    return sorted(paths)


def crlf(raw, path):
    if b"\r" in raw:
        raise BundleError(f"{path} contains CR at the committed revision; expected LF")
    return raw.replace(b"\n", b"\r\n")


def sha256(raw):
    return hashlib.sha256(raw).hexdigest()


def members(root, ref, version):
    committed = blob(root, ref, "VERSION").decode("utf-8").strip()
    if committed != version:
        raise BundleError(f"VERSION at {ref} is {committed!r}, not {version!r}")
    files = {path: crlf(blob(root, ref, path), path) for path in module_paths(root, ref)}
    for document in DOCUMENTS:
        files[document] = blob(root, ref, document)
    manifest = "".join(f"{sha256(files[p])}  {p}\n" for p in sorted(files)).encode()
    files["MANIFEST.sha256"] = manifest
    return files


def write_zip(files, prefix):
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w") as archive:
        for path in sorted(files):
            info = zipfile.ZipInfo(f"{prefix}/{path}", date_time=FIXED_TIME)
            info.compress_type = zipfile.ZIP_DEFLATED
            info.create_system = 3
            info.external_attr = 0o100644 << 16
            archive.writestr(info, files[path], compresslevel=9)
    return buffer.getvalue()


def changelog_section(changelog, version):
    lines = changelog.splitlines()
    start = None
    for i, line in enumerate(lines):
        match = HEADING.match(line)
        if match and match.group(1) == version:
            start = i
            break
    if start is None:
        raise BundleError(f"CHANGELOG.md has no dated '## [{version}] - YYYY-MM-DD' section")
    end = next((j for j in range(start + 1, len(lines)) if lines[j].startswith("## ")),
               len(lines))
    body = "\n".join(lines[start + 1:end]).strip()
    if not body:
        raise BundleError(f"CHANGELOG.md section {version} is empty")
    return HEADING.match(lines[start]).group(2), body


def release_notes(root, ref, version, assets, repository):
    changelog = blob(root, ref, "CHANGELOG.md").decode("utf-8")
    released, body = changelog_section(changelog, version)
    base = f"https://github.com/{repository}/blob/v{version}"
    table = ["| Asset | Bytes | SHA-256 |", "| --- | ---: | --- |"]
    table += [f"| `{name}` | {len(raw):,} | `{sha256(raw)}` |" for name, raw in assets]
    return "\n".join([
        f"Released {released}. From [CHANGELOG.md]({base}/CHANGELOG.md):",
        "",
        body,
        "",
        "## Assets",
        "",
        *table,
        "",
        "The bundle and SHA256SUMS carry a GitHub build-provenance attestation. Verify "
        f"with `sha256sum --check SHA256SUMS` and `gh attestation verify <asset> --repo "
        f"{repository}`. Installation: [INSTALLATION.md]({base}/INSTALLATION.md). "
        f"Report vulnerabilities privately as described in [SECURITY.md]({base}/SECURITY.md).",
        "",
    ])


def build(root, ref, version, out, notes_path=None, repository=None):
    if not SEMVER.fullmatch(version):
        raise BundleError(f"{version!r} is not a MAJOR.MINOR.PATCH version")
    prefix = f"{NAME}-{version}"
    bundle = write_zip(members(root, ref, version), prefix)
    os.makedirs(out, exist_ok=True)
    zip_name = f"{prefix}.zip"
    with open(os.path.join(out, zip_name), "wb") as f:
        f.write(bundle)
    with open(os.path.join(out, "SHA256SUMS"), "w", encoding="utf-8", newline="\n") as f:
        f.write(f"{sha256(bundle)}  {zip_name}\n")
    if notes_path:
        if not repository or not re.fullmatch(r"[\w.-]+/[\w.-]+", repository):
            raise BundleError("release notes need --repo owner/name")
        notes = release_notes(root, ref, version, [(zip_name, bundle)], repository)
        with open(notes_path, "w", encoding="utf-8", newline="\n") as f:
            f.write(notes)
    return zip_name, bundle


def self_test():
    failures = []

    def check(condition, message):
        if not condition:
            failures.append(message)

    tmp = tempfile.mkdtemp()
    try:
        def put(path, text):
            full = os.path.join(tmp, path)
            os.makedirs(os.path.dirname(full) or tmp, exist_ok=True)
            with open(full, "w", encoding="utf-8", newline="\n") as f:
                f.write(text)

        git(tmp, "init", "-q")
        put("src/A.bas", 'Attribute VB_Name = "A"\nPublic Function K_STATS_A() As Double\nEnd Function\n')
        put("src/notes.txt", "not a module\n")
        put("tests/T.bas", 'Attribute VB_Name = "T"\n')
        for document in DOCUMENTS:
            put(document, "1.2.3\n" if document == "VERSION" else f"{document}\n")
        put("CHANGELOG.md", "# Changelog\n\n## [Unreleased]\n\n## [1.2.3] - 2026-10-12\n\n"
            "### Added\n\n- First release.\n\n## [1.2.2] - 2026-09-01\n\n- Older.\n")
        put("benchmark/big.csv", "excluded\n")
        git(tmp, "add", "-A")
        git(tmp, "-c", "user.name=t", "-c", "user.email=t@t", "commit", "-qm", "fixture")

        out1, out2 = os.path.join(tmp, "d1"), os.path.join(tmp, "d2")
        notes = os.path.join(tmp, "notes.md")
        name, first = build(tmp, "HEAD", "1.2.3", out1, notes, "owner/repo")
        _, second = build(tmp, "HEAD", "1.2.3", out2)
        check(name == f"{NAME}-1.2.3.zip", "bundle name carries the version")
        check(first == second, "two builds of one revision are byte-identical")
        with zipfile.ZipFile(io.BytesIO(first)) as archive:
            names = archive.namelist()
            check(names == sorted(names), "entries are sorted")
            expected = sorted(f"{NAME}-1.2.3/{p}" for p in
                              ("src/A.bas", "tests/T.bas", "MANIFEST.sha256", *DOCUMENTS))
            check(names == expected, f"bundle holds exactly the declared members: {names}")
            module = archive.read(f"{NAME}-1.2.3/src/A.bas")
            check(module.count(b"\r\n") == 3 and b"\n" not in module.replace(b"\r\n", b""),
                  ".bas members are CRLF")
            check(archive.read(f"{NAME}-1.2.3/README.md") == b"README.md\n",
                  "documents are kept as committed")
            check(all(i.date_time == FIXED_TIME for i in archive.infolist()),
                  "timestamps are fixed")
            manifest = archive.read(f"{NAME}-1.2.3/MANIFEST.sha256").decode()
            check(f"{sha256(module)}  src/A.bas" in manifest, "manifest hashes the CRLF module")
        sums = open(os.path.join(out1, "SHA256SUMS"), encoding="utf-8").read()
        check(sums == f"{sha256(first)}  {name}\n", "SHA256SUMS names the bundle digest")
        text = open(notes, encoding="utf-8").read()
        check("First release." in text and "Older." not in text,
              "notes hold exactly the version's changelog section")
        check(sha256(first) in text and "Released 2026-10-12" in text,
              "notes carry the asset digest and the release date")
        check("https://github.com/owner/repo/blob/v1.2.3/CHANGELOG.md" in text
              and "--repo owner/repo" in text and "](CHANGELOG.md)" not in text,
              "notes link absolutely at the tag and name the repository")

        for label, args in (("version differs from VERSION", ("HEAD", "1.2.4")),
                            ("non-SemVer version", ("HEAD", "1.2")),
                            ("unknown revision", ("no-such-ref", "1.2.3"))):
            try:
                build(tmp, args[0], args[1], os.path.join(tmp, "bad"))
                failures.append(f"{label} must be refused")
            except BundleError:
                pass
        put("CHANGELOG.md", "# Changelog\n\n## [Unreleased]\n\n- Pending.\n")
        put("tests/T.bas", 'Attribute VB_Name = "T"\r\n')
        git(tmp, "add", "-A")
        git(tmp, "-c", "core.autocrlf=false", "-c", "user.name=t", "-c", "user.email=t@t",
            "commit", "-qm", "undated")
        for label, call in (
                ("an undated release",
                 lambda: release_notes(tmp, "HEAD", "1.2.3", [], "owner/repo")),
                ("a module committed with CR", lambda: members(tmp, "HEAD", "1.2.3"))):
            try:
                call()
                failures.append(f"{label} must be refused")
            except BundleError:
                pass
    finally:
        shutil.rmtree(tmp, ignore_errors=True)

    if failures:
        print("SELF-TEST FAIL: release bundle builder")
        for failure in failures:
            print("  - " + failure)
        return 1
    print("SELF-TEST PASS: release bundle builder (byte-identical rebuild, exact members, "
          "CRLF modules, fixed timestamps, member manifest, SHA256SUMS, changelog-section "
          "notes, and refusal of version mismatch, non-SemVer, unknown revision, undated "
          "release and CR in a committed module)")
    return 0


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__.split("\n")[0])
    ap.add_argument("--self-test", action="store_true")
    ap.add_argument("--version", help="MAJOR.MINOR.PATCH, equal to VERSION at --ref")
    ap.add_argument("--ref", default="HEAD", help="Git revision to package")
    ap.add_argument("--out", default="dist")
    ap.add_argument("--notes", help="also write release notes to this path")
    ap.add_argument("--repo", default=os.environ.get("GITHUB_REPOSITORY"),
                    help="owner/name for links in the notes (default: $GITHUB_REPOSITORY)")
    args = ap.parse_args(argv)
    if args.self_test:
        return self_test()
    if not args.version:
        ap.error("--version is required")
    root = git(os.getcwd(), "rev-parse", "--show-toplevel").decode().strip()
    try:
        name, bundle = build(root, args.ref, args.version, args.out, args.notes, args.repo)
    except BundleError as exc:
        print(f"FAIL: release bundle\n  - {exc}")
        return 1
    print(f"PASS: {name} ({len(bundle):,} bytes) sha256 {sha256(bundle)}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
