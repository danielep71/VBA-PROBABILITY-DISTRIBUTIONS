"""Tests for tools/check_binaries.py, in temporary Git repositories."""

from __future__ import annotations

import contextlib
import hashlib
import io
import json
import shutil
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "tools"))

import check_binaries  # noqa: E402

PNG = b"\x89PNG\r\n\x1a\n\0\0\0\rIHDR"
JPEG = b"\xff\xd8\xff\xe0\0\x10JFIF\0"
WORKBOOK = b"PK\x03\x04\x14\0\x06\0[Content_Types].xml\0vbaProject.bin"
WORKBOOK_PATH = "examples/STATS-Distributions demo.xlsm"
ALLOWLIST = check_binaries.ALLOWLIST


def allowlist(files: dict[str, bytes], reason: str = "Reviewed demo workbook") -> object:
    return {"schema_version": 1,
            "files": {name: {"sha256": hashlib.sha256(content).hexdigest(), "reason": reason}
                      for name, content in files.items()}}


@unittest.skipIf(shutil.which("git") is None, "git is not installed")
class BinaryCheckTest(unittest.TestCase):
    def repository(self, files: dict[str, bytes], listed: object = None) -> list[str]:
        """Stage files in a new repository with this repository's .gitattributes; check it."""
        with tempfile.TemporaryDirectory(prefix="binaries-") as directory:
            root = Path(directory)
            subprocess.run(["git", "init", "-q"], cwd=root, check=True)
            subprocess.run(["git", "config", "core.autocrlf", "false"], cwd=root, check=True)
            files = {".gitattributes": (ROOT / ".gitattributes").read_bytes(), **files}
            if listed is not None:
                files[ALLOWLIST] = json.dumps(listed).encode()
            for name, content in files.items():
                (root / name).parent.mkdir(parents=True, exist_ok=True)
                (root / name).write_bytes(content)
            subprocess.run(["git", "add", "-f", "."], cwd=root, check=True)
            return check_binaries.check(root)[0]

    def assertFinding(self, findings: list[str], text: str) -> None:
        self.assertTrue(any(text in finding for finding in findings), findings)

    def test_repository_passes(self) -> None:
        findings, passed = check_binaries.check(ROOT)
        self.assertEqual(findings, [])
        self.assertIn(WORKBOOK_PATH, {binary.path for binary in passed})

    def test_text_files_pass(self) -> None:
        self.assertEqual(self.repository({
            "README.md": b"# Title\n", "data.csv": b"a,b\n1,2\n",
            "src/M_STATS_PROBDIST_Demo.bas": b'Attribute VB_Name = "M_STATS_PROBDIST_Demo"\r\n',
            "assets/diagram.svg": b"<svg/>\n"}), [])

    def test_approved_images_pass(self) -> None:
        self.assertEqual(self.repository({"assets/Home2.jpg": JPEG, "assets/banner2.png": PNG,
                                          "assets/photo.jpeg": JPEG}), [])

    def test_image_outside_the_approved_paths_fails(self) -> None:
        for name in ("docs/banner.png", "assets/nested/banner.png", "banner.png"):
            with self.subTest(name=name):
                self.assertFinding(self.repository({name: PNG}),
                                   f"{name}: binary file outside the approved paths")

    def test_approved_path_needs_the_image_signature(self) -> None:
        self.assertFinding(self.repository({"assets/banner.png": WORKBOOK}),
                           "its bytes are not that image format")
        self.assertFinding(self.repository({"assets/photo.jpg": PNG}),
                           "assets/photo.jpg: binary file outside the approved paths")
        self.assertFinding(self.repository({"assets/anim.gif": b"GIF89a\0"}),
                           "assets/anim.gif: binary file outside the approved paths")

    def test_digest_pinned_workbook_passes(self) -> None:
        self.assertEqual(self.repository({WORKBOOK_PATH: WORKBOOK},
                                         allowlist({WORKBOOK_PATH: WORKBOOK})), [])

    def test_changed_workbook_fails(self) -> None:
        changed = WORKBOOK + b"\0new macro"
        findings = self.repository({WORKBOOK_PATH: changed}, allowlist({WORKBOOK_PATH: WORKBOOK}))
        new_digest = hashlib.sha256(changed).hexdigest()
        self.assertFinding(findings, f"{WORKBOOK_PATH}: SHA-256 {new_digest} differs from "
                                     f"the reviewed {hashlib.sha256(WORKBOOK).hexdigest()}")
        self.assertFinding(findings, f"then set its sha256 to {new_digest}")

    def test_allowlisted_workbook_replaced_by_text_still_fails(self) -> None:
        self.assertFinding(self.repository({WORKBOOK_PATH: b"plain text now\n"},
                                           allowlist({WORKBOOK_PATH: WORKBOOK})),
                           f"{WORKBOOK_PATH}: SHA-256")

    def test_unlisted_binary_fails(self) -> None:
        for name in (WORKBOOK_PATH, "book.xlsx", "docs/report.pdf", "archive.zip",
                     "tools/helper.exe", "src/frmRun.frx"):
            with self.subTest(name=name):
                findings = self.repository({name: WORKBOOK})
                self.assertFinding(findings, f"{name}: binary file outside the approved paths")
                self.assertFinding(findings, f'"sha256": "{hashlib.sha256(WORKBOOK).hexdigest()}"')
        # Listing a different file does not admit this one.
        self.assertFinding(self.repository({WORKBOOK_PATH: WORKBOOK, "copy.xlsm": WORKBOOK},
                                           allowlist({WORKBOOK_PATH: WORKBOOK})),
                           "copy.xlsm: binary file outside the approved paths")

    def test_nul_file_without_binary_attributes_fails(self) -> None:
        for name, attribute in (("notes.txt", "set"), ("payload.bin", "auto"),
                                ("assets/banner.dat", "auto")):
            with self.subTest(name=name):
                content = b"text\0more"
                # Not even a matching allowlist entry admits it.
                findings = self.repository({name: content}, allowlist({name: content}))
                self.assertFinding(findings, f"{name}: contains a NUL byte but .gitattributes "
                                             f"treats it as text (text: {attribute})")

    def test_stale_allowlist_entry_fails(self) -> None:
        self.assertFinding(self.repository({}, allowlist({WORKBOOK_PATH: WORKBOOK})),
                           f"{ALLOWLIST}: {WORKBOOK_PATH} is not a tracked file")

    def test_unstaged_allowlist_is_not_read(self) -> None:
        with tempfile.TemporaryDirectory(prefix="binaries-") as directory:
            root = Path(directory)
            subprocess.run(["git", "init", "-q"], cwd=root, check=True)
            (root / ".gitattributes").write_bytes((ROOT / ".gitattributes").read_bytes())
            (root / "book.xlsm").write_bytes(WORKBOOK)
            subprocess.run(["git", "add", "."], cwd=root, check=True)
            (root / ".github").mkdir()
            (root / ALLOWLIST).write_text(json.dumps(allowlist({"book.xlsm": WORKBOOK})))
            self.assertFinding(check_binaries.check(root)[0], "book.xlsm: binary file")

    def test_unstaged_gitattributes_are_not_read(self) -> None:
        # Review finding: attributes must come from the index, like the contents.
        payload = b"data\0more"
        with tempfile.TemporaryDirectory(prefix="binaries-") as directory:
            root = Path(directory)
            subprocess.run(["git", "init", "-q"], cwd=root, check=True)
            (root / ".gitattributes").write_bytes(b"* text=auto\n")
            (root / "payload.bin").write_bytes(payload)
            (root / ".github").mkdir()
            (root / ALLOWLIST).write_text(json.dumps(allowlist({"payload.bin": payload})))
            subprocess.run(["git", "add", "."], cwd=root, check=True)
            # An unstaged rule that would make the staged file binary is ignored.
            (root / ".gitattributes").write_bytes(b"* text=auto\npayload.bin -text\n")
            self.assertFinding(check_binaries.check(root)[0],
                               "payload.bin: contains a NUL byte but .gitattributes treats it as text")

    def test_malformed_allowlist(self) -> None:
        digest = "a" * 64
        for listed, text in (([], "needs schema_version 1"),
                             ({"schema_version": 2, "files": {}}, "needs schema_version 1"),
                             ({"schema_version": 1, "files": []}, "needs schema_version 1"),
                             ({"schema_version": 1,
                               "files": {"a.xlsm": {"sha256": "x", "reason": "r"}}},
                              "a.xlsm needs a sha256 of 64 lowercase hex digits"),
                             ({"schema_version": 1,
                               "files": {"a.xlsm": {"sha256": digest.upper(), "reason": "r"}}},
                              "a.xlsm needs a sha256"),
                             ({"schema_version": 1,
                               "files": {"a.xlsm": {"sha256": digest, "reason": " "}}},
                              "a.xlsm needs a sha256 of 64 lowercase hex digits and a non-empty "
                              "reason")):
            with self.subTest(listed=listed):
                self.assertFinding(self.repository({}, listed), text)
        self.assertFinding(self.repository({ALLOWLIST: b"{not json"}), "not valid UTF-8 JSON")

    def test_main_reports_the_verdict(self) -> None:
        with contextlib.redirect_stdout(io.StringIO()) as printed:
            self.assertEqual(check_binaries.main(["--root", str(ROOT)]), 0)
        self.assertIn(f"ok: {WORKBOOK_PATH} (sha256 ", printed.getvalue())
        self.assertIn("binaries: PASS", printed.getvalue())


if __name__ == "__main__":
    unittest.main()
