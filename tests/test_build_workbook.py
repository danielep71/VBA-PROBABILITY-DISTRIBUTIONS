"""Tests for the offline workbook build: tools/devbuild/vba_project.py,
build_workbook.py and verify_workbook.py.

The compression is checked against the [MS-OVBA] compression examples; the
compound file against a raw parse of its header, FAT and directory; the dir,
PROJECT and PROJECTwm streams record by record. Builds run in temporary Git
repositories, plus one build of this repository's own sources, which also
proves that a rebuild is byte-identical and that no tracked file changes.
"""

from __future__ import annotations

import contextlib
import hashlib
import io
import json
import random
import shutil
import struct
import subprocess
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path
from typing import Any, ClassVar

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "tools" / "devbuild"))

import build_workbook as B  # noqa: E402
import vba_project as V  # noqa: E402
import verify_workbook as W  # noqa: E402

PROJECT_ID = "{0B9A3C1D-1111-4222-8333-444455556666}"


def digest(data: bytes) -> str:
    return "sha256:" + hashlib.sha256(data).hexdigest()


def standard(name: str, body: str = "") -> V.Module:
    return V.Module(name, f'Attribute VB_Name = "{name}"\r\nOption Explicit\r\n{body}')


def sample_modules() -> list[V.Module]:
    large = "".join(f"' line {i}: some repeated text, repeated text, repeated text\r\n"
                    for i in range(400))
    return [V.document_module("ThisWorkbook", True), V.document_module("shBuild", False),
            standard("M_STATS_PROBDIST_SMALL"), standard("M_STATS_PROBDIST_LARGE", large)]


# ============================================================================ compression
class CompressionTest(unittest.TestCase):
    # [MS-OVBA] section 3.2, compression examples: (uncompressed, compressed).
    NO_COMPRESSION = (b"abcdefghijklmnopqrstuv.",
                      bytes.fromhex("0119B000616263646566676800696A6B6C6D6E6F7000"
                                    "717273747576" "2E"))
    NORMAL = (b"#aaabcdefaaaaghijaaaaaklaaamnopqaaaaaaaaaaaarstuvwxyzaaa",
              bytes.fromhex("012FB000236161616263646582660070616768696A013808616B6C00306D6E6F70"
                            "06710270041072737475761077" "78797A00" "3C"))
    MAXIMUM = (b"a" * 73, bytes.fromhex("0103B002614500"))

    def test_specification_examples_decompress(self) -> None:
        for name, (plain, packed) in {"no compression": self.NO_COMPRESSION,
                                      "normal": self.NORMAL,
                                      "maximum": self.MAXIMUM}.items():
            with self.subTest(example=name):
                self.assertEqual(V.decompress(packed), plain)

    def test_specification_examples_compress_to_their_bytes(self) -> None:
        self.assertEqual(V.compress(self.NO_COMPRESSION[0]), self.NO_COMPRESSION[1])
        self.assertEqual(V.compress(self.MAXIMUM[0]), self.MAXIMUM[1])

    def test_normal_example_takes_the_nearest_of_equal_matches(self) -> None:
        # The specification's matching takes the nearest of several equally long
        # matches; the published normal example takes a farther one in some
        # copy tokens. Both decode to the same text. These bytes pin the
        # builder's choice, so that a change of matcher, which would change every
        # built workbook, fails here first.
        normal = V.compress(self.NORMAL[0])
        self.assertEqual(normal, bytes.fromhex(
            "012FB000236161616263646582660070616768696A013808616B6C00206D6E6F70"
            "0671027004007273747576107778797A002C"))
        self.assertEqual(len(normal), len(self.NORMAL[1]))
        self.assertEqual(V.decompress(normal), self.NORMAL[0])

    def test_round_trip(self) -> None:
        rng = random.Random(8)
        samples = [b"", b"a", b"abc", b"abc" * 3000, bytes(rng.randrange(256) for _ in range(9000)),
                   b"Option Explicit\r\n" * 700, bytes(range(256)) * 40, b"x" * 4096,
                   b"x" * 4097, (ROOT / "src/M_STATS_PROBDIST_CORE.bas").read_bytes()]
        for data in samples:
            with self.subTest(length=len(data)):
                packed = V.compress(data)
                self.assertEqual(packed[0], 0x01)
                self.assertEqual(V.decompress(packed), data)

    def test_chunks_follow_the_container_layout(self) -> None:
        # Every chunk but the last decompresses to exactly 4096 bytes; a 4096-byte
        # chunk that does not compress is stored raw (flag bit clear).
        rng = random.Random(3)
        noise = bytes(rng.randrange(256) for _ in range(4096))
        packed = V.compress(noise + b"tail")
        header = struct.unpack_from("<H", packed, 1)[0]
        self.assertEqual(header, 0x3FFF)
        self.assertEqual(packed[3:3 + 4096], noise)
        self.assertEqual(V.decompress(packed), noise + b"tail")

    def test_incompressible_last_chunk_is_refused(self) -> None:
        rng = random.Random(97)
        noise = "".join(chr(rng.randrange(33, 127)) for _ in range(4000))
        module = standard("M_STATS_PROBDIST_NOISE", f"' {noise}\r\n")
        with self.assertRaisesRegex(ValueError, "module M_STATS_PROBDIST_NOISE: the last"):
            V.build_vba_project([module], PROJECT_ID)

    def test_malformed_containers_are_refused(self) -> None:
        for name, data in {"no signature": b"", "wrong signature": b"\x02\x00\xb0",
                           "bad chunk signature": b"\x01\x00\x00\x00",
                           "truncated header": b"\x01\x00",
                           "token before its chunk": b"\x01\x02\xb0\x01\x00\x00"}.items():
            with self.subTest(case=name), self.assertRaises(ValueError):
                V.decompress(data)


class EncryptionTest(unittest.TestCase):
    def test_project_protection_fields_round_trip(self) -> None:
        rng = random.Random("{ID}")
        for data in (b"\x00", b"\xff", struct.pack("<I", 0), b"abcdefgh"):
            self.assertEqual(V.decrypt(V._encrypt(data, "{ID}", rng)), data)


# ============================================================================ compound file
def raw_header(blob: bytes) -> dict[str, Any]:
    """The header fields of [MS-CFB] 2.2, unpacked independently of the reader."""
    fields = struct.unpack_from("<8s16sHHHHH6sIIIIIIIII", blob, 0)
    names = ("signature", "clsid", "minor", "major", "byte_order", "sector_shift",
             "mini_shift", "reserved", "n_dir", "n_fat", "dir_start", "transaction",
             "cutoff", "minifat_start", "n_minifat", "difat_start", "n_difat")
    header = dict(zip(names, fields))
    header["difat"] = struct.unpack_from("<109I", blob, 76)
    return header


class CompoundFileTest(unittest.TestCase):
    blob = V.build_vba_project(sample_modules(), PROJECT_ID)

    def test_header_is_version_3_with_512_byte_sectors(self) -> None:
        header = raw_header(self.blob)
        self.assertEqual(header["signature"], V.CFB_SIGNATURE)
        self.assertEqual(header["clsid"], b"\x00" * 16)
        self.assertEqual((header["minor"], header["major"], header["byte_order"]),
                         (0x3E, 3, 0xFFFE))
        self.assertEqual((header["sector_shift"], header["mini_shift"]), (9, 6))
        self.assertEqual((header["n_dir"], header["transaction"], header["cutoff"]),
                         (0, 0, 4096))
        self.assertEqual((header["difat_start"], header["n_difat"]), (V.ENDOFCHAIN, 0))
        self.assertEqual(len(self.blob) % 512, 0)
        n_fat = header["n_fat"]
        self.assertTrue(all(s == V.FREESECT for s in header["difat"][n_fat:]))

    def test_every_sector_is_used_once(self) -> None:
        header = raw_header(self.blob)
        sectors = len(self.blob) // 512 - 1
        fat_bytes = b"".join(self.blob[512 * (s + 1):512 * (s + 2)]
                             for s in header["difat"][:header["n_fat"]])
        fat = struct.unpack(f"<{len(fat_bytes) // 4}I", fat_bytes)
        self.assertEqual([fat[s] for s in header["difat"][:header["n_fat"]]],
                         [V.FATSECT] * header["n_fat"])
        self.assertTrue(all(value == V.FREESECT for value in fat[sectors:]))
        cfb = V.CompoundFile(self.blob)
        starts = [header["dir_start"], header["minifat_start"], cfb.entries[0]["start"]]
        starts += [e["start"] for e in cfb.entries
                   if e["kind"] == 2 and e["size"] >= header["cutoff"]]
        used: list[int] = []
        for start in starts:
            used += V.CompoundFile._chain(start, list(fat))
        used += list(header["difat"][:header["n_fat"]])
        self.assertEqual(sorted(used), list(range(sectors)))

    def test_directory_tree_holds_the_project_streams(self) -> None:
        cfb = V.CompoundFile(self.blob)
        self.assertEqual(cfb.entries[0]["name"], "Root Entry")
        self.assertEqual(sorted(cfb.children(0)), ["PROJECT", "PROJECTwm", "VBA"])
        vba = cfb.children(cfb.children(0)["VBA"])
        self.assertEqual(sorted(vba), sorted(["_VBA_PROJECT", "dir",
                                              *(m.name for m in sample_modules())]))
        for sid in (0, cfb.children(0)["VBA"]):
            # Each storage's children form a binary search tree in [MS-CFB] order.
            def in_order(node: int) -> list[str]:
                if node == V.NOSTREAM:
                    return []
                entry = cfb.entries[node]
                return [*in_order(entry["left"]), entry["name"], *in_order(entry["right"])]
            names = in_order(cfb.entries[sid]["child"])
            self.assertEqual(names, sorted(names, key=lambda n: (len(n), n.upper())))

    def test_every_module_reads_back_unchanged(self) -> None:
        cfb = V.CompoundFile(self.blob)
        for module in sample_modules():
            with self.subTest(module=module.name):
                self.assertEqual(V.decompress(cfb.read(f"VBA/{module.name}")).decode("cp1252"),
                                 module.source)

    def test_vba_project_stream_asks_for_compilation_from_source(self) -> None:
        cfb = V.CompoundFile(self.blob)
        self.assertEqual(cfb.read("VBA/_VBA_PROJECT"), bytes.fromhex("CC61FFFF000000"))
        names = [entry["name"] for entry in cfb.entries]
        self.assertFalse([name for name in names if name.startswith("__SRP_")])

    def test_unknown_paths_and_long_names_are_refused(self) -> None:
        cfb = V.CompoundFile(self.blob)
        with self.assertRaises(KeyError):
            cfb.read("VBA/M_NONE")
        with self.assertRaises(KeyError):
            cfb.read("VBA")
        with self.assertRaisesRegex(ValueError, "longer than 31"):
            V.write_cfb({"X" * 32: b""})
        with self.assertRaisesRegex(ValueError, "1 to 31 characters"):
            V.Module("M" * 32, "")
        with self.assertRaisesRegex(ValueError, "used twice"):
            V.build_vba_project([standard("M_A"), standard("m_a")], PROJECT_ID)
        with self.assertRaisesRegex(ValueError, "not a compound file"):
            V.CompoundFile(b"PK" + self.blob[2:])

    def test_output_is_deterministic_for_a_project_id(self) -> None:
        self.assertEqual(self.blob, V.build_vba_project(sample_modules(), PROJECT_ID))
        other = "{11111111-0000-0000-0000-000000000000}"
        self.assertNotEqual(self.blob, V.build_vba_project(sample_modules(), other))


# ============================================================================ project streams
def dir_records(data: bytes) -> list[tuple[int, bytes]]:
    """[MS-OVBA] 2.3.4.2 records as (Id, payload). PROJECTVERSION (0x0009) is the
    one record whose size field is a constant 4 followed by six bytes."""
    records, pos = [], 0
    while pos < len(data):
        rid, size = struct.unpack_from("<HI", data, pos)
        pos += 6
        if rid == 0x0009:
            size = 6
        records.append((rid, data[pos:pos + size]))
        pos += size
    return records


class ProjectStreamsTest(unittest.TestCase):
    cfb = V.CompoundFile(V.build_vba_project(sample_modules(), PROJECT_ID))

    def test_dir_stream_records(self) -> None:
        records = dir_records(V.decompress(self.cfb.read("VBA/dir")))
        ids = [rid for rid, _ in records]
        self.assertEqual(ids[:13], [0x0001, 0x0002, 0x0014, 0x0003, 0x0004, 0x0005, 0x0040,
                                    0x0006, 0x003D, 0x0007, 0x0008, 0x0009, 0x000C])
        values = dict(records)
        self.assertEqual(struct.unpack("<I", values[0x0001])[0], 1)          # Win32
        self.assertEqual(struct.unpack("<H", values[0x0003])[0], 1252)       # code page
        self.assertEqual(values[0x0004], b"VBAProject")
        self.assertEqual(struct.unpack("<IH", values[0x0009]), (1, 0))
        self.assertEqual(struct.unpack("<H", values[0x000F])[0], len(sample_modules()))
        self.assertEqual(records[-1], (0x0010, b""))
        modules: list[dict[int, bytes]] = []
        for rid, payload in records[ids.index(0x0013) + 1:-1]:
            if rid == 0x0019:
                modules.append({})
            modules[-1][rid] = payload
        self.assertEqual(len(modules), len(sample_modules()))
        for module, fields in zip(sample_modules(), modules):
            with self.subTest(module=module.name):
                self.assertEqual(fields[0x0019], module.name.encode("cp1252"))
                self.assertEqual(fields[0x0047], module.name.encode("utf-16-le"))
                self.assertEqual(fields[0x001A], module.name.encode("cp1252"))
                self.assertEqual(struct.unpack("<I", fields[0x0031])[0], 0)  # no p-code
                kind = 0x0021 if module.kind == V.STANDARD else 0x0022
                self.assertIn(kind, fields)
                self.assertIn(0x002B, fields)

    def test_project_stream_declares_every_module(self) -> None:
        text = self.cfb.read("PROJECT").decode("cp1252")
        self.assertTrue(text.startswith(f'ID="{PROJECT_ID}"\r\n'))
        lines = text.split("\r\n")
        self.assertEqual([line for line in lines if line.startswith(("Module=", "Document="))],
                         ["Document=ThisWorkbook/&H00000000", "Document=shBuild/&H00000000",
                          "Module=M_STATS_PROBDIST_SMALL", "Module=M_STATS_PROBDIST_LARGE"])
        fields = {line.split("=", 1)[0]: line.split("=", 1)[1].strip('"') for line in lines
                  if line.startswith(("CMG=", "DPB=", "GC="))}
        self.assertEqual(V.decrypt(fields["CMG"]), struct.pack("<I", 0))  # no protection
        self.assertEqual(V.decrypt(fields["DPB"]), b"\x00")                # no password
        self.assertEqual(V.decrypt(fields["GC"]), b"\xff")                 # visible
        workspace = lines[lines.index("[Workspace]") + 1:-1]
        self.assertEqual(workspace, [f"{m.name}=0, 0, 0, 0, C" for m in sample_modules()])

    def test_projectwm_pairs_every_name(self) -> None:
        data = self.cfb.read("PROJECTwm")
        expected = b"".join(m.name.encode("cp1252") + b"\x00" + m.name.encode("utf-16-le")
                            + b"\x00\x00" for m in sample_modules()) + b"\x00\x00"
        self.assertEqual(data, expected)


# ============================================================================ builds
FIXTURE_POLICY = {"schema_version": 1, "regression_sources": ["src/M_STATS_PROBDIST_A.bas",
                                                             "tests/M_STATS_PROBDIST_TEST.bas"]}
FIXTURE_FILES: dict[str, bytes] = {
    ".gitattributes": b"*.bas text eol=crlf\n",
    B.POLICY: json.dumps(FIXTURE_POLICY, indent=2).encode() + b"\n",
    "src/M_STATS_PROBDIST_A.bas": (b'Attribute VB_Name = "M_STATS_PROBDIST_A"\r\nOption Explicit\r\n'
                                   b"Public Function K_STATS_A() As Double\r\n    K_STATS_A = 1\r\n"
                                   b"End Function\r\n"),
    "tests/M_STATS_PROBDIST_TEST.bas": (b'Attribute VB_Name = "M_STATS_PROBDIST_TEST"\r\n'
                                        b"Option Explicit\r\nPublic Sub Test_All()\r\n"
                                        b"    Debug.Print K_STATS_A()\r\nEnd Sub\r\n"),
}


def run_git(root: Path, *args: str) -> str:
    return subprocess.run(["git", "-C", str(root), *args], capture_output=True, text=True,
                          check=True).stdout


@contextlib.contextmanager
def git_tree(files: dict[str, bytes] | None = None, commit: bool = True) -> Any:
    """A committed Git repository holding the fixture files, changed by ``files``
    (a None value removes a file)."""
    with tempfile.TemporaryDirectory(prefix="workbook-build-") as temporary:
        root = Path(temporary) / "repo"
        root.mkdir()
        run_git(root, "init", "-q")
        for key, value in (("user.name", "Fixture"), ("user.email", "fixture@example.invalid"),
                           ("core.autocrlf", "false"), ("commit.gpgsign", "false")):
            run_git(root, "config", key, value)
        merged: dict[str, bytes | None] = {**FIXTURE_FILES, **(files or {})}
        for name, data in merged.items():
            if data is not None:
                (root / name).parent.mkdir(parents=True, exist_ok=True)
                (root / name).write_bytes(data)
        if commit:
            run_git(root, "add", "-A")
            run_git(root, "commit", "-qm", "Fixture")
        yield root


@unittest.skipIf(shutil.which("git") is None, "git is not installed")
class BuildTest(unittest.TestCase):
    def build(self, root: Path) -> tuple[Path, Path, B.Build]:
        out = root.parent / "out"
        result = B.build(root, out / "w.xlsm")
        return out / "w.xlsm", out / "w.manifest.json", result

    def test_a_clean_build_is_verified(self) -> None:
        with git_tree() as root:
            workbook, manifest, result = self.build(root)
            verdict = W.verify(root, manifest)
            self.assertEqual((verdict.status, verdict.findings), (W.VERIFIED, []))
            commit = run_git(root, "rev-parse", "HEAD").strip()
            self.assertEqual(result.commit, commit)
            self.assertTrue(result.clean)
            self.assertEqual(W.verify(root, manifest, commit="HEAD").status, W.VERIFIED)
            with zipfile.ZipFile(workbook) as package:
                self.assertEqual(W.read_stamp(package, B.STAMP_COMMIT), commit)
                self.assertEqual(W.read_stamp(package, B.STAMP_TIME), result.time)

    def test_manifest_digests(self) -> None:
        with git_tree() as root:
            workbook, manifest, _ = self.build(root)
            data = json.loads(manifest.read_text(encoding="utf-8"))
            self.assertEqual(data["format"], B.FORMAT)
            self.assertEqual(data["workbook"], {"file": "w.xlsm",
                                                "sha256": digest(workbook.read_bytes())})
            self.assertEqual([item["path"] for item in data["modules"]],
                             FIXTURE_POLICY["regression_sources"])
            for item in data["modules"]:
                with self.subTest(module=item["name"]):
                    blob = subprocess.run(["git", "-C", str(root), "cat-file", "blob",
                                           f"HEAD:{item['path']}"], capture_output=True,
                                          check=True).stdout
                    self.assertNotIn(b"\r\n", blob)  # stored with LF, as in this repository
                    self.assertEqual(item["git_blob_sha256"], digest(blob))
                    self.assertEqual(item["module_sha256"],
                                     digest(FIXTURE_FILES[item["path"]]))

    def test_a_rebuild_is_byte_identical(self) -> None:
        with git_tree() as root:
            workbook, manifest, _ = self.build(root)
            first = (workbook.read_bytes(), manifest.read_bytes())
            workbook.unlink()
            manifest.unlink()
            self.build(root)
            self.assertEqual((workbook.read_bytes(), manifest.read_bytes()), first)

    def test_package_is_stored_in_a_fixed_order_and_time(self) -> None:
        with git_tree() as root:
            workbook, _, _ = self.build(root)
            with zipfile.ZipFile(workbook) as package:
                infos = package.infolist()
                self.assertEqual([info.filename for info in infos][-1], "xl/vbaProject.bin")
                for info in infos:
                    self.assertEqual((info.date_time, info.compress_type, info.create_system),
                                     (B.ZIP_TIME, zipfile.ZIP_STORED, 0))
                self.assertIn("macroEnabled", package.read("[Content_Types].xml").decode())
                book = package.read("xl/workbook.xml").decode()
                self.assertIn('codeName="ThisWorkbook"', book)
                self.assertIn('<definedName name="build_commit">Build!$B$3</definedName>', book)

    def test_a_dirty_tree_is_stamped_and_unverified(self) -> None:
        with git_tree() as root:
            path = root / "src/M_STATS_PROBDIST_A.bas"
            path.write_bytes(path.read_bytes() + b"' changed\r\n")
            workbook, manifest, result = self.build(root)
            self.assertFalse(result.clean)
            with zipfile.ZipFile(workbook) as package:
                self.assertEqual(W.read_stamp(package, B.STAMP_COMMIT),
                                 result.commit + B.NOT_EVIDENCE)
            verdict = W.verify(root, manifest)
            self.assertEqual((verdict.status, verdict.findings), (W.UNVERIFIED, []))
            self.assertTrue(any("working-tree change" in note for note in verdict.notes))
            with contextlib.redirect_stdout(io.StringIO()):
                self.assertEqual(W.main([str(manifest), "--root", str(root)]), 0)
                self.assertEqual(W.main([str(manifest), "--root", str(root),
                                         "--require-verified"]), 1)

    def test_missing_and_extra_modules_stop_the_build(self) -> None:
        cases = {
            "not found": {"tests/M_STATS_PROBDIST_TEST.bas": None},
            "not listed": {"src/M_STATS_PROBDIST_EXTRA.bas":
                           b'Attribute VB_Name = "M_STATS_PROBDIST_EXTRA"\r\n'},
            "must be Attribute VB_Name": {"src/M_STATS_PROBDIST_A.bas":
                                          b'Attribute VB_Name = "M_OTHER"\r\n'},
            "directly under": {B.POLICY: json.dumps(
                {"regression_sources": ["benchmark/M_X.bas"]}).encode()},
            "lists a module twice": {B.POLICY: json.dumps(
                {"regression_sources": ["src/M_STATS_PROBDIST_A.bas",
                                        "src/M_STATS_PROBDIST_A.bas"]}).encode()},
            "non-empty list": {B.POLICY: json.dumps({"regression_sources": []}).encode()},
            "not cp1252 text": {"src/M_STATS_PROBDIST_A.bas":
                                b'Attribute VB_Name = "M_STATS_PROBDIST_A"\r\n\x81\r\n'},
        }
        for expected, files in cases.items():
            with self.subTest(case=expected), git_tree(files) as root:
                with self.assertRaisesRegex(B.BuildError, expected):
                    self.build(root)
                with contextlib.redirect_stdout(io.StringIO()) as printed:
                    code = B.main([str(root.parent / "o.xlsm"), "--root", str(root)])
                self.assertEqual(code, 1)
                self.assertIn("FAIL:", printed.getvalue())
                self.assertFalse((root.parent / "o.xlsm").exists())

    def test_main_reports_the_build(self) -> None:
        with git_tree() as root, contextlib.redirect_stdout(io.StringIO()) as printed:
            self.assertEqual(B.main([str(root.parent / "o.xlsm"), "--root", str(root)]), 0)
            self.assertIn("modules: M_STATS_PROBDIST_A, M_STATS_PROBDIST_TEST", printed.getvalue())
            self.assertTrue((root.parent / "o.manifest.json").is_file())


@unittest.skipIf(shutil.which("git") is None, "git is not installed")
class VerifyTest(unittest.TestCase):
    def setUp(self) -> None:
        context = git_tree()
        self.root = context.__enter__()
        self.addCleanup(context.__exit__, None, None, None)
        self.workbook = self.root.parent / "out/w.xlsm"
        self.manifest = self.root.parent / "out/w.manifest.json"
        B.build(self.root, self.workbook)
        self.data = json.loads(self.manifest.read_text(encoding="utf-8"))

    def rewrite(self, data: dict[str, Any]) -> None:
        self.manifest.write_text(json.dumps(data), encoding="utf-8")

    def assert_fails(self, expected: str) -> None:
        verdict = W.verify(self.root, self.manifest)
        self.assertEqual(verdict.status, W.FAIL)
        self.assertTrue(any(expected in finding for finding in verdict.findings),
                        verdict.findings)
        with contextlib.redirect_stdout(io.StringIO()) as printed:
            self.assertEqual(W.main([str(self.manifest), "--root", str(self.root)]), 1)
        self.assertIn("Binding: FAIL", printed.getvalue())

    def test_nothing_to_check_is_unverified(self) -> None:
        self.assertEqual(W.verify(self.root, self.root / "none.json").status, W.UNVERIFIED)
        self.workbook.unlink()
        verdict = W.verify(self.root, self.manifest)
        self.assertEqual((verdict.status, verdict.findings), (W.UNVERIFIED, []))
        self.assertIn("no workbook", verdict.notes[0])

    def test_an_unknown_commit_is_unverified(self) -> None:
        # Verified from a repository without the commit, the sources cannot be checked.
        with git_tree({"README.md": b"another history\n"}) as other:
            verdict = W.verify(other, self.manifest)
        self.assertEqual((verdict.status, verdict.findings), (W.UNVERIFIED, []))
        self.assertTrue(any("sources are not checked" in note for note in verdict.notes))

    def test_a_changed_workbook_fails(self) -> None:
        self.workbook.write_bytes(self.workbook.read_bytes() + b"\x00")
        self.assert_fails("sha256")

    def test_a_malformed_manifest_fails(self) -> None:
        self.manifest.write_text("{}", encoding="utf-8")
        self.assert_fails("not a valid manifest")
        self.rewrite({**self.data, "format": "other"})
        self.assert_fails("not a valid manifest")

    def test_a_wrong_module_digest_fails(self) -> None:
        modules = [dict(item) for item in self.data["modules"]]
        modules[0]["module_sha256"] = "sha256:" + "0" * 64
        self.rewrite({**self.data, "modules": modules})
        self.assert_fails("module M_STATS_PROBDIST_A has sha256")

    def test_a_wrong_blob_digest_fails(self) -> None:
        modules = [dict(item) for item in self.data["modules"]]
        modules[1]["git_blob_sha256"] = "sha256:" + "0" * 64
        self.rewrite({**self.data, "modules": modules})
        self.assert_fails("Git blob sha256")

    def test_a_missing_module_fails(self) -> None:
        self.rewrite({**self.data, "modules": self.data["modules"][:1]})
        self.assert_fails("missing from the workbook")

    def test_an_extra_module_fails(self) -> None:
        extra = {**self.data["modules"][0], "name": "M_STATS_PROBDIST_B",
                 "path": "src/M_STATS_PROBDIST_B.bas"}
        self.rewrite({**self.data, "modules": [*self.data["modules"], extra]})
        self.assert_fails("not in regression_sources")

    def test_another_commit_fails(self) -> None:
        (self.root / "README.md").write_text("next\n", encoding="utf-8")
        run_git(self.root, "add", "-A")
        run_git(self.root, "commit", "-qm", "Next")
        verdict = W.verify(self.root, self.manifest, commit="HEAD")
        self.assertEqual(verdict.status, W.FAIL)
        self.assertIn("expected", verdict.findings[0])
        self.assertEqual(W.verify(self.root, self.manifest, commit="HEAD~1").status, W.VERIFIED)

    def test_other_document_modules_fail(self) -> None:
        vba_project = {**self.data["vba_project"], "document_modules": ["ThisWorkbook", "shOther"]}
        self.rewrite({**self.data, "vba_project": vba_project})
        self.assert_fails("the VBA project declares")

    def test_a_reordered_module_list_fails(self) -> None:
        self.rewrite({**self.data, "modules": self.data["modules"][::-1]})
        self.assert_fails("module order")


@unittest.skipIf(shutil.which("git") is None, "git is not installed")
class RepositoryBuildTest(unittest.TestCase):
    """One build of this repository's own sources (about three seconds each)."""

    built: ClassVar[list[tuple[bytes, bytes]]] = []
    verdict: ClassVar[W.Verdict]
    result: ClassVar[B.Build]

    @classmethod
    def setUpClass(cls) -> None:
        listed = B.load_policy(ROOT)
        cls.before = {path: (ROOT / path).read_bytes() for path in listed}
        with tempfile.TemporaryDirectory(prefix="workbook-build-") as temporary:
            for run in ("a", "b"):
                target = Path(temporary) / run / "w.xlsm"
                cls.result = B.build(ROOT, target)
                cls.built.append((target.read_bytes(), cls.result.manifest.read_bytes()))
            cls.verdict = W.verify(ROOT, Path(temporary) / "a" / "w.manifest.json")

    def test_the_build_is_byte_identical_and_bound(self) -> None:
        self.assertEqual(self.built[0], self.built[1])
        self.assertEqual(self.verdict.findings, [])
        self.assertEqual(self.verdict.status,
                         W.VERIFIED if self.result.clean else W.UNVERIFIED)

    def test_the_policy_modules_are_embedded_and_unchanged(self) -> None:
        listed = B.load_policy(ROOT)
        self.assertEqual([source.path for source in self.result.sources], listed)
        for path, data in self.before.items():
            self.assertEqual((ROOT / path).read_bytes(), data, path)
        with zipfile.ZipFile(io.BytesIO(self.built[0][0])) as package:
            cfb = V.CompoundFile(package.read("xl/vbaProject.bin"))
        for path in listed:
            with self.subTest(module=path):
                text = V.decompress(cfb.read(f"VBA/{Path(path).stem}"))
                self.assertEqual(text, B.module_text((ROOT / path).read_bytes()).encode("cp1252"))


if __name__ == "__main__":
    unittest.main()
