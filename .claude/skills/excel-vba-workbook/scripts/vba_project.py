"""Write, and read back, an Office VBA project storage (vbaProject.bin).

Implements, from the [MS-OVBA] and [MS-CFB] specifications:
  - the VBA compression algorithm (2.4.1) and its decompression,
  - the dir stream (2.3.4.2) and the PROJECT / PROJECTwm streams (2.3.1),
  - the Data Encryption of the CMG / DPB / GC project fields (2.4.3),
  - a version-3 compound file writer (512-byte sectors, mini stream) and a
    minimal reader of the same layout.

_VBA_PROJECT is written with Version 0xFFFF and no performance cache (no
p-code), which tells the host to compile the project from source when it
opens it. Standard modules and the document modules of the workbook and its
sheets are supported; class modules and forms are not, because this
repository has none.

The output depends only on the modules and the project ID: there is no clock,
no random seed and no unordered iteration. tools/devbuild/verify_workbook.py
reads a built project back with decompress() and CompoundFile, and
tests/test_build_workbook.py checks both against the specification's
compression examples and against raw parses of the compound file. Standard
library only.
"""

from __future__ import annotations

import random
import struct
from dataclasses import dataclass
from typing import Any

ENDOFCHAIN = 0xFFFFFFFE
FREESECT = 0xFFFFFFFF
FATSECT = 0xFFFFFFFD
NOSTREAM = 0xFFFFFFFF
CFB_SIGNATURE = bytes.fromhex("D0CF11E0A1B11AE1")
SECTOR = 512
MINI_SECTOR = 64
MINI_CUTOFF = 4096
CHUNK = 4096

WORKBOOK_BASE = "0{00020819-0000-0000-C000-000000000046}"
WORKSHEET_BASE = "0{00020820-0000-0000-C000-000000000046}"
# _VBA_PROJECT: Reserved1 0x61CC, Version 0xFFFF (compile from source), Reserved2, Reserved3.
VBA_PROJECT_STREAM = bytes([0xCC, 0x61, 0xFF, 0xFF, 0x00, 0x00, 0x00])
VBA_ENCODING = "cp1252"
# VBA identifiers and compound-file entry names are both limited to 31 characters.
MAX_NAME = 31


# ============================================================================ compression
def _copy_token_params(diff: int) -> tuple[int, int]:
    """[MS-OVBA] 2.4.1.3.19.1 CopyToken Help: bit count of the offset, maximum length."""
    bit_count = 4
    while (1 << bit_count) < diff:
        bit_count += 1
    max_len = (0xFFFF >> bit_count) + 3
    return bit_count, max_len


def compress(data: bytes) -> bytes:
    """[MS-OVBA] 2.4.1.3.6: a CompressedContainer of ``data``."""
    out = bytearray([0x01])
    # Empty data is the signature byte alone: the algorithm's chunk loop never runs.
    for pos in range(0, len(data), CHUNK):
        out += _compress_chunk(data[pos:pos + CHUNK])
    return bytes(out)


# The MS-OVBA chunk compression algorithm as one loop, kept in the
# specification's shape so it can be checked against it step by step.
def _compress_chunk(chunk: bytes) -> bytes:  # noqa: C901
    body = bytearray()
    cur = 0
    n = len(chunk)
    index: dict[bytes, list[int]] = {}

    def add(i: int) -> None:
        if i + 3 <= n:
            index.setdefault(chunk[i:i + 3], []).append(i)

    while cur < n:
        flag_pos = len(body)
        body.append(0)
        flag = 0
        for bit in range(8):
            if cur >= n:
                break
            best_len, best_off = 0, 0
            if cur > 0:
                bit_count, max_len = _copy_token_params(cur)
                limit = min(max_len, n - cur)
                # The nearest candidates first; a longer match replaces a shorter one only.
                for cand in reversed(index.get(chunk[cur:cur + 3], [])[-256:]):
                    off = cur - cand
                    if off > (1 << bit_count):
                        break
                    length = 0
                    while length < limit and chunk[cand + length] == chunk[cur + length]:
                        length += 1
                    if length > best_len:
                        best_len, best_off = length, off
                        if length == limit:
                            break
            if best_len >= 3:
                bit_count, _ = _copy_token_params(cur)
                token = ((best_off - 1) << (16 - bit_count)) | (best_len - 3)
                body += struct.pack("<H", token)
                flag |= 1 << bit
                for i in range(cur, cur + best_len):
                    add(i)
                cur += best_len
            else:
                body.append(chunk[cur])
                add(cur)
                cur += 1
        body[flag_pos] = flag
    if len(body) + 2 > CHUNK + 2:
        if len(chunk) != CHUNK:
            # [MS-OVBA] allows a raw chunk only as 4096 bytes and a short chunk only
            # last; a short last chunk that will not compress cannot be represented
            # without padding the module source, so stop instead of writing it.
            raise ValueError(f"the last {len(chunk)} bytes do not compress into one chunk")
        return struct.pack("<H", (CHUNK - 1) | (0b011 << 12)) + chunk
    header = ((len(body) + 2 - 3) & 0x0FFF) | (0b011 << 12) | (1 << 15)
    return struct.pack("<H", header) + bytes(body)


def decompress(data: bytes) -> bytes:
    """[MS-OVBA] 2.4.1.3.1: the bytes of a CompressedContainer."""
    if not data or data[0] != 0x01:
        raise ValueError("compressed container: missing signature byte 0x01")
    out = bytearray()
    pos = 1
    while pos < len(data):
        if pos + 2 > len(data):
            raise ValueError("compressed container: truncated chunk header")
        start = pos
        header = struct.unpack_from("<H", data, pos)[0]
        pos += 2
        if (header >> 12) & 0x7 != 0b011:
            raise ValueError(f"compressed container: bad chunk signature at byte {start}")
        end = min(start + (header & 0x0FFF) + 3, len(data))
        if not header & 0x8000:
            out += data[pos:pos + CHUNK]
            pos += CHUNK
            continue
        chunk_start = len(out)
        while pos < end:
            flags = data[pos]
            pos += 1
            for bit in range(8):
                if pos >= end:
                    break
                if not (flags >> bit) & 1:
                    out.append(data[pos])
                    pos += 1
                    continue
                if pos + 2 > end:
                    raise ValueError("compressed container: truncated copy token")
                token = struct.unpack_from("<H", data, pos)[0]
                pos += 2
                bit_count, _ = _copy_token_params(len(out) - chunk_start)
                length = (token & (0xFFFF >> bit_count)) + 3
                offset = (token >> (16 - bit_count)) + 1
                source = len(out) - offset
                if source < chunk_start:
                    raise ValueError("compressed container: copy token before its chunk")
                for i in range(length):
                    out.append(out[source + i])
    return bytes(out)


# ============================================================================ data encryption
def _encrypt(data: bytes, project_id: str, rng: random.Random) -> str:
    """[MS-OVBA] 2.4.3.2 Encryption, with the seed and ignored bytes drawn from ``rng``."""
    seed = rng.randrange(256)
    version = 2
    proj_key = sum(project_id.encode(VBA_ENCODING)) & 0xFF
    out = bytearray([seed, seed ^ version, seed ^ proj_key])
    unenc1, enc1, enc2 = proj_key, seed ^ proj_key, seed ^ version

    def emit(byte: int) -> None:
        nonlocal unenc1, enc1, enc2
        b = byte ^ ((enc2 + unenc1) & 0xFF)
        out.append(b)
        enc2, enc1, unenc1 = enc1, b, byte

    for _ in range((seed & 6) // 2):
        emit(rng.randrange(256))
    for b in struct.pack("<I", len(data)):
        emit(b)
    for b in data:
        emit(b)
    return out.hex().upper()


def decrypt(hex_text: str) -> bytes:
    """[MS-OVBA] 2.4.3.3 Decryption."""
    raw = bytes.fromhex(hex_text)
    seed, version_enc, proj_key_enc = raw[0], raw[1], raw[2]
    if seed ^ version_enc != 2:
        raise ValueError("encrypted field: bad version")
    proj_key = seed ^ proj_key_enc
    unenc1, enc1, enc2 = proj_key, proj_key_enc, version_enc
    pos = 3
    plain = bytearray()
    for _ in range((seed & 6) // 2 + 4):
        b = raw[pos] ^ ((enc2 + unenc1) & 0xFF)
        enc2, enc1, unenc1 = enc1, raw[pos], b
        plain.append(b)
        pos += 1
    length = struct.unpack("<I", bytes(plain[-4:]))[0]
    data = bytearray()
    for _ in range(length):
        b = raw[pos] ^ ((enc2 + unenc1) & 0xFF)
        enc2, enc1, unenc1 = enc1, raw[pos], b
        data.append(b)
        pos += 1
    return bytes(data)


# ============================================================================ project streams
STANDARD, DOCUMENT = "standard", "document"


@dataclass(frozen=True)
class Module:
    name: str
    source: str            # full module text, CRLF line endings, including Attribute lines
    kind: str = STANDARD   # STANDARD or DOCUMENT

    def __post_init__(self) -> None:
        if self.kind not in (STANDARD, DOCUMENT):
            raise ValueError(f"module {self.name}: unknown kind {self.kind!r}")
        if not self.name or len(self.name) > MAX_NAME:
            raise ValueError(f"module {self.name!r}: a name of 1 to {MAX_NAME} characters "
                             "is required")


def document_module(name: str, workbook: bool) -> Module:
    """The empty document module of the workbook or of one worksheet."""
    base = WORKBOOK_BASE if workbook else WORKSHEET_BASE
    text = "\r\n".join([
        f'Attribute VB_Name = "{name}"',
        f'Attribute VB_Base = "{base}"',
        "Attribute VB_GlobalNameSpace = False",
        "Attribute VB_Creatable = False",
        "Attribute VB_PredeclaredId = True",
        "Attribute VB_Exposed = True",
        "Attribute VB_TemplateDerived = False",
        "Attribute VB_Customizable = True",
        "Option Explicit",
        "",
    ])
    return Module(name, text, DOCUMENT)


def _rec(rid: int, payload: bytes) -> bytes:
    return struct.pack("<HI", rid, len(payload)) + payload


def _u16(text: str) -> bytes:
    return text.encode("utf-16-le")


def _mb(text: str) -> bytes:
    return text.encode(VBA_ENCODING)


REFERENCES = [
    ("stdole", r"*\G{00020430-0000-0000-C000-000000000046}#2.0#0#C:\Windows\System32\stdole2.tlb#OLE Automation"),
    ("Office", r"*\G{2DF8D04C-5BFA-101B-BDE5-00AA0044DE52}#2.0#0#C:\Program Files\Common Files\Microsoft Shared\OFFICE16\MSO.DLL#Microsoft Office 16.0 Object Library"),
]


def build_dir(project_name: str, modules: list[Module]) -> bytes:
    """[MS-OVBA] 2.3.4.2: the uncompressed dir stream."""
    d = bytearray()
    d += _rec(0x0001, struct.pack("<I", 1))          # SYSKIND: Win32
    d += _rec(0x0002, struct.pack("<I", 0x0409))     # LCID
    d += _rec(0x0014, struct.pack("<I", 0x0409))     # LCIDINVOKE
    d += _rec(0x0003, struct.pack("<H", 1252))       # CODEPAGE
    d += _rec(0x0004, _mb(project_name))             # NAME
    d += _rec(0x0005, b"") + _rec(0x0040, b"")       # DOCSTRING (+unicode)
    d += _rec(0x0006, b"") + _rec(0x003D, b"")       # HELPFILEPATH 1 / 2
    d += _rec(0x0007, struct.pack("<I", 0))          # HELPCONTEXT
    d += _rec(0x0008, struct.pack("<I", 0))          # LIBFLAGS
    d += struct.pack("<HIIH", 0x0009, 4, 1, 0)       # VERSION: reserved=4, major, minor
    d += _rec(0x000C, b"") + _rec(0x003C, b"")       # CONSTANTS (+unicode)
    for ref_name, libid in REFERENCES:
        d += _rec(0x0016, _mb(ref_name)) + _rec(0x003E, _u16(ref_name))
        lib = _mb(libid)
        d += _rec(0x000D, struct.pack("<I", len(lib)) + lib + struct.pack("<IH", 0, 0))
    d += _rec(0x000F, struct.pack("<H", len(modules)))
    d += _rec(0x0013, struct.pack("<H", 0xFFFF))     # PROJECTCOOKIE
    for m in modules:
        d += _rec(0x0019, _mb(m.name)) + _rec(0x0047, _u16(m.name))
        d += _rec(0x001A, _mb(m.name)) + _rec(0x0032, _u16(m.name))
        d += _rec(0x001C, b"") + _rec(0x0048, b"")
        d += _rec(0x0031, struct.pack("<I", 0))      # MODULEOFFSET: source at 0, no p-code
        d += _rec(0x001E, struct.pack("<I", 0))      # HELPCONTEXT
        d += _rec(0x002C, struct.pack("<H", 0xFFFF)) # MODULECOOKIE
        d += _rec(0x0021 if m.kind == STANDARD else 0x0022, b"")  # MODULETYPE
        d += _rec(0x002B, b"")                       # module terminator
    d += _rec(0x0010, b"")                           # dir terminator
    return bytes(d)


def build_project_text(project_id: str, modules: list[Module]) -> bytes:
    """[MS-OVBA] 2.3.1: the PROJECT stream, deterministic for a project ID."""
    rng = random.Random(project_id)
    lines = [f'ID="{project_id}"']
    for m in modules:
        lines.append(f"Document={m.name}/&H00000000" if m.kind == DOCUMENT else f"Module={m.name}")
    lines += [
        'Name="VBAProject"',
        'HelpContextID="0"',
        'VersionCompatible32="393222000"',
        f'CMG="{_encrypt(struct.pack("<I", 0), project_id, rng)}"',
        f'DPB="{_encrypt(bytes([0]), project_id, rng)}"',
        f'GC="{_encrypt(bytes([0xFF]), project_id, rng)}"',
        "",
        "[Host Extender Info]",
        "&H00000001={3832D640-CF90-11CF-8E43-00A0C911005A};VBE;&H00000000",
        "",
        "[Workspace]",
    ]
    for m in modules:
        lines.append(f"{m.name}=0, 0, 0, 0, C")
    return ("\r\n".join(lines) + "\r\n").encode(VBA_ENCODING)


def build_projectwm(modules: list[Module]) -> bytes:
    """[MS-OVBA] 2.3.3: the PROJECTwm stream of module name pairs."""
    out = bytearray()
    for m in modules:
        out += _mb(m.name) + b"\x00" + _u16(m.name) + b"\x00\x00"
    out += b"\x00\x00"
    return bytes(out)


# ============================================================================ compound file
@dataclass
class _Entry:
    name: str
    kind: int                    # 1 storage, 2 stream, 5 root
    data: bytes = b""
    children: list["_Entry"] | None = None
    sid: int = 0
    left: int = NOSTREAM
    right: int = NOSTREAM
    child: int = NOSTREAM
    start: int = ENDOFCHAIN
    size: int = 0


def _cfb_key(e: _Entry) -> tuple[int, str]:
    """[MS-CFB] 2.6.4: siblings compare by name length, then by upper-case name."""
    return (len(e.name), e.name.upper())


# Writes the MS-CFB layout in one pass (directory, FAT, mini stream); its
# branches are the format's cases, which a split would only scatter.
def write_cfb(tree: dict[str, Any]) -> bytes:  # noqa: C901
    """tree: {name: bytes | dict}; returns the compound file bytes. Every
    timestamp and CLSID is zero, so equal trees give equal bytes."""
    root = _Entry("Root Entry", 5, children=[])
    flat: list[_Entry] = [root]

    def build(parent: _Entry, node: dict[str, Any]) -> None:
        assert parent.children is not None  # build() only receives storage entries
        for name, value in node.items():
            if len(name) > MAX_NAME:
                raise ValueError(f"compound file entry {name!r} is longer than {MAX_NAME} "
                                 "characters")
            if isinstance(value, dict):
                e = _Entry(name, 1, children=[])
                parent.children.append(e)
                flat.append(e)
                build(e, value)
            else:
                e = _Entry(name, 2, data=value, size=len(value))
                parent.children.append(e)
                flat.append(e)

    build(root, tree)
    for i, e in enumerate(flat):
        e.sid = i

    def link(entries: list[_Entry]) -> int:
        if not entries:
            return NOSTREAM
        entries = sorted(entries, key=_cfb_key)
        mid = len(entries) // 2
        node = entries[mid]
        node.left = link(entries[:mid])
        node.right = link(entries[mid + 1:])
        return node.sid

    for e in flat:
        if e.children is not None:
            e.child = link(e.children)

    fat: list[int] = []
    sectors: list[bytes] = []

    def alloc(data: bytes) -> int:
        count = max(1, -(-len(data) // SECTOR))
        first = len(sectors)
        for i in range(count):
            sectors.append(data[i * SECTOR:(i + 1) * SECTOR].ljust(SECTOR, b"\x00"))
            fat.append(first + i + 1 if i < count - 1 else ENDOFCHAIN)
        return first

    # Mini stream for small streams.
    mini = bytearray()
    minifat: list[int] = []
    for e in flat:
        if e.kind == 2 and e.size < MINI_CUTOFF:
            if e.size == 0:
                e.start = ENDOFCHAIN
                continue
            count = -(-e.size // MINI_SECTOR)
            first = len(minifat)
            for i in range(count):
                minifat.append(first + i + 1 if i < count - 1 else ENDOFCHAIN)
            e.start = first
            mini += e.data.ljust(count * MINI_SECTOR, b"\x00")
    for e in flat:
        if e.kind == 2 and e.size >= MINI_CUTOFF:
            e.start = alloc(e.data)
    if mini:
        root.start = alloc(bytes(mini))
        root.size = len(mini)
    minifat_start, n_minifat = ENDOFCHAIN, 0
    if minifat:
        mf = b"".join(struct.pack("<I", v) for v in minifat)
        minifat_start = alloc(mf.ljust(-(-len(mf) // SECTOR) * SECTOR, b"\xff"))
        n_minifat = -(-len(mf) // SECTOR)

    # Directory.
    dir_bytes = bytearray()
    for e in flat:
        name = e.name.encode("utf-16-le") + b"\x00\x00"
        entry = name.ljust(64, b"\x00")
        entry += struct.pack("<HBB", len(name), e.kind, 1)
        entry += struct.pack("<III", e.left, e.right, e.child)
        entry += b"\x00" * 16 + struct.pack("<I", 0) + b"\x00" * 16
        entry += struct.pack("<IQ", e.start if e.kind != 1 else 0, e.size if e.kind != 1 else 0)
        dir_bytes += entry
    while len(dir_bytes) % SECTOR:
        dir_bytes += b"\x00" * 64 + struct.pack("<HBB", 0, 0, 0) \
            + struct.pack("<III", NOSTREAM, NOSTREAM, NOSTREAM) \
            + b"\x00" * 36 + struct.pack("<IQ", 0, 0)
    dir_start = alloc(bytes(dir_bytes))

    # FAT sectors (they describe themselves too).
    n_fat = 1
    while n_fat * (SECTOR // 4) < len(sectors) + n_fat:
        n_fat += 1
    if n_fat > 109:
        raise ValueError("the project needs a DIFAT sector, which is not implemented")
    fat_start = len(sectors)
    fat += [FATSECT] * n_fat
    fat += [FREESECT] * (n_fat * (SECTOR // 4) - len(fat))
    fat_bytes = b"".join(struct.pack("<I", v) for v in fat)
    for i in range(n_fat):
        sectors.append(fat_bytes[i * SECTOR:(i + 1) * SECTOR])

    header = bytearray()
    header += CFB_SIGNATURE + b"\x00" * 16
    header += struct.pack("<HHHHH", 0x003E, 0x0003, 0xFFFE, 9, 6) + b"\x00" * 6
    header += struct.pack("<IIIII", 0, n_fat, dir_start, 0, MINI_CUTOFF)
    header += struct.pack("<IIII", minifat_start, n_minifat, ENDOFCHAIN, 0)
    difat = [fat_start + i for i in range(n_fat)] + [FREESECT] * (109 - n_fat)
    header += b"".join(struct.pack("<I", v) for v in difat)
    assert len(header) == SECTOR
    return bytes(header) + b"".join(sectors)


class CompoundFile:
    """[MS-CFB] version-3 reader of what write_cfb() writes: FAT, directory
    tree and mini stream. It refuses a cyclic chain and an unknown path."""

    def __init__(self, blob: bytes) -> None:
        if blob[:8] != CFB_SIGNATURE:
            raise ValueError("not a compound file")
        self.blob = blob
        self.sector = 1 << struct.unpack_from("<H", blob, 30)[0]
        self.mini = 1 << struct.unpack_from("<H", blob, 32)[0]
        n_fat, dir_start = struct.unpack_from("<II", blob, 44)
        self.cutoff, minifat_start = struct.unpack_from("<II", blob, 56)
        if n_fat > 109:
            raise ValueError("compound file with a DIFAT sector")
        difat = struct.unpack_from("<109I", blob, 76)[:n_fat]
        fat_bytes = b"".join(self._sector(s) for s in difat)
        self.fat = list(struct.unpack(f"<{len(fat_bytes) // 4}I", fat_bytes))
        directory = self._chain_bytes(dir_start)
        self.entries = [self._entry(directory[i:i + 128]) for i in range(0, len(directory), 128)]
        root = self.entries[0]
        self.mini_stream = self._chain_bytes(root["start"])[:root["size"]] \
            if root["start"] != ENDOFCHAIN else b""
        minifat_bytes = self._chain_bytes(minifat_start) if minifat_start != ENDOFCHAIN else b""
        self.minifat = list(struct.unpack(f"<{len(minifat_bytes) // 4}I", minifat_bytes))

    def _sector(self, index: int) -> bytes:
        offset = self.sector * (index + 1)
        if offset + self.sector > len(self.blob):
            raise ValueError(f"compound file: sector {index} is past the end")
        return self.blob[offset:offset + self.sector]

    @staticmethod
    def _chain(start: int, table: list[int]) -> list[int]:
        chain: list[int] = []
        seen: set[int] = set()
        while start != ENDOFCHAIN:
            if start in seen or start >= len(table):
                raise ValueError("compound file: broken sector chain")
            seen.add(start)
            chain.append(start)
            start = table[start]
        return chain

    def _chain_bytes(self, start: int) -> bytes:
        return b"".join(self._sector(s) for s in self._chain(start, self.fat))

    @staticmethod
    def _entry(raw: bytes) -> dict[str, Any]:
        name_length = struct.unpack_from("<H", raw, 64)[0]
        return {"name": raw[:max(0, name_length - 2)].decode("utf-16-le"), "kind": raw[66],
                "left": struct.unpack_from("<I", raw, 68)[0],
                "right": struct.unpack_from("<I", raw, 72)[0],
                "child": struct.unpack_from("<I", raw, 76)[0],
                "start": struct.unpack_from("<I", raw, 116)[0],
                "size": struct.unpack_from("<Q", raw, 120)[0]}

    def children(self, sid: int) -> dict[str, int]:
        """The entries under storage ``sid``, by name."""
        found: dict[str, int] = {}
        stack = [self.entries[sid]["child"]]
        while stack:
            node = stack.pop()
            if node == NOSTREAM:
                continue
            if node in found.values() or node >= len(self.entries):
                raise ValueError("compound file: broken directory tree")
            entry = self.entries[node]
            found[entry["name"]] = node
            stack += [entry["left"], entry["right"]]
        return found

    def read(self, path: str) -> bytes:
        """The bytes of the stream at ``path``, such as ``VBA/dir``."""
        sid = 0
        for part in path.split("/"):
            children = self.children(sid)
            if part not in children:
                raise KeyError(f"compound file: no entry {path}")
            sid = children[part]
        entry = self.entries[sid]
        if entry["kind"] != 2:
            raise KeyError(f"compound file: {path} is not a stream")
        if entry["size"] >= self.cutoff:
            return self._chain_bytes(entry["start"])[:entry["size"]]
        if entry["size"] == 0:
            return b""
        chunks = [self.mini_stream[s * self.mini:(s + 1) * self.mini]
                  for s in self._chain(entry["start"], self.minifat)]
        return b"".join(chunks)[:entry["size"]]


# ============================================================================ top level
def build_vba_project(modules: list[Module], project_id: str) -> bytes:
    """The vbaProject.bin of ``modules``, in their order, for a fixed project ID."""
    names = [m.name.casefold() for m in modules]
    if len(set(names)) != len(names):
        raise ValueError("a module name is used twice")
    vba: dict[str, bytes] = {"_VBA_PROJECT": VBA_PROJECT_STREAM,
                             "dir": compress(build_dir("VBAProject", modules))}
    for m in modules:
        try:
            vba[m.name] = compress(m.source.encode(VBA_ENCODING))
        except ValueError as error:
            raise ValueError(f"module {m.name}: {error}") from error
    return write_cfb({"PROJECT": build_project_text(project_id, modules),
                      "PROJECTwm": build_projectwm(modules),
                      "VBA": vba})
