"""Build a vbaProject.bin from VBA source files (stdlib only).

Implements just enough of:
  - [MS-OVBA] VBA project structure, compression and data encryption
  - [MS-CFB]  Compound File Binary format (writer)

The performance cache is omitted and _VBA_PROJECT.Version is 0xFFFF,
so Excel recompiles the project from source when the file is opened.
"""
from __future__ import annotations

import random
import struct
from dataclasses import dataclass

# ---------------------------------------------------------------------------
# MS-OVBA 2.4.1 compression
# ---------------------------------------------------------------------------


def _copy_token_bits(decompressed_current: int, chunk_start: int) -> tuple[int, int, int, int]:
    difference = decompressed_current - chunk_start
    bit_count = max((difference - 1).bit_length(), 4)
    length_mask = 0xFFFF >> bit_count
    offset_mask = ~length_mask & 0xFFFF
    maximum_length = (0xFFFF >> bit_count) + 3
    return length_mask, offset_mask, bit_count, maximum_length


def _compress_chunk(data: bytes) -> bytes:
    """Compress one chunk (<= 4096 bytes). Returns header + payload."""
    out = bytearray()
    pos = 0
    end = len(data)
    index: dict[bytes, list[int]] = {}
    while pos < end:
        flag_index = len(out)
        out.append(0)
        flags = 0
        for bit in range(8):
            if pos >= end:
                break
            _, _, bit_count, max_len = _copy_token_bits(pos, 0)
            best_len = 0
            best_off = 0
            # Look up earlier positions that start with the same 3 bytes.
            for cand in reversed(index.get(data[pos : pos + 3], [])[-64:]):
                length = 0
                while (
                    pos + length < end
                    and length < max_len
                    and data[cand + length] == data[pos + length]
                ):
                    length += 1
                if length > best_len:
                    best_len = length
                    best_off = pos - cand
                    if length == max_len:
                        break
            if best_len >= 3:
                token = ((best_off - 1) << (16 - bit_count)) | (best_len - 3)
                out += struct.pack("<H", token)
                flags |= 1 << bit
                step = best_len
            else:
                out.append(data[pos])
                step = 1
            for p in range(pos, pos + step):
                index.setdefault(data[p : p + 3], []).append(p)
            pos += step
        out[flag_index] = flags
    if len(out) > 4096:
        # Fall back to a raw chunk (only allowed for full 4096-byte chunks).
        assert len(data) == 4096
        header = (4096 - 1) | (0b011 << 12)
        return struct.pack("<H", header) + data
    header = (len(out) + 2 - 3) | (0b011 << 12) | 0x8000
    return struct.pack("<H", header) + bytes(out)


def compress(data: bytes) -> bytes:
    out = bytearray(b"\x01")
    for i in range(0, len(data), 4096):
        out += _compress_chunk(data[i : i + 4096])
    return bytes(out)


def decompress(data: bytes) -> bytes:
    """Inverse of compress(); used by tests."""
    assert data[0] == 1
    out = bytearray()
    pos = 1
    while pos < len(data):
        header = struct.unpack_from("<H", data, pos)[0]
        size = (header & 0x0FFF) + 3
        compressed = header & 0x8000
        chunk = data[pos + 2 : pos + size]
        pos += size
        chunk_start = len(out)
        if not compressed:
            out += chunk
            continue
        i = 0
        while i < len(chunk):
            flags = chunk[i]
            i += 1
            for bit in range(8):
                if i >= len(chunk):
                    break
                if flags & (1 << bit):
                    token = struct.unpack_from("<H", chunk, i)[0]
                    i += 2
                    length_mask, _, bit_count, _ = _copy_token_bits(len(out), chunk_start)
                    length = (token & length_mask) + 3
                    offset = (token >> (16 - bit_count)) + 1
                    for _ in range(length):
                        out.append(out[-offset])
                else:
                    out.append(chunk[i])
                    i += 1
    return bytes(out)


# ---------------------------------------------------------------------------
# MS-OVBA 2.4.3 data encryption (used for CMG / DPB / GC in PROJECT stream)
# ---------------------------------------------------------------------------


def _encrypt(data: bytes, project_id: str, rng: random.Random) -> str:
    seed = rng.randrange(256)
    version = 2
    proj_key = sum(project_id.encode("ascii")) & 0xFF
    out = bytearray([seed, seed ^ version, seed ^ proj_key])
    unencrypted_byte1 = proj_key
    encrypted_byte1 = out[2]
    encrypted_byte2 = out[1]

    def emit(byte: int) -> None:
        nonlocal unencrypted_byte1, encrypted_byte1, encrypted_byte2
        enc = byte ^ ((encrypted_byte2 + unencrypted_byte1) & 0xFF)
        out.append(enc)
        encrypted_byte2 = encrypted_byte1
        encrypted_byte1 = enc
        unencrypted_byte1 = byte

    for _ in range((seed & 6) // 2):
        emit(rng.randrange(256))
    for b in struct.pack("<I", len(data)):
        emit(b)
    for b in data:
        emit(b)
    return out.hex().upper()


# ---------------------------------------------------------------------------
# VBA project
# ---------------------------------------------------------------------------

DOC_BASE_WORKBOOK = "0{00020819-0000-0000-C000-000000000046}"
DOC_BASE_WORKSHEET = "0{00020820-0000-0000-C000-000000000046}"
CLASS_BASE = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"


@dataclass
class Module:
    name: str
    kind: str  # "standard" | "class" | "document"
    code: str  # body without Attribute header lines
    base: str = ""  # VB_Base for document modules

    def source(self) -> str:
        lines = [f'Attribute VB_Name = "{self.name}"']
        if self.kind == "class":
            lines += [
                f'Attribute VB_Base = "{CLASS_BASE}"',
                "Attribute VB_GlobalNameSpace = False",
                "Attribute VB_Creatable = False",
                "Attribute VB_PredeclaredId = False",
                "Attribute VB_Exposed = False",
                "Attribute VB_TemplateDerived = False",
                "Attribute VB_Customizable = False",
            ]
        elif self.kind == "document":
            lines += [
                f'Attribute VB_Base = "{self.base}"',
                "Attribute VB_GlobalNameSpace = False",
                "Attribute VB_Creatable = False",
                "Attribute VB_PredeclaredId = True",
                "Attribute VB_Exposed = True",
                "Attribute VB_TemplateDerived = False",
                "Attribute VB_Customizable = True",
            ]
        body = self.code.replace("\r\n", "\n").strip("\n")
        text = "\r\n".join(lines) + "\r\n" + (body.replace("\n", "\r\n") + "\r\n" if body else "")
        return text


def parse_source_file(name: str, text: str, kind: str) -> Module:
    """Turn an exported .bas / .cls file into a Module (drops header lines)."""
    body_lines = []
    in_begin = False
    for line in text.replace("\r\n", "\n").split("\n"):
        stripped = line.strip()
        if stripped.startswith("VERSION ") and stripped.endswith("CLASS"):
            continue
        if stripped == "BEGIN":
            in_begin = True
            continue
        if in_begin:
            if stripped == "END":
                in_begin = False
            continue
        if stripped.startswith("Attribute VB_"):
            continue
        body_lines.append(line)
    return Module(name=name, kind=kind, code="\n".join(body_lines))


def _rec(rec_id: int, payload: bytes) -> bytes:
    return struct.pack("<HI", rec_id, len(payload)) + payload


def _encode(s: str, codepage: int) -> bytes:
    return s.encode(f"cp{codepage}")


def _dir_stream(modules: list[Module], codepage: int) -> bytes:
    d = bytearray()
    d += _rec(0x0001, struct.pack("<I", 1))  # SYSKIND: Win32
    d += _rec(0x0002, struct.pack("<I", 0x0409))  # LCID
    d += _rec(0x0014, struct.pack("<I", 0x0409))  # LCIDINVOKE
    d += _rec(0x0003, struct.pack("<H", codepage))  # CODEPAGE
    d += _rec(0x0004, b"VBAProject")  # NAME
    d += _rec(0x0005, b"") + _rec(0x0040, b"")  # DOCSTRING
    d += _rec(0x0006, b"") + _rec(0x003D, b"")  # HELPFILEPATH
    d += _rec(0x0007, struct.pack("<I", 0))  # HELPCONTEXT
    d += _rec(0x0008, struct.pack("<I", 0))  # LIBFLAGS
    d += struct.pack("<HIIH", 0x0009, 4, 1, 0)  # VERSION
    d += _rec(0x000C, b"") + _rec(0x003C, b"")  # CONSTANTS
    # No references: Excel adds VBA / Excel / stdole itself on recompile.
    d += _rec(0x000F, struct.pack("<H", len(modules)))  # MODULES
    d += _rec(0x0013, struct.pack("<H", 0xFFFF))  # PROJECTCOOKIE
    for m in modules:
        name = _encode(m.name, codepage)
        uname = m.name.encode("utf-16-le")
        d += _rec(0x0019, name)
        d += _rec(0x0047, uname)
        d += _rec(0x001A, name) + _rec(0x0032, uname)
        d += _rec(0x001C, b"") + _rec(0x0048, b"")
        d += _rec(0x0031, struct.pack("<I", 0))  # MODULEOFFSET
        d += _rec(0x001E, struct.pack("<I", 0))
        d += _rec(0x002C, struct.pack("<H", 0xFFFF))
        d += struct.pack("<HI", 0x0021 if m.kind == "standard" else 0x0022, 0)
        d += struct.pack("<HI", 0x002B, 0)  # module terminator
    d += struct.pack("<HI", 0x0010, 0)  # dir terminator
    return bytes(d)


def _project_stream(modules: list[Module], codepage: int, rng: random.Random) -> bytes:
    project_id = "{00000000-0000-0000-0000-000000000000}"
    lines = [f'ID="{project_id}"']
    for m in modules:
        if m.kind == "document":
            lines.append(f"Document={m.name}/&H00000000")
        elif m.kind == "class":
            lines.append(f"Class={m.name}")
        else:
            lines.append(f"Module={m.name}")
    lines += [
        'Name="VBAProject"',
        'HelpContextID="0"',
        'VersionCompatible32="393222000"',
        'CMG="' + _encrypt(struct.pack("<I", 0), project_id, rng) + '"',
        'DPB="' + _encrypt(bytes([0x00]), project_id, rng) + '"',
        'GC="' + _encrypt(bytes([0xFF]), project_id, rng) + '"',
        "",
        "[Host Extender Info]",
        "&H00000001={3832D640-CF90-11CF-8E43-00A0C911005A};VBE;&H00000000",
        "",
        "[Workspace]",
    ]
    for m in modules:
        lines.append(f"{m.name}=0, 0, 0, 0, C")
    return _encode("\r\n".join(lines) + "\r\n", codepage)


def _projectwm_stream(modules: list[Module], codepage: int) -> bytes:
    out = bytearray()
    for m in modules:
        out += _encode(m.name, codepage) + b"\x00"
        out += m.name.encode("utf-16-le") + b"\x00\x00"
    out += b"\x00\x00"
    return bytes(out)


def build_vba_project(modules: list[Module], codepage: int = 1252, seed: int = 0) -> bytes:
    rng = random.Random(seed)
    vba_streams = {
        "_VBA_PROJECT": struct.pack("<HHBH", 0x61CC, 0xFFFF, 0, 0),
        "dir": compress(_dir_stream(modules, codepage)),
    }
    for m in modules:
        vba_streams[m.name] = compress(_encode(m.source(), codepage))
    tree = {
        "PROJECT": _project_stream(modules, codepage, rng),
        "PROJECTwm": _projectwm_stream(modules, codepage),
        "VBA": vba_streams,
    }
    return write_cfb(tree)


# ---------------------------------------------------------------------------
# MS-CFB writer (version 3, 512-byte sectors)
# ---------------------------------------------------------------------------

SECTOR = 512
MINI_SECTOR = 64
MINI_CUTOFF = 4096
FREESECT = 0xFFFFFFFF
ENDOFCHAIN = 0xFFFFFFFE
FATSECT = 0xFFFFFFFD
NOSTREAM = 0xFFFFFFFF


@dataclass
class _Entry:
    name: str
    type: int  # 1 storage, 2 stream, 5 root
    data: bytes = b""
    children: list[int] | None = None
    left: int = NOSTREAM
    right: int = NOSTREAM
    child: int = NOSTREAM
    start: int = ENDOFCHAIN
    size: int = 0


def _cfb_key(name: str) -> tuple[int, str]:
    return (len(name), name.upper())


def write_cfb(tree: dict) -> bytes:
    entries: list[_Entry] = [_Entry("Root Entry", 5, children=[])]

    def add(parent: int, node: dict) -> None:
        for name, value in node.items():
            idx = len(entries)
            if isinstance(value, dict):
                entries.append(_Entry(name, 1, children=[]))
                entries[parent].children.append(idx)  # type: ignore[union-attr]
                add(idx, value)
            else:
                entries.append(_Entry(name, 2, data=value))
                entries[parent].children.append(idx)  # type: ignore[union-attr]

    add(0, tree)

    # Balanced binary search tree for each storage's children (all black).
    def build_tree(ids: list[int]) -> int:
        if not ids:
            return NOSTREAM
        mid = len(ids) // 2
        root = ids[mid]
        entries[root].left = build_tree(ids[:mid])
        entries[root].right = build_tree(ids[mid + 1 :])
        return root

    for e in entries:
        if e.children is not None:
            ordered = sorted(e.children, key=lambda i: _cfb_key(entries[i].name))
            e.child = build_tree(ordered)

    # Mini stream for small streams.
    mini_stream = bytearray()
    mini_fat: list[int] = []
    big_streams: list[_Entry] = []
    for e in entries:
        if e.type != 2:
            continue
        e.size = len(e.data)
        if e.size == 0:
            e.start = ENDOFCHAIN
        elif e.size < MINI_CUTOFF:
            first = len(mini_fat)
            count = -(-e.size // MINI_SECTOR)
            for i in range(count):
                mini_fat.append(first + i + 1 if i < count - 1 else ENDOFCHAIN)
            e.start = first
            mini_stream += e.data.ljust(count * MINI_SECTOR, b"\x00")
        else:
            big_streams.append(e)

    # Lay out sectors: [directory][minifat][ministream][big streams][fat]
    sectors: list[bytes] = []
    fat: list[int] = []

    def alloc(data: bytes) -> int:
        if not data:
            return ENDOFCHAIN
        count = -(-len(data) // SECTOR)
        first = len(sectors)
        padded = data.ljust(count * SECTOR, b"\x00")
        for i in range(count):
            sectors.append(padded[i * SECTOR : (i + 1) * SECTOR])
            fat.append(first + i + 1 if i < count - 1 else ENDOFCHAIN)
        return first

    dir_count = -(-len(entries) // 4)
    dir_start = alloc(b"\x00" * dir_count * SECTOR)  # placeholder, filled later

    mini_fat_bytes = b"".join(struct.pack("<I", v) for v in mini_fat)
    if mini_fat_bytes:
        mini_fat_bytes = mini_fat_bytes.ljust(-(-len(mini_fat_bytes) // SECTOR) * SECTOR, b"\xff")
    mini_fat_start = alloc(mini_fat_bytes)
    mini_fat_sectors = len(mini_fat_bytes) // SECTOR

    root = entries[0]
    root.start = alloc(bytes(mini_stream))
    root.size = len(mini_stream)

    for e in big_streams:
        e.start = alloc(e.data)

    # FAT sectors (each holds 128 entries, including entries for themselves).
    fat_sector_count = 0
    while True:
        total = len(fat) + fat_sector_count
        needed = -(-total // 128)
        if needed == fat_sector_count:
            break
        fat_sector_count = needed
    assert fat_sector_count <= 109, "file too large for header-only DIFAT"
    fat_start = len(sectors)
    fat.extend([FATSECT] * fat_sector_count)
    fat.extend([FREESECT] * (fat_sector_count * 128 - len(fat)))
    fat_bytes = b"".join(struct.pack("<I", v) for v in fat)
    for i in range(fat_sector_count):
        sectors.append(fat_bytes[i * SECTOR : (i + 1) * SECTOR])

    # Directory entries.
    dir_bytes = bytearray()
    for e in entries:
        name = e.name.encode("utf-16-le") + b"\x00\x00"
        assert len(name) <= 64
        dir_bytes += name.ljust(64, b"\x00")
        dir_bytes += struct.pack("<HBB", len(name), e.type, 1)  # color: black
        dir_bytes += struct.pack("<III", e.left, e.right, e.child)
        dir_bytes += b"\x00" * 16  # CLSID
        dir_bytes += struct.pack("<I", 0)  # state bits
        dir_bytes += b"\x00" * 16  # creation / modified time
        start = e.start if e.type != 1 else 0
        size = e.size if e.type != 1 else 0
        dir_bytes += struct.pack("<IQ", start, size)
    empty = bytearray(b"\x00" * 64) + struct.pack("<HBB", 0, 0, 0) + struct.pack("<III", NOSTREAM, NOSTREAM, NOSTREAM) + b"\x00" * 44
    while len(dir_bytes) < dir_count * SECTOR:
        dir_bytes += empty
    for i in range(dir_count):
        sectors[dir_start + i] = bytes(dir_bytes[i * SECTOR : (i + 1) * SECTOR])

    header = bytearray()
    header += bytes.fromhex("D0CF11E0A1B11AE1")
    header += b"\x00" * 16
    header += struct.pack("<HHHHH", 0x003E, 0x0003, 0xFFFE, 9, 6)
    header += b"\x00" * 6
    header += struct.pack("<I", 0)  # directory sectors (0 for v3)
    header += struct.pack("<I", fat_sector_count)
    header += struct.pack("<I", dir_start)
    header += struct.pack("<I", 0)  # transaction signature
    header += struct.pack("<I", MINI_CUTOFF)
    header += struct.pack("<I", mini_fat_start if mini_fat_sectors else ENDOFCHAIN)
    header += struct.pack("<I", mini_fat_sectors)
    header += struct.pack("<I", ENDOFCHAIN)  # first DIFAT sector
    header += struct.pack("<I", 0)  # DIFAT sector count
    difat = [fat_start + i for i in range(fat_sector_count)] + [FREESECT] * (109 - fat_sector_count)
    header += b"".join(struct.pack("<I", v) for v in difat)
    assert len(header) == SECTOR
    return bytes(header) + b"".join(sectors)
