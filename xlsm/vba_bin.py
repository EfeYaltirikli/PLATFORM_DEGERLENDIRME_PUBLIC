# -*- coding: utf-8 -*-
"""Builds a vbaProject.bin (MS-OVBA inside an MS-CFB file) from VBA source
text, without Excel.  The modules carry source only (no compiled p-code):
_VBA_PROJECT says version 0xFFFF, so Excel compiles the source on open.

    build(modules, path)    modules = [(name, kind, source)]
                            kind: "std" | "doc_wb" | "doc_ws"
"""

import random
import struct
import uuid

ENDOFCHAIN = 0xFFFFFFFE
FREESECT = 0xFFFFFFFF
FATSECT = 0xFFFFFFFD
NOSTREAM = 0xFFFFFFFF
CODEPAGE = 1252
STDOLE = False

DOC_HEAD = {
    "doc_wb": "0{00020819-0000-0000-C000-000000000046}",
    "doc_ws": "0{00020820-0000-0000-C000-000000000046}",
}


# ---------------------------------------------------------------------------
#  MS-OVBA 2.4.1 compression
# ---------------------------------------------------------------------------

def _copy_token_help(diff):
    bit = 4
    while (1 << bit) < diff:
        bit += 1
    length_mask = 0xFFFF >> bit
    max_len = length_mask + 3
    return bit, length_mask, max_len


def _compress_chunk(data):
    out = bytearray()
    pos = 0
    n = len(data)
    while pos < n:
        flag = 0
        flag_at = len(out)
        out.append(0)
        for bit in range(8):
            if pos >= n:
                break
            best_len, best_off = 0, 0
            if pos > 0:
                bc, _lm, max_len = _copy_token_help(pos)
                lim = min(max_len, n - pos)
                for cand in range(pos - 1, -1, -1):
                    ln = 0
                    while ln < lim and data[cand + ln] == data[pos + ln]:
                        ln += 1
                    if ln > best_len:
                        best_len, best_off = ln, pos - cand
                        if ln == lim:
                            break
            if best_len >= 3:
                bc, _lm, _ml = _copy_token_help(pos)
                tok = ((best_off - 1) << (16 - bc)) | (best_len - 3)
                out += struct.pack("<H", tok)
                flag |= 1 << bit
                pos += best_len
            else:
                out.append(data[pos])
                pos += 1
        out[flag_at] = flag
    return bytes(out)


def compress(data):
    out = bytearray(b"\x01")
    for i in range(0, len(data), 4096):
        chunk = data[i:i + 4096]
        comp = _compress_chunk(chunk)
        if len(comp) > 4096 and len(chunk) == 4096:
            out += struct.pack("<H", 0x3000 | (4096 + 2 - 3)) + chunk
        else:
            size = len(comp) + 2
            out += struct.pack("<H", 0xB000 | (size - 3)) + comp
    return bytes(out)


def decompress(data):
    """For the tests: the reverse of compress()."""
    assert data[0] == 1
    out = bytearray()
    pos = 1
    while pos < len(data):
        hdr = struct.unpack_from("<H", data, pos)[0]
        size = (hdr & 0x0FFF) + 3
        end = pos + size
        pos += 2
        start = len(out)
        if not hdr & 0x8000:
            out += data[pos:pos + 4096]
            pos += 4096
            continue
        while pos < end:
            flag = data[pos]
            pos += 1
            for bit in range(8):
                if pos >= end:
                    break
                if flag & (1 << bit):
                    tok = struct.unpack_from("<H", data, pos)[0]
                    pos += 2
                    bc, lm, _ml = _copy_token_help(len(out) - start)
                    ln = (tok & lm) + 3
                    off = (tok >> (16 - bc)) + 1
                    for _k in range(ln):
                        out.append(out[-off])
                else:
                    out.append(data[pos])
                    pos += 1
    return bytes(out)


# ---------------------------------------------------------------------------
#  PROJECT stream password/protection fields (MS-OVBA 2.4.3)
# ---------------------------------------------------------------------------

def _encrypt(data, project_id, seed=None):
    seed = random.randint(0, 255) if seed is None else seed
    version = 2
    proj_key = sum(project_id.encode("ascii")) & 0xFF
    out = bytearray([seed, seed ^ version, seed ^ proj_key])
    unenc1, enc1, enc2 = proj_key, seed ^ proj_key, seed ^ version
    for _c in range((seed & 6) // 2):
        temp = random.randint(0, 255)
        b = temp ^ ((enc2 + unenc1) & 0xFF)
        out.append(b)
        enc2, enc1, unenc1 = enc1, b, temp
    for temp in struct.pack("<I", len(data)) + bytes(data):
        b = temp ^ ((enc2 + unenc1) & 0xFF)
        out.append(b)
        enc2, enc1, unenc1 = enc1, b, temp
    return out.hex().upper()


# ---------------------------------------------------------------------------
#  the VBA streams
# ---------------------------------------------------------------------------

def _rec(rid, payload):
    return struct.pack("<HI", rid, len(payload)) + payload


def _dir_stream(modules, project_name="VBAProject"):
    cp = lambda s: s.encode("cp%d" % CODEPAGE)
    u16 = lambda s: s.encode("utf-16-le")
    d = bytearray()
    d += _rec(0x0001, struct.pack("<I", 3))           # SYSKIND 64-bit (as Excel writes it)
    d += _rec(0x004A, struct.pack("<I", 0x00020000))  # COMPATVERSION
    d += _rec(0x0002, struct.pack("<I", 0x409))       # LCID
    d += _rec(0x0014, struct.pack("<I", 0x409))       # LCIDINVOKE
    d += _rec(0x0003, struct.pack("<H", CODEPAGE))    # CODEPAGE
    d += _rec(0x0004, cp(project_name))               # NAME
    d += _rec(0x0005, b"") + _rec(0x0040, b"")        # DOCSTRING
    d += _rec(0x0006, b"") + _rec(0x003D, b"")        # HELPFILEPATH
    d += _rec(0x0007, struct.pack("<I", 0))           # HELPCONTEXT
    d += _rec(0x0008, struct.pack("<I", 0))           # LIBFLAGS
    d += struct.pack("<HIIH", 0x0009, 4, 1, 0)        # VERSION (size 4, but 6 bytes)
    d += _rec(0x000C, b"") + _rec(0x003C, b"")        # CONSTANTS
    # reference: stdole (optional - Excel's own files on Mac have none)
    if not STDOLE:
        return _modules_part(d, modules)
    name = "stdole"
    libid = ("*\\G{00020430-0000-0000-C000-000000000046}#2.0#0#"
             "C:\\Windows\\System32\\stdole2.tlb#OLE Automation")
    d += _rec(0x0016, cp(name)) + _rec(0x003E, u16(name))
    lb = cp(libid)
    d += _rec(0x000D, struct.pack("<I", len(lb)) + lb + struct.pack("<IH", 0, 0))
    return _modules_part(d, modules)


def _modules_part(d, modules):
    cp = lambda s: s.encode("cp%d" % CODEPAGE)
    u16 = lambda s: s.encode("utf-16-le")
    d += _rec(0x000F, struct.pack("<H", len(modules)))
    d += _rec(0x0013, struct.pack("<H", 0xFFFF))      # PROJECTCOOKIE
    for mname, kind, _src in modules:
        d += _rec(0x0019, cp(mname)) + _rec(0x0047, u16(mname))
        d += _rec(0x001A, cp(mname)) + _rec(0x0032, u16(mname))
        d += _rec(0x001C, b"") + _rec(0x0048, b"")
        d += _rec(0x0031, struct.pack("<I", 0))       # MODULEOFFSET: source at 0
        d += _rec(0x001E, struct.pack("<I", 0))       # HELPCONTEXT
        d += _rec(0x002C, struct.pack("<H", 0xFFFF))  # COOKIE
        d += _rec(0x0021 if kind == "std" else 0x0022, b"")
        d += _rec(0x002B, b"")                        # module terminator
    d += _rec(0x0010, b"")                            # dir terminator
    return bytes(d)


def _module_source(name, kind, src):
    lines = ['Attribute VB_Name = "%s"' % name]
    if kind != "std":
        lines += ['Attribute VB_Base = "%s"' % DOC_HEAD[kind],
                  "Attribute VB_GlobalNameSpace = False",
                  "Attribute VB_Creatable = False",
                  "Attribute VB_PredeclaredId = True",
                  "Attribute VB_Exposed = True",
                  "Attribute VB_TemplateDerived = False",
                  "Attribute VB_Customizable = True"]
    body = [l for l in src.replace("\r\n", "\n").split("\n")
            if not l.startswith("Attribute VB_Name")]
    text = "\r\n".join(lines + body).rstrip("\r\n") + "\r\n"
    return text.encode("cp%d" % CODEPAGE)


def _project_stream(modules, project_id):
    out = ['ID="%s"' % project_id]
    for name, kind, _s in modules:
        out.append(("Document=%s/&H00000000" % name) if kind != "std" else ("Module=%s" % name))
    out += ['Name="VBAProject"', 'HelpContextID="0"', 'VersionCompatible32="393222000"',
            'CMG="%s"' % _encrypt(b"\x00\x00\x00\x00", project_id),
            'DPB="%s"' % _encrypt(b"\x00", project_id),
            'GC="%s"' % _encrypt(b"\xFF", project_id),
            "", "[Host Extender Info]",
            "&H00000001={3832D640-CF90-11CF-8E43-00A0C911005A};VBE;&H00000000",
            "", "[Workspace]"]
    for name, _k, _s in modules:
        out.append("%s=0, 0, 0, 0, C" % name)
    return ("\r\n".join(out) + "\r\n").encode("cp%d" % CODEPAGE)


def _project_wm(modules):
    b = bytearray()
    for name, _k, _s in modules:
        b += name.encode("cp%d" % CODEPAGE) + b"\x00" + name.encode("utf-16-le") + b"\x00\x00"
    return bytes(b + b"\x00\x00")


# ---------------------------------------------------------------------------
#  MS-CFB writer (version 3, 512-byte sectors)
# ---------------------------------------------------------------------------

class _Node:
    def __init__(self, name, kind, data=b""):
        self.name, self.kind, self.data = name, kind, data   # kind 5 root, 1 storage, 2 stream
        self.children = []
        self.left = self.right = self.child = NOSTREAM
        self.start, self.size = 0, 0


def _key(n):
    return (len(n.name), n.name.upper())


def _tree(nodes):
    """Balanced binary tree of the sorted siblings; returns the root id."""
    if not nodes:
        return NOSTREAM
    m = len(nodes) // 2
    nodes[m].left = _tree(nodes[:m])
    nodes[m].right = _tree(nodes[m + 1:])
    return nodes[m].sid


def write_cfb(root, path):
    order = []

    def walk(n):
        n.sid = len(order)
        order.append(n)
        for ch in sorted(n.children, key=_key):
            walk(ch)
    walk(root)
    for n in order:
        if n.kind in (1, 5):
            n.child = _tree(sorted(n.children, key=_key))

    streams = [n for n in order if n.kind == 2]
    mini = [n for n in streams if len(n.data) < 4096]
    big = [n for n in streams if len(n.data) >= 4096]

    ministream = bytearray()
    minifat = []
    for n in mini:
        n.size = len(n.data)
        if not n.data:
            n.start = ENDOFCHAIN
            continue
        cnt = (len(n.data) + 63) // 64
        n.start = len(ministream) // 64
        for k in range(cnt):
            minifat.append(n.start + k + 1 if k < cnt - 1 else ENDOFCHAIN)
        ministream += n.data + b"\x00" * (cnt * 64 - len(n.data))

    sec = lambda nbytes: (nbytes + 511) // 512
    n_dir = sec(len(order) * 128)
    n_mfat = sec(len(minifat) * 4)
    n_ms = sec(len(ministream))
    n_big = sum(sec(len(n.data)) for n in big)
    n_fat = 1
    while True:
        total = n_fat + n_dir + n_mfat + n_ms + n_big
        need = (total + 127) // 128
        if need <= n_fat:
            break
        n_fat = need
    assert n_fat <= 109

    fat = [FREESECT] * (n_fat * 128)
    cur = 0

    def chain(count):
        nonlocal cur
        start = cur
        for k in range(count):
            fat[cur] = cur + 1 if k < count - 1 else ENDOFCHAIN
            cur += 1
        return start if count else ENDOFCHAIN

    for k in range(n_fat):
        fat[cur] = FATSECT
        cur += 1
    dir_start = chain(n_dir)
    mfat_start = chain(n_mfat)
    ms_start = chain(n_ms)
    for n in big:
        n.size = len(n.data)
        n.start = chain(sec(len(n.data)))
    root.start = ms_start if n_ms else ENDOFCHAIN
    root.size = len(ministream)

    hdr = bytearray(512)
    hdr[0:8] = bytes.fromhex("D0CF11E0A1B11AE1")
    struct.pack_into("<HHHHH", hdr, 24, 0x003E, 0x0003, 0xFFFE, 9, 6)
    struct.pack_into("<I", hdr, 44, n_fat)
    struct.pack_into("<I", hdr, 48, dir_start)
    struct.pack_into("<I", hdr, 56, 0x1000)
    struct.pack_into("<I", hdr, 60, mfat_start if n_mfat else ENDOFCHAIN)
    struct.pack_into("<I", hdr, 64, n_mfat)
    struct.pack_into("<I", hdr, 68, ENDOFCHAIN)
    struct.pack_into("<I", hdr, 72, 0)
    for k in range(109):
        struct.pack_into("<I", hdr, 76 + 4 * k, k if k < n_fat else FREESECT)

    body = bytearray()
    body += b"".join(struct.pack("<I", x) for x in fat)
    d = bytearray()
    for n in order:
        e = bytearray(128)
        nm = n.name.encode("utf-16-le")
        e[0:len(nm)] = nm
        struct.pack_into("<HBB", e, 64, len(nm) + 2, n.kind, 1)
        struct.pack_into("<III", e, 68, n.left, n.right, n.child)
        struct.pack_into("<II", e, 116, n.start if n.kind != 1 else 0, n.size if n.kind != 1 else 0)
        d += e
    while len(d) < n_dir * 512:
        e = bytearray(128)
        struct.pack_into("<III", e, 68, NOSTREAM, NOSTREAM, NOSTREAM)
        d += e
    body += d
    mf = b"".join(struct.pack("<I", x) for x in minifat)
    body += mf + b"\xFF" * (n_mfat * 512 - len(mf))
    body += ministream + b"\x00" * (n_ms * 512 - len(ministream))
    for n in big:
        body += n.data + b"\x00" * (sec(len(n.data)) * 512 - len(n.data))
    with open(path, "wb") as f:
        f.write(bytes(hdr) + bytes(body))


def build(modules, path):
    project_id = "{%s}" % str(uuid.uuid4()).upper()
    root = _Node("Root Entry", 5)
    vba = _Node("VBA", 1)
    root.children = [vba,
                     _Node("PROJECT", 2, _project_stream(modules, project_id)),
                     _Node("PROJECTwm", 2, _project_wm(modules))]
    vba.children = [_Node("_VBA_PROJECT", 2, b"\xCC\x61\xFF\xFF\x00\x00\x00"),
                    _Node("dir", 2, compress(_dir_stream(modules)))]
    for name, kind, src in modules:
        vba.children.append(_Node(name, 2, compress(_module_source(name, kind, src))))
    write_cfb(root, path)
