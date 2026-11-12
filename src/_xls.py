"""Doc va ghi file .xls (BIFF8 trong vo OLE2) khong can Excel hay thu vien ngoai."""
import os
import struct

BOF = 0x0809
EOF_ = 0x000A
BOUNDSHEET = 0x0085
SST = 0x00FC
CONTINUE = 0x003C
LABELSST = 0x00FD
LABEL = 0x0204
NUMBER = 0x0203
RK = 0x027E
MULRK = 0x00BD
BLANK = 0x0201
MULBLANK = 0x00BE
FORMULA = 0x0006
ROW = 0x0208
FONT = 0x0031
FORMAT = 0x041E
XF = 0x00E0
STYLE = 0x0293
COLWIDTH = 0x007D
MERGEDCELLS = 0x00E5


def _records(data):
    out, pos, n = [], 0, len(data)
    while pos + 4 <= n:
        rid, ln = struct.unpack("<HH", data[pos:pos + 4])
        out.append((rid, data[pos + 4:pos + 4 + ln]))
        pos += 4 + ln
    return out


def _rk(rk):
    if rk & 0x02:
        val = rk >> 2
        if val & 0x20000000:
            val -= 0x40000000
        num = float(val)
    else:
        (num,) = struct.unpack("<d", struct.pack("<Q", (rk & 0xFFFFFFFC) << 32))
    if rk & 0x01:
        num /= 100.0
    return num


class _Sst:
    def __init__(self, blocks):
        self.blocks, self.bi, self.pos = blocks, 0, 0

    def eof(self):
        return self.bi >= len(self.blocks)

    def read(self, n):
        out = b""
        while n > 0 and not self.eof():
            b = self.blocks[self.bi]
            if self.pos >= len(b):
                self.bi += 1
                self.pos = 0
                continue
            take = min(n, len(b) - self.pos)
            out += b[self.pos:self.pos + take]
            self.pos += take
            n -= take
        return out

    def parse(self):
        head = self.read(8)
        if len(head) < 8:
            return []
        _total, unique = struct.unpack("<II", head)
        strings = []
        while len(strings) < unique and not self.eof():
            try:
                strings.append(self._one())
            except (IndexError, struct.error):
                break
        return strings

    def _one(self):
        (nchar,) = struct.unpack("<H", self.read(2))
        (opts,) = struct.unpack("<B", self.read(1))
        high = bool(opts & 0x01)
        rich = bool(opts & 0x08)
        far = bool(opts & 0x04)
        nrun = cbext = 0
        if rich:
            (nrun,) = struct.unpack("<H", self.read(2))
        if far:
            (cbext,) = struct.unpack("<I", self.read(4))
        pieces, left = [], nchar
        while left > 0:
            b = self.blocks[self.bi]
            avail = len(b) - self.pos
            width = 2 if high else 1
            if avail < width:
                self.bi += 1
                self.pos = 0
                if self.eof():
                    break
                (o,) = struct.unpack("<B", self.read(1))
                high = bool(o & 0x01)
                continue
            take = min(left, avail // width)
            raw = self.read(take * width)
            pieces.append(raw.decode("utf-16-le" if high else "latin-1", "replace"))
            left -= take
            if left > 0 and not self.eof():
                (o,) = struct.unpack("<B", self.read(1))
                high = bool(o & 0x01)
        if nrun:
            self.read(4 * nrun)
        if cbext:
            self.read(cbext)
        return "".join(pieces)


def read_xls(path):
    with open(path, "rb") as fh:
        data = fh.read()
    if data[:8] != b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1":
        raise ValueError("File khong phai .xls (OLE2). Hay luu lai dang Excel 97-2003.")
    stream = read_ole_stream(data)
    records = _records(stream)

    blocks, started = [], False
    for rid, body in records:
        if rid == SST:
            started = True
            blocks.append(body)
        elif rid == CONTINUE and started:
            blocks.append(body)
        elif started:
            break
    sst = _Sst(blocks).parse() if blocks else []

    cells, cols, rows, merges, sheets, fonts, formats = {}, {}, {}, [], [], [], {}
    bofs = 0
    for rid, body in records:
        if rid == BOF:
            bofs += 1
            continue
        if bofs < 1:
            continue
        try:
            if rid == BOUNDSHEET:
                (nchar,) = struct.unpack("<B", body[6:7])
                opts = body[7]
                wide = 2 if opts & 1 else 1
                raw = body[8:8 + nchar * wide]
                sheets.append(raw.decode("utf-16-le" if opts & 1 else "latin-1", "replace"))
            elif rid == FORMAT:
                idx = struct.unpack("<H", body[:2])[0]
                (nchar,) = struct.unpack("<H", body[2:4])
                opts = body[4]
                wide = 2 if opts & 1 else 1
                raw = body[5:5 + nchar * wide]
                formats[idx] = raw.decode("utf-16-le" if opts & 1 else "latin-1", "replace")
            elif rid == FONT:
                height = struct.unpack("<H", body[:2])[0]
                weight = struct.unpack("<H", body[4:6])[0]
                (nchar,) = struct.unpack("<B", body[14:15])
                opts = body[15]
                wide = 2 if opts & 1 else 1
                raw = body[16:16 + nchar * wide]
                fonts.append({"size": height / 20.0, "bold": weight >= 700,
                              "name": raw.decode("utf-16-le" if opts & 1 else "latin-1", "replace")})
            elif rid == COLWIDTH:
                c0, c1, wd = struct.unpack("<HHH", body[:6])
                for c in range(c0, min(c1 + 1, c0 + 400)):
                    cols[c] = wd
            elif rid == ROW:
                r = struct.unpack("<H", body[:2])[0]
                if len(body) >= 8:
                    ht = struct.unpack("<H", body[6:8])[0]
                    if ht:
                        rows[r] = ht
            elif rid == MERGEDCELLS:
                (cnt,) = struct.unpack("<H", body[:2])
                for i in range(cnt):
                    merges.append(struct.unpack("<HHHH", body[2 + i * 8:10 + i * 8]))
            elif rid == LABELSST:
                r, c, xf, idx = struct.unpack("<HHHI", body[:10])
                cells[(r, c)] = ("s", sst[idx] if idx < len(sst) else "", xf)
            elif rid == LABEL:
                r, c, xf = struct.unpack("<HHH", body[:6])
                (nchar,) = struct.unpack("<H", body[6:8])
                opts = body[8]
                wide = 2 if opts & 1 else 1
                raw = body[9:9 + nchar * wide]
                cells[(r, c)] = ("s", raw.decode("utf-16-le" if opts & 1 else "latin-1", "replace"), xf)
            elif rid == NUMBER:
                r, c, xf = struct.unpack("<HHH", body[:6])
                (v,) = struct.unpack("<d", body[6:14])
                cells[(r, c)] = ("n", v, xf)
            elif rid == RK:
                r, c, xf, rk = struct.unpack("<HHHI", body[:10])
                cells[(r, c)] = ("n", _rk(rk), xf)
            elif rid == MULRK:
                r, c0 = struct.unpack("<HH", body[:4])
                n = (len(body) - 6) // 6
                for i in range(n):
                    xf, rk = struct.unpack("<HI", body[4 + i * 6:10 + i * 6])
                    cells[(r, c0 + i)] = ("n", _rk(rk), xf)
            elif rid == BLANK:
                r, c, xf = struct.unpack("<HHH", body[:6])
                cells.setdefault((r, c), ("b", None, xf))
            elif rid == MULBLANK:
                r, c0 = struct.unpack("<HH", body[:4])
                n = (len(body) - 6) // 2
                for i in range(n):
                    (xf,) = struct.unpack("<H", body[4 + i * 2:6 + i * 2])
                    cells.setdefault((r, c0 + i), ("b", None, xf))
            elif rid == FORMULA:
                r, c, xf = struct.unpack("<HHH", body[:6])
                cells.setdefault((r, c), ("f", None, xf))
        except (struct.error, IndexError):
            continue
    return {"cells": cells, "cols": cols, "rows": rows, "merges": merges,
            "sheets": sheets, "formats": formats, "fonts": fonts}


# ---------------------------------------------------------------- OLE2 doc
def read_ole_stream(data, names=("Workbook", "Book")):
    sect_shift = struct.unpack("<H", data[30:32])[0]
    ssz = 1 << sect_shift
    num_fat = struct.unpack("<I", data[44:48])[0]
    first_dir = struct.unpack("<I", data[48:52])[0]
    mini_cutoff = struct.unpack("<I", data[56:60])[0]
    first_minifat = struct.unpack("<I", data[60:64])[0]
    num_minifat = struct.unpack("<I", data[64:68])[0]
    difat_start = struct.unpack("<I", data[68:72])[0]

    fat_sects = []
    for i in range(109):
        v = struct.unpack("<I", data[76 + i * 4:80 + i * 4])[0]
        if v < 0xFFFFFFFA:
            fat_sects.append(v)
    sec, guard = difat_start, 0
    while sec < 0xFFFFFFFA and guard < 1024:
        off = 512 + sec * ssz
        ents = list(struct.unpack("<%dI" % (ssz // 4), data[off:off + ssz]))
        fat_sects.extend([e for e in ents[:-1] if e < 0xFFFFFFFA])
        sec = ents[-1]
        guard += 1

    def sector(s):
        off = 512 + s * ssz
        chunk = data[off:off + ssz]
        return chunk + b"\x00" * (ssz - len(chunk))

    fat = []
    for s in fat_sects[:max(1, num_fat)]:
        fat.extend(struct.unpack("<%dI" % (ssz // 4), sector(s)))
    if not fat:
        raise ValueError("File .xls thieu bang FAT")
    if len(fat) <= 1:
        raise ValueError("File .xls chi co 1 sector, khong hop le")

    def chain(start):
        out, cur, g = [], start, 0
        while cur < 0xFFFFFFFA and g < 200000:
            out.append(cur)
            cur = fat[cur] if cur < len(fat) else 0xFFFFFFFE
            g += 1
        return out

    def read_chain(start, size=None):
        b = b"".join(sector(s) for s in chain(start))
        return b[:size] if size else b

    directory = read_chain(first_dir)
    entries = []
    for i in range(len(directory) // 128):
        e = directory[i * 128:(i + 1) * 128]
        (nlen,) = struct.unpack("<H", e[64:66])
        typ = e[66]
        if typ not in (1, 2, 5):
            continue
        name = e[:max(0, nlen - 2)].decode("utf-16-le", "replace") if nlen >= 2 else ""
        start, size = struct.unpack("<II", e[116:124])
        entries.append({"name": name, "type": typ, "start": start, "size": size})
    if not entries:
        raise ValueError("File .xls bi hong (khong co bang muc luc)")

    minifat = []
    if num_minifat and first_minifat < 0xFFFFFFFA:
        raw = read_chain(first_minifat)
        minifat = list(struct.unpack("<%dI" % (len(raw) // 4), raw))
    root = entries[0]
    mini_stream = read_chain(root["start"], root["size"]) if root["size"] else b""

    def read_mini(start, size):
        out, cur, g, mssz = b"", start, 0, 64
        while cur < 0xFFFFFFFA and g < 200000:
            out += mini_stream[cur * mssz:(cur + 1) * mssz]
            cur = minifat[cur] if cur < len(minifat) else 0xFFFFFFFE
            g += 1
        return out[:size]

    # Neu Root khong co mini stream (size = 0) thi moi stream deu nam trong FAT
    # thuong - ke ca stream nho hon nguong mini (file do chinh chuong trinh nay ghi).
    # Stream do chinh tool ghi luon bat dau bang BOF -> kiem tra truoc.
    def looks_biff(sector_index):
        head = sector(sector_index)[:4]
        return len(head) == 4 and head[:2] == b"\x09\x08"

    use_mini = bool(root["size"]) and bool(minifat)
    for want in names:
        for e in entries:
            if e["type"] == 2 and e["name"] == want:
                if use_mini and e["size"] < mini_cutoff and not looks_biff(e["start"]):
                    data_mini = read_mini(e["start"], e["size"])
                    if data_mini[:2] == b"\x09\x08":
                        return data_mini
                    return data_mini
                return read_chain(e["start"], e["size"])
    raise ValueError("Khong tim thay du lieu bang tinh trong file .xls")


# ================================================================ GHI .XLS
def _rec(rid, body):
    return struct.pack("<HH", rid, len(body)) + body


def _cstr(text):
    return text.encode("latin-1", "replace") + b"\x00"


def _xlstr(text):
    """Chuoi BIFF8: do dai 2 byte + co bao + utf-16le."""
    text = "" if text is None else str(text)
    return struct.pack("<HB", len(text), 0x01) + text.encode("utf-16-le")


def _short_str(text):
    text = "" if text is None else str(text)
    return struct.pack("<BB", len(text), 0x01) + text.encode("utf-16-le")


class Styles:
    HALIGN = {"general": 0, "left": 1, "center": 2, "right": 3, "fill": 4, "justify": 5}
    VALIGN = {"top": 0, "center": 1, "bottom": 2, "justify": 3}

    def __init__(self):
        self.fonts = [{"name": "Arial", "size": 10, "bold": False, "italic": False,
                       "colour": 0x08, "underline": 0}]
        self.formats = []
        self.xfs = [self._base_spec()]
        self._index = {self._key(self.xfs[0]): 15}

    @staticmethod
    def _base_spec():
        return {"font": 0, "numfmt": 0, "halign": "general", "valign": "bottom",
                "wrap": False, "border": False, "back": None}

    @staticmethod
    def _key(spec):
        return (spec["font"], spec["numfmt"], spec["halign"], spec["valign"],
                spec["wrap"], spec["border"], spec["back"])

    def font(self, name="Arial", size=10, bold=False, italic=False, colour=0x08, underline=0):
        spec = {"name": name, "size": size, "bold": bold, "italic": italic,
                "colour": colour, "underline": underline}
        for i, f in enumerate(self.fonts):
            if f == spec:
                return i
        self.fonts.append(spec)
        return len(self.fonts) - 1

    def numfmt(self, code):
        if not code:
            return 0
        for idx, c in self.formats:
            if c == code:
                return idx
        idx = 164 + len(self.formats)
        self.formats.append((idx, code))
        return idx

    def xf(self, font=0, numfmt=0, halign="general", valign="bottom", wrap=False,
           border=False, back=None):
        spec = {"font": font, "numfmt": numfmt, "halign": halign, "valign": valign,
                "wrap": wrap, "border": border, "back": back}
        key = self._key(spec)
        if key in self._index:
            return self._index[key]
        self.xfs.append(spec)
        idx = 15 + len(self.xfs) - 1
        self._index[key] = idx
        return idx


class Sheet:
    def __init__(self, name, max_row=200, max_col=40):
        self.name = name
        self.cells = {}
        self.merges = []
        self.col_widths = {}
        self.row_heights = {}
        self.max_row = max_row
        self.max_col = max_col

    def text(self, r, c, value, xf=15):
        if value is None or value == "":
            return self.blank(r, c, xf)
        self.cells[(r, c)] = ("s", str(value), xf)

    def number(self, r, c, value, xf=15):
        if value is None or value == "":
            return self.blank(r, c, xf)
        try:
            self.cells[(r, c)] = ("n", float(value), xf)
        except (TypeError, ValueError):
            self.cells[(r, c)] = ("s", str(value), xf)

    def blank(self, r, c, xf=15):
        self.cells[(r, c)] = ("b", None, xf)

    def merge(self, r1, r2, c1, c2):
        self.merges.append((r1, r2, c1, c2))

    def width(self, c, chars):
        self.col_widths[c] = int(round(chars * 256))

    def height(self, r, twips):
        self.row_heights[r] = int(twips)


class _SstBuilder:
    def __init__(self):
        self.strings, self.map = [], {}

    def add(self, text):
        text = "" if text is None else str(text)
        if text in self.map:
            return self.map[text]
        self.map[text] = len(self.strings)
        self.strings.append(text)
        return self.map[text]

    def build(self, limit=8000):
        """SST + EXTSST. Tra ve (bytes, so lan xuat hien cua tung chuoi)."""
        counts = {}
        for s in self.strings:
            counts[s] = counts.get(s, 0) + 1
        parts = [bytearray(struct.pack("<II", len(self.strings), len(self.strings)))]
        for s in self.strings:
            chunk = _xlstr(s)
            if len(parts[-1]) + len(chunk) > limit and len(parts[-1]) > 8:
                parts.append(bytearray())
            parts[-1] += chunk
        out = _rec(SST, bytes(parts[0]))
        for p in parts[1:]:
            out += _rec(CONTINUE, bytes(p))
        return out, counts


def _extsst(strings, counts):
    """EXTSST: bang chi muc vi tri chuoi trong SST (LibreOffice/Excel de doc)."""
    cint, cinfo = 8, 8          # so bucket = ceil(n/8), moi entry 8 byte
    nbucket = max(1, (len(strings) + cint - 1) // cint)
    body = struct.pack("<H", cinfo)
    base = 0
    for i in range(nbucket):
        s = strings[i * cint]
        body += struct.pack("<IH", base, len(s))
        base += 3 + 2 * len(s)
    return _rec(0x00FF, body)


_PALETTE = [(0x08, (0, 0, 0)), (0x09, (255, 255, 255)), (0x0A, (255, 0, 0)),
            (0x0B, (0, 255, 0)), (0x0C, (0, 0, 255)), (0x0D, (255, 255, 0)),
            (0x0E, (255, 0, 255)), (0x0F, (0, 255, 255)), (0x16, (192, 192, 192)),
            (0x1A, (204, 255, 204)), (0x15, (255, 255, 153)), (0x2A, (153, 204, 255)),
            (0x37, (255, 204, 153))]


def _pal(rgb):
    best, bd = 0x16, None
    for idx, col in _PALETTE:
        d = sum((a - b) ** 2 for a, b in zip(col, rgb))
        if bd is None or d < bd:
            bd, best = d, idx
    return best


def _style_records(styles):
    out = b""
    for f in styles.fonts:
        body = struct.pack("<HHHHH", int(round(f["size"] * 20)),
                           0x0002 if f["italic"] else 0,
                           f["colour"], 700 if f["bold"] else 400, f["underline"])
        body += struct.pack("<HHBBB", 0, 0, 0x00, 0x00, 0x00)
        body += _cstr(f["name"][:31])
        out += _rec(FONT, body)
    for idx, code in styles.formats:
        out += _rec(FORMAT, struct.pack("<H", idx) + _xlstr(code))

    def xf_body(spec):
        body = bytearray()
        body += struct.pack("<HHHH", spec["numfmt"], spec["font"], 0xFFF5, 0x0020)
        body += bytes([Styles.HALIGN.get(spec["halign"], 0),
                       Styles.VALIGN.get(spec["valign"], 2),
                       1 if spec["wrap"] else 0, 0, 0, 0])
        b = 1 if spec["border"] else 0
        body += bytes([b, 0, b, 0])
        body += bytes([0, 0, 0, 0])
        if spec["back"] is None:
            body += bytes([0, 0, 0, 0])
        else:
            idx = _pal(spec["back"])
            body += bytes([1, idx, 0x41, idx])
        body += struct.pack("<HH", 0x20C0, 0x00C0)
        return bytes(body)

    default = Styles._base_spec()
    for _ in range(15):
        out += _rec(XF, xf_body(default))
    for spec in styles.xfs:
        out += _rec(XF, xf_body(spec))
    out += _rec(STYLE, struct.pack("<HHH", 0x8010, 0, 0x8000))
    return out


def _sheet_stream(sheet, sst):
    """Sinh stream cua 1 sheet theo dung thu tu record ma Excel/LO yeu cau."""
    out = b""
    out += _rec(BOF, struct.pack("<HHHHII", 0x0600, 0x0010, 0x0DBB, 0x07CD, 0x000200D9, 0x00080606))
    out += _rec(0x000D, struct.pack("<H", 1))                    # CALCMODE auto
    out += _rec(0x000C, struct.pack("<H", 100))                  # CALCCOUNT
    out += _rec(0x000F, struct.pack("<H", 1))                    # REFMODE A1
    out += _rec(0x0011, struct.pack("<H", 0))                    # ITERATION
    out += _rec(0x0010, struct.pack("<d", 0.001))                # DELTA
    out += _rec(0x005F, struct.pack("<H", 1))                    # SAVERECALC
    out += _rec(0x0082, struct.pack("<H", 1))                    # GRIDSET
    out += _rec(0x0080, struct.pack("<HHHH", 0, 0, 0, 0))        # GUTS
    out += _rec(0x0225, struct.pack("<HH", 0x0001, 0x00FF))      # DEFAULTROWHEIGHT
    out += _rec(0x0081, struct.pack("<H", 0x04C1))               # WSBOOL
    out += _rec(0x0014, b"")                                     # HEADER
    out += _rec(0x0015, b"")                                     # FOOTER
    out += _rec(0x0083, struct.pack("<H", 1))                    # HCENTER
    out += _rec(0x0084, struct.pack("<H", 0))                    # VCENTER
    out += _rec(0x0026, struct.pack("<d", 0.25))                 # LEFTMARGIN
    out += _rec(0x0027, struct.pack("<d", 0.25))                 # RIGHTMARGIN
    out += _rec(0x0028, struct.pack("<d", 0.75))                 # TOPMARGIN
    out += _rec(0x0029, struct.pack("<d", 0.75))                 # BOTTOMMARGIN
    out += _rec(0x00A1, struct.pack("<HHHHHHHH", 9, 100, 1, 1, 1, 0, 600, 600)
                + struct.pack("<dd", 0.3, 0.3) + struct.pack("<H", 1))  # SETUP
    out += _rec(0x0055, struct.pack("<H", 8))                    # DEFCOLWIDTH
    for c in sorted(sheet.col_widths):
        out += _rec(COLWIDTH, struct.pack("<HHHH", c, c, sheet.col_widths[c], 0))
    out += _rec(0x0200, struct.pack("<IIHHH", 0, sheet.max_row, 0, sheet.max_col, 0))
    out += _rec(0x023E, struct.pack("<HHHHHHHHHH", 0, 0, 0, 0, 0, 0, 0, 0, 0, 0))

    grouped = {}
    for (r, c), item in sheet.cells.items():
        grouped.setdefault(r, []).append(item + (c,))
    for r in sorted(grouped):
        items = sorted(grouped[r], key=lambda x: x[3])
        h = sheet.row_heights.get(r, 255)
        out += _rec(ROW, struct.pack("<HHHHHHHH", r, items[0][3], items[-1][3], h, 0, 0, 0, 0x0001))
        for kind, value, xf, c in items:
            if kind == "n" and value is not None:
                out += _rec(NUMBER, struct.pack("<HHH", r, c, xf) + struct.pack("<d", value))
            elif kind == "s":
                out += _rec(LABELSST, struct.pack("<HHHI", r, c, xf, sst.add(value)))
            else:
                out += _rec(BLANK, struct.pack("<HHH", r, c, xf))
    if sheet.merges:
        body = struct.pack("<H", len(sheet.merges))
        for (r1, r2, c1, c2) in sheet.merges:
            body += struct.pack("<HHHH", r1, r2, c1, c2)
        out += _rec(MERGEDCELLS, body)
    out += _rec(EOF_, b"")
    return out


def _globals_stream(sheet_offsets, styles, sst):
    out = b""
    out += _rec(BOF, struct.pack("<HHHHII", 0x0600, 0x0005, 0x0DBB, 0x07CC, 0x00000041, 0x00000130))
    out += _rec(0x00E1, struct.pack("<H", 0x04B0))
    out += _rec(0x00E2, b"")
    out += _rec(0x00C1, struct.pack("<H", 0))
    out += _rec(0x005C, _cstr("TienAnAuto"))
    out += _rec(0x0042, struct.pack("<H", 1200))
    out += _rec(0x0161, struct.pack("<H", 0))
    out += _rec(0x01C0, b"")
    out += _rec(0x013D, struct.pack("<H", 1))
    out += _rec(0x009C, struct.pack("<H", 14))
    out += _rec(0x0019, struct.pack("<H", 0))
    out += _rec(0x0012, struct.pack("<H", 0))
    out += _rec(0x0013, struct.pack("<H", 0))
    out += _rec(0x01AF, struct.pack("<H", 0))
    out += _rec(0x01BC, struct.pack("<H", 0))
    out += _rec(0x003D, struct.pack("<HHHHHHHH", 0, 0, 16384, 8192, 0x0026, 0, 0, 1))
    out += _rec(0x0040, struct.pack("<H", 0))
    out += _rec(0x008D, struct.pack("<H", 0))
    out += _rec(0x0022, struct.pack("<H", 0))
    out += _rec(0x000E, struct.pack("<H", 0))
    out += _rec(0x01B7, struct.pack("<H", 0))
    out += _rec(0x00DA, struct.pack("<H", 0))
    for name, offset in sheet_offsets:
        out += _rec(BOUNDSHEET, struct.pack("<IBB", offset, 0, 0) + _short_str(name))
    out += _style_records(styles)
    sst_bytes, counts = sst.build()
    out += sst_bytes
    out += _extsst(sst.strings, counts)
    out += _rec(EOF_, b"")
    return out


def _ole_wrap(stream):
    """Vo OLE2 v3: header 512 byte la sector -1, du lieu tinh tu byte 512 (sector 0)."""
    data = stream + b"\x00" * (512 - (len(stream) % 512))   # luon du 1 sector cuoi
    n_data = len(data) // 512
    fat_sect = n_data
    dir_sect = n_data + 1
    if dir_sect + 1 > 128:
        raise ValueError("Du lieu qua lon cho writer don gian")
    fat = []
    for i in range(128):
        if i < n_data - 1:
            fat.append(i + 1)
        elif i in (n_data - 1, fat_sect, dir_sect):
            fat.append(0xFFFFFFFE)
        else:
            fat.append(0xFFFFFFFF)

    def dirent(name, typ, start, size, child=0xFFFFFFFF):
        n = name.encode("utf-16-le") + b"\x00\x00"
        body = n + b"\x00" * (64 - len(n))
        body += struct.pack("<HBB", len(n), typ, 1)
        body += struct.pack("<III", 0xFFFFFFFF, 0xFFFFFFFF, child)
        body += b"\x00" * 36
        body += struct.pack("<II", start, size)
        body += b"\x00" * 4
        return body

    # Root chua mini stream rong nhung van tro toi sector hop le (gioi han 127)
    dirents = dirent("Root Entry", 5, 0x0000007F, 0, child=1)
    dirents += dirent("Workbook", 2, 0, len(stream))
    dirents += b"\x00" * (128 * 2)

    header = bytearray(512)
    header[0:8] = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"
    struct.pack_into("<HHHHH", header, 24, 0x003E, 0x0003, 0xFFFE, 9, 6)
    struct.pack_into("<I", header, 40, 0)
    struct.pack_into("<I", header, 44, 1)
    struct.pack_into("<I", header, 48, dir_sect)
    struct.pack_into("<I", header, 52, 0)
    struct.pack_into("<I", header, 56, 4096)
    struct.pack_into("<I", header, 60, 0xFFFFFFFE)
    struct.pack_into("<I", header, 64, 0)
    struct.pack_into("<I", header, 68, 0xFFFFFFFE)
    struct.pack_into("<I", header, 72, 0)
    struct.pack_into("<I", header, 76, fat_sect)
    for i in range(1, 109):
        struct.pack_into("<I", header, 76 + i * 4, 0xFFFFFFFF)
    fat_bytes = b"".join(struct.pack("<I", v) for v in fat)
    return bytes(header) + data + fat_bytes + dirents


def write_xls(path, sheets, styles=None):
    styles = styles or Styles()
    sst = _SstBuilder()
    offsets = [0] * len(sheets)
    for _ in range(4):
        blobs = [_sheet_stream(s, sst) for s in sheets]
        globals_stream = _globals_stream([(s.name, offsets[i]) for i, s in enumerate(sheets)],
                                         styles, sst)
        cursor, new_offsets = len(globals_stream), []
        for blob in blobs:
            new_offsets.append(cursor)
            cursor += len(blob)
        if new_offsets == offsets:
            break
        offsets = new_offsets
    blobs = [_sheet_stream(s, sst) for s in sheets]
    globals_stream = _globals_stream([(s.name, offsets[i]) for i, s in enumerate(sheets)],
                                     styles, sst)
    body = globals_stream + b"".join(blobs)
    ole = _ole_wrap(body)
    tmp = path + ".tmp"
    with open(tmp, "wb") as fh:
        fh.write(ole)
    os.replace(tmp, path)
