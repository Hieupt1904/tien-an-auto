"""Ghi file quyet toan .xls dung dinh dang file mau (BIFF8 + vo OLE2 tu dung).

Cau truc sheet trong khuon (da kiem chung voi Excel/LibreOffice):
    header (BOF..COLINFO)  -> do rong cot, vung in, dinh dang chung
    ROW header + o cua hang tieu de 0..3 (ten cot + nhan ngay)
    ROW + o cho tung mat hang (ten + so lieu)
    ROW + o cua dong tong (cong thuc SUM)
    PANE / SETUP / WINDOW2 / EOF
Trinh doc doc tuan tu theo vi tri record, nen moi hang phai dat dung thu tu:
ROW header truoc, cac o sau.

RIENG: file mau la "Excel 95 workbook" (build id cu), LibreOffice chi doc dung
khi ca bang nam trong MOT khoi (khong chen ROW header cho tung hang). Vi vay
ta chuyen sang build id BIFF8 chuan cua Excel 97 va chen ROW header day du.
"""
import base64
import json
import os
import struct

ROW = 0x0208
FORMULA = 0x0006
ARRAY = 0x0221
MULRK = 0x00BD
NUMBER = 0x0203
LABELSST = 0x00FD
LABEL = 0x0204
BLANK = 0x0201
MULBLANK = 0x00BE
DIMENSIONS = 0x0200
PANE = 0x0041
EOF_ = 0x000A
BOF = 0x0809
COLINFO = 0x007D


def rec(rid, body):
    return struct.pack("<HH", rid, len(body)) + body


def rk_of(value):
    iv = int(value)
    if iv == value and -(2 ** 29) < iv < 2 ** 29:
        return ((iv << 2) | 0x02) & 0xFFFFFFFF
    scaled = round(value * 100.0)
    if abs(scaled - value * 100.0) < 1e-7 and -(2 ** 29) < scaled < 2 ** 29:
        return ((int(scaled) << 2) | 0x03) & 0xFFFFFFFF
    return None


def _exact(text):
    import unicodedata
    s = unicodedata.normalize("NFD", str(text or ""))
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")
    for a, b in (("\u0111", "d"), ("\u0110", "D")):
        s = s.replace(a, b)
    return " ".join(s.split()).lower()


def _unicode_field(text):
    text = "" if text is None else str(text)
    return struct.pack("<HB", len(text), 0x01) + text.encode("utf-16-le")


class Template:
    """Khuon doc tu template.json (sinh boi tools/build_template.py)."""

    def __init__(self, path):
        self.path = os.path.abspath(path)
        with open(path, "r", encoding="utf-8") as fh:
            data = json.load(fh)
        self.globals = base64.b64decode(data["globals"])
        self.sheet_header = base64.b64decode(data["sheet_header"])
        self.title_blob = base64.b64decode(data["title_blob"])
        self.totals_blob = base64.b64decode(data.get("totals_blob") or b"")
        self.tail = base64.b64decode(data["tail"])
        self.col_xf = dict((int(k), v) for k, v in data["col_xf"].items())
        self.name_xf = int(data.get("name_xf") or self.col_xf.get(0, 15))
        self.sst = list(data.get("sst") or [])
        self.sst_index = dict((s, i) for i, s in enumerate(self.sst))
        self.first_row = int(data.get("first_row") or 4)
        self.data_rows = int(data.get("data_rows") or 27)
        self.totals_row = data.get("totals_row")
        self.name_col = int(data.get("name_col") or 0)
        self.source_file = data.get("source_file") or ""
        self.day_columns = dict((int(k), v) for k, v in (data.get("day_columns") or {}).items())
        total = data.get("total_col")
        self.total_col = int(total) if total is not None else None
        self._globals_patched = None

    # --- thong tin -------------------------------------------------
    def columns(self):
        return sorted(self.col_xf)

    def max_col(self):
        cols = self.columns()
        return max(cols) if cols else 10

    def day_column_map(self):
        """[(cot so luong, cot thanh tien)] cua tung ngay, theo nhan o dong tieu de."""
        return [(c, c + 1) for c in sorted(self.day_columns)]

    def day_labels(self):
        return [self.day_columns[k] for k in sorted(self.day_columns)]

    def total_pair(self):
        if self.total_col is not None:
            return (self.total_col, self.total_col + 1)
        return None

    def sst_id(self, text):
        """Chi so chuoi dung DUNG ten (chi bo dau tieng Viet, giu dau ngoac)."""
        if text is None:
            return None
        text = str(text)
        if text in self.sst_index:
            return self.sst_index[text]
        target = _exact(text)
        for i, s in enumerate(self.sst):
            if _exact(s) == target:
                return i
        return None

    def column_widths(self):
        """Do rong cot hien tai: {cot: width/256} (doc tu COLINFO trong sheet_header)."""
        out = {}
        blob = self.sheet_header
        pos = 0
        while pos + 4 <= len(blob):
            rid, ln = struct.unpack("<HH", blob[pos:pos + 4])
            body = blob[pos + 4:pos + 4 + ln]
            if rid == COLINFO and len(body) >= 8:
                c0, c1, wd, _xf = struct.unpack("<HHHH", body[:8])
                for c in range(c0, min(c1, c0 + 400) + 1):
                    out[c] = wd / 256.0
            pos += 4 + ln
        return out

    def sheet_header_wide(self, min_width=9.0, n_rows=None, n_cols=None):
        """sheet_header voi moi cot tu cot ngay tro di rong it nhat min_width (tranh ###)."""
        blob = self.sheet_header
        cols = self.column_widths()
        if not cols:
            return blob
        wide_from = min(self.day_columns) if self.day_columns else 0
        out = b""
        pos = 0
        while pos + 4 <= len(blob):
            rid, ln = struct.unpack("<HH", blob[pos:pos + 4])
            body = bytearray(blob[pos + 4:pos + 4 + ln])
            if rid == COLINFO and len(body) >= 8:
                c0, c1, wd, xf = struct.unpack("<HHHH", bytes(body[:8]))
                want = max(wd / 256.0, min_width) if c0 >= wide_from else wd / 256.0
                struct.pack_into("<H", body, 4, int(round(want * 256)))
            elif rid == DIMENSIONS and len(body) >= 14 and n_rows is not None:
                struct.pack_into("<IIHHH", body, 0, 0, n_rows, 0, n_cols or 0, 0)
            out += rec(rid, bytes(body))
            pos += 4 + ln
        return out

    def sheet_header_dims(self, n_rows, n_cols):
        """sheet_header voi DIMENSIONS cap nhat theo so dong/cot thuc te."""
        blob = self.sheet_header
        out = b""
        pos = 0
        seen = False
        while pos + 4 <= len(blob):
            rid, ln = struct.unpack("<HH", blob[pos:pos + 4])
            body = bytearray(blob[pos + 4:pos + 4 + ln])
            if rid == DIMENSIONS and len(body) >= 14:
                struct.pack_into("<IIHHH", body, 0, 0, n_rows, 0, n_cols, 0)
                seen = True
            out += rec(rid, bytes(body))
            pos += 4 + ln
        if not seen:
            out = rec(DIMENSIONS, struct.pack("<IIHHH", 0, n_rows, 0, n_cols, 0)) + out
        return out

    def globals_stream(self):
        """Globals cua file mau, khong them bot.

        XF trong BIFF duoc danh so 0..133 (file mau co du 134 dinh dang); khi ghi ta
        tham chieu dung chi so do, nen phai dung nguyen bang XF goc.
        """
        return self.globals

    # --- record ket thuc ------------------------------------------
    def tail_fixed(self):
        """Record cuoi sheet; doi SETUP sang A4 ngang, in vua 1 trang."""
        out = b""
        pos = 0
        blob = self.tail
        while pos + 4 <= len(blob):
            rid, ln = struct.unpack("<HH", blob[pos:pos + 4])
            body = blob[pos + 4:pos + 4 + ln]
            if rid == PANE:
                pos += 4 + ln
                continue
            if rid == 0x00A1 and len(body) >= 12:      # SETUP
                b = bytearray(body)
                struct.pack_into("<H", b, 0, 9)        # A4
                struct.pack_into("<H", b, 2, 1)        # kho giay
                struct.pack_into("<H", b, 6, 1)        # ngang
                body = bytes(b)
            out += rec(rid, body)
            pos += 4 + ln
        return out


# ------------------------------------------------------------------ records
def _emit_mulrk(row, run):
    """run: [(col, xf, rk)] tang dan; MULRK chi ghi duoc cot LIEN TIEP."""
    out = b""
    i = 0
    while i < len(run):
        start = run[i][0]
        chunk = [run[i]]
        j = i + 1
        while j < len(run) and run[j][0] == chunk[-1][0] + 1:
            chunk.append(run[j])
            j += 1
        body = struct.pack("<HH", row, start)
        for col, xf, rk in chunk:
            body += struct.pack("<HI", xf, rk)
        body += struct.pack("<H", chunk[-1][0])       # colLast
        out += rec(MULRK, body)
        i = j
    return out


def _mulrk_row(row, values):
    out = b""
    run = []
    for col, xf, val in values:
        rk = rk_of(val)
        if rk is None:
            if run:
                out += _emit_mulrk(row, run)
                run = []
            out += rec(NUMBER, struct.pack("<HHH", row, col, xf) + struct.pack("<d", float(val)))
        else:
            run.append((col, xf, rk))
    if run:
        out += _emit_mulrk(row, run)
    return out


def _emit_mulblank(row, run):
    """run: [(col, xf)] tang dan -> BLANK le, MULBLANK (>=3 cot) khi lien tiep.

    BIFF8 MULBLANK = row(2), colFirst(2), <xf(2) cho TUNG o>, colLast(2).
    Moi o trong dai mang XF RIENG cua no (mau that: cot 13..42 la
    119,120,119,120,... vi so luong / thanh tien dung 2 kieu khac nhau).
    Vi vay KHONG duoc gop cac cot khac XF vao cung mot dai.
    """
    out = b""
    i = 0
    while i < len(run):
        start = run[i][0]
        chunk = [run[i]]
        j = i + 1
        while j < len(run) and run[j][0] == chunk[-1][0] + 1:
            chunk.append(run[j])
            j += 1
        if len(chunk) >= 3:
            body = struct.pack("<HH", row, start)
            for _c, xf in chunk:
                body += struct.pack("<H", xf)          # XF rieng cho TUNG o
            body += struct.pack("<H", chunk[-1][0])
            out += rec(MULBLANK, body)
        else:
            for c, xf in chunk:
                out += rec(BLANK, struct.pack("<HHH", row, c, xf))
        i = j
    return out


def _row_header(row, last_col, height=0x01B3):
    """ROW record BIFF8: rw, colMic, colMac, miyRw, ...

    colMac phai la so COT (khong phai chi so cot) tuc last_col + 1; va phai phu
    moi o cua hang. Dat sai cho (vi du de 0) lam Excel bao
    "File error: data may have been lost" du LibreOffice van mo duoc.
    """
    return rec(ROW, struct.pack("<HHHHHHHH", row, 0, max(last_col, 0) + 1,
                                height, 0, 0, 0, 1))


def _ranges_in_formula(cce):
    out = []
    i = 0
    while i < len(cce):
        ptg = cce[i]
        if ptg in (0x24, 0x44, 0x64, 0x25, 0x45, 0x65, 0x3A, 0x3B, 0x3C,
                   0x5A, 0x5B, 0x5C):
            if i + 7 > len(cce):
                break
            r_first, r_last = struct.unpack("<HH", cce[i + 1:i + 5])
            out.append((i + 1, r_first, r_last))
            i += 7
        elif ptg in (0x15, 0x16, 0x17, 0x18, 0x19, 0x1E, 0x1F, 0x1D):
            i += {0x15: 9, 0x16: 1, 0x17: 3, 0x18: 1, 0x19: 2, 0x1E: 1, 0x1F: 1, 0x1D: 1}[ptg]
        else:
            i += 1
    return out


def _rewrite_formula(body, new_last_row, calc=None, rebuild=None, row=None, is_total=False):
    """Noi vung SUM den dong du lieu cuoi; tinh lai ket qua luu san."""
    body = bytearray(body)
    cce = struct.unpack("<H", body[20:22])[0]
    cce_start = 22
    if rebuild is not None:
        # cong thuc cot TONG: viet lai thanh tong cac cot nguon
        xf, sources = rebuild
        value = 0.0
        if is_total:
            # dong tong: cong toan bo bang tren cot nguon
            for src in sources:
                for (r_t, c_t), v in (calc or {}).items():
                    if c_t == src and isinstance(v, (int, float)):
                        value += v
        else:
            for src in sources:
                v = (calc or {}).get((row, src))
                if isinstance(v, (int, float)):
                    value += v
        target = struct.unpack("<H", body[2:4])[0]
        # re-save the record: 6 header bytes, 8 bytes value, 2 flags, 4 chn, 4 cce = 24 bytes
        new = struct.pack("<HHH", row, target, xf) + struct.pack("<d", value) + \
            struct.pack("<H", 0x0020) + bytes.fromhex("1f0004ff")
        # Cong thuc cot TONG: "=B5+D5+F5+..." (ptgRef 7 byte + ptgAdd).
        # ptgRef (0x44): row(2) col(2) + grbit tuong doi (2 byte) = 7 byte.
        # ptgRef 0x44 = 5 byte: row(2) + col(2). Voi bit co 0xC000, "col" la
        # DO LECH so voi o dang chua cong thuc (khong phai so cot tuyet doi).
        cce_new = b""
        for i, c in enumerate(sources):
            if i:
                cce_new += b"\x03"                      # ptgAdd
            # "=SUM(B5,D5,F5,...)" -> cac ptgRef (0x44, 5 byte) + ptgFuncVar SUM.
        # LibreOffice doc ptgRef 5 byte la: rw(2) + col(2) lech 2 so voi cot that,
        # nen bu +2 cho dung cot.
        cce_new = b""
        for c in sources:
            cce_new += bytes([0x44]) + struct.pack("<HH", row, c)
        cce_new += bytes([0x22, len(sources)]) + struct.pack("<H", 0x0004)
        cce_field = struct.pack("<H", len(cce_new))
        return bytes(new + cce_field + cce_new)

    ranges = _ranges_in_formula(body[cce_start:cce_start + cce])
    for (off, r_first, r_last) in ranges:
        if r_last != new_last_row:
            struct.pack_into("<H", body, cce_start + off + 2, new_last_row)
    if calc is not None and ranges:
        r_first = ranges[0][1]
        col = struct.unpack("<H", body[2:4])[0]
        total = 0.0
        for r in range(r_first, max(new_last_row, r_first) + 1):
            v = calc.get((r, col))
            if isinstance(v, (int, float)):
                total += v
        struct.pack_into("<d", body, 6, total)
    return bytes(body)


def _shift_row(blob, new_row, new_last_row, calc=None, rebuild=None):
    out = b""
    pos = 0
    while pos + 4 <= len(blob):
        rid, ln = struct.unpack("<HH", blob[pos:pos + 4])
        body = bytearray(blob[pos + 4:pos + 4 + ln])
        if rid in (ROW, MULRK, NUMBER, LABELSST, BLANK, MULBLANK, FORMULA) and len(body) >= 2:
            struct.pack_into("<H", body, 0, new_row)
        if rid == FORMULA:
            col = struct.unpack("<H", bytes(body[2:4]))[0]
            body = bytearray(_rewrite_formula(bytes(body), new_last_row, calc,
                                             (rebuild or {}).get(col), new_row,
                                             is_total=bool(rebuild)))
        out += rec(rid, bytes(body))
        pos += 4 + ln
    return out


def _cells_for_row(tpl, r, item, total_row, last_data, calc, rebuild):
    """Sinh toan bo record o cua mot hang, theo dung thu tu cot tang dan.

    O trong khuon dung bang XF 0..133 cua file mau; MULRK chi gom cac cot lien
    tiep cung mot kieu, va giu nguyen cong thuc o dong tong.
    """
    vals = []
    name = item.get("name") if item is not None else None
    if name:
        sid = tpl.sst_id(name)
        if sid is not None:
            vals.append((tpl.name_col, tpl.name_xf, "s", sid))
        else:
            vals.append((tpl.name_col, tpl.name_xf, "t", name))
    tp = tpl.total_pair()
    if item is not None:
        for col in sorted(item.get("values") or {}):
            v = item["values"][col]
            if v is None:
                continue
            if tp and col in tp:
                continue                       # cot TONG ghi bang cong thuc, khong ghi so
            vals.append((col, tpl.col_xf.get(col, 15), "n", float(v)))
    vals.sort(key=lambda t: t[0])

    # O KHONG CO SO LIEU van phai ghi de giu KHUNG KE: file mau ghi MULBLANK cho
    # toan bo cot con lai cua hang. Neu bo qua, bang se ho rong va mat vien.
    have = dict((col, xf) for col, xf, _k, _p in vals)
    blank_run = []

    def flush_blanks():
        nonlocal out, blank_run
        if blank_run:
            out += _emit_mulblank(r, blank_run)
            blank_run = []

    out = b""
    run = []
    for col in range(0, tpl.max_col() + 1):
        if col not in have:
            # cot trong: dung XF cua chinh cot do (giu vien + dinh dang)
            if run:
                out += _emit_mulrk(r, run)
                run = []
            xf = tpl.col_xf.get(col, tpl.name_xf)
            if blank_run and (blank_run[-1][1] != xf or blank_run[-1][0] != col - 1):
                flush_blanks()
            blank_run.append((col, xf))
            continue
        if blank_run:
            flush_blanks()
        xf = have[col]
        kind, payload = None, None
        for c2, x2, k2, p2 in vals:
            if c2 == col:
                kind, payload = k2, p2
                break
        if kind != "n":
            if run:
                out += _emit_mulrk(r, run)
                run = []
            if kind == "s":
                out += rec(LABELSST, struct.pack("<HHHI", r, col, xf, payload))
            else:
                out += rec(LABEL, struct.pack("<HHH", r, col, xf) + _unicode_field(payload))
            continue
        rk = rk_of(payload)
        if rk is None:
            if run:
                out += _emit_mulrk(r, run)
                run = []
            out += rec(NUMBER, struct.pack("<HHH", r, col, xf) + struct.pack("<d", payload))
            continue
        if run and (run[-1][1] != xf or run[-1][0] != col - 1):
            out += _emit_mulrk(r, run)
            run = []
        run.append((col, xf, rk))
    if run:
        out += _emit_mulrk(r, run)
    flush_blanks()
    return out


def _total_formula_row(tpl, r, sources, xf, value, col):
    """Cong thuc "=SUM(B5,D5,...)" cho cot TONG o mot dong du lieu."""
    cce = b""
    for c in sources:
        cce += bytes([0x44]) + struct.pack("<HH", r, c)
    cce += bytes([0x22, len(sources)]) + struct.pack("<H", 0x0004)
    body = struct.pack("<HHH", r, col, xf)
    body += struct.pack("<d", float(value))
    body += struct.pack("<H", 0x0020)
    body += bytes.fromhex("1f0004ff")
    body += struct.pack("<H", len(cce)) + cce
    return rec(FORMULA, body)


def build_stream(tpl, rows, freeze=True):
    """rows: [{'name': str, 'values': {col: value}}] -> stream BIFF8 day du."""
    max_col = tpl.max_col()
    last_data = tpl.first_row + len(rows) - 1
    total_row = tpl.first_row + len(rows)
    # sheet_header co san DIMENSIONS tu file mau: chi cap nhat so dong/cot trong
    # chinh record do. Them mot DIMENSIONS thu hai se lam Excel bao
    # "File error: data may have been lost".
    out = [tpl.globals_stream(),
           tpl.sheet_header_wide(n_rows=total_row + 1, n_cols=max_col + 1)]
    out.append(tpl.title_blob)

    tp = tpl.total_pair()
    q_src = [q for q, _a in tpl.day_column_map()]
    a_src = [a for _q, a in tpl.day_column_map()]
    for i, item in enumerate(rows):
        r = tpl.first_row + i
        out.append(_row_header(r, max_col))
        out.append(_cells_for_row(tpl, r, item, total_row, last_data, None, None))
        if tp and item is not None:
            vals = item.get("values") or {}
            qty = sum(vals.get(c, 0) or 0 for c in q_src)
            amt = sum(vals.get(c, 0) or 0 for c in a_src)
            out.append(_total_formula_row(tpl, r, q_src,
                                          tpl.col_xf.get(tp[0], 121), qty, tp[0]))
            out.append(_total_formula_row(tpl, r, a_src,
                                          tpl.col_xf.get(tp[1], 122), amt, tp[1]))

    calc = {}
    for i, item in enumerate(rows):
        for col, val in (item.get("values") or {}).items():
            calc[(tpl.first_row + i, col)] = val
    if tpl.totals_blob:
        rebuild = {}
        tp = tpl.total_pair()
        if tp:
            rebuild[tp[0]] = (tpl.col_xf.get(tp[0], 121), [q for q, _a in tpl.day_column_map()])
            rebuild[tp[1]] = (tpl.col_xf.get(tp[1], 122), [a for _q, a in tpl.day_column_map()])
        out.append(_shift_row(tpl.totals_blob, total_row, last_data, calc, rebuild))
    out.append(tpl.tail_fixed())
    return b"".join(out)


def write_xls(path, tpl, rows, freeze=True):
    import _ole
    stream = build_stream(tpl, rows, freeze)
    tmp = path + ".tmp"
    _ole.write_ole_workbook(tmp, stream)
    os.replace(tmp, path)
    return path
