"""Ghi file quyet toan .xls dung dinh dang file mau (BIFF8 + vo OLE2 tu dung).

Nguyen ly (da kiem chung voi LibreOffice/Excel):
  * .xls = globals (BOF..SST..EOF: font, so dinh dang, bang chuoi) + phan sheet.
  * Phan sheet gom: header (BOF..COLINFO), ROW header cua cac hang tieu de,
    cac o tieu de (lay NGUYEN tu file mau: ten cot, ngay, dinh dang),
    cac hang du lieu, dong tong, va record ket thuc.
  * Trinh doc doc tuan tu theo vi tri record -> moi hang phai dung thu tu:
    ROW truoc, cac o sau. O so cua 1 hang gom trong 1 MULRK.
  * Font, so dinh dang (#,##0), do rong cot, duong ke, tron o: khong doi.
"""
import base64
import json
import os
import struct

ROW = 0x0208
MULRK = 0x00BD
NUMBER = 0x0203
LABELSST = 0x00FD
DIMENSIONS = 0x0200
PANE = 0x0041
EOF_ = 0x000A


def rec(rid, body):
    return struct.pack("<HH", rid, len(body)) + body


def rk_of(value):
    """float -> RK (chinh xac tuyet doi cho so nguyen nho va so 2 chu so le)."""
    iv = int(value)
    if iv == value and -(2 ** 29) < iv < 2 ** 29:
        return ((iv << 2) | 0x02) & 0xFFFFFFFF
    scaled = round(value * 100.0)
    if abs(scaled - value * 100.0) < 1e-7 and -(2 ** 29) < scaled < 2 ** 29:
        return ((int(scaled) << 2) | 0x03) & 0xFFFFFFFF
    return None


LABEL = 0x0204


def _unicode_field(text):
    """Chuoi Unicode BIFF8 (do dai 2 byte, co bao 1 = utf-16le)."""
    text = "" if text is None else str(text)
    return struct.pack("<HB", len(text), 0x01) + text.encode("utf-16-le")


def _exact(text):
    """Chuan hoa de SO KHOP: giu nguyen dau cau va khoang trang, chi bo dau tieng Viet."""
    import unicodedata
    s = unicodedata.normalize("NFD", str(text or ""))
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")
    for a, b in (("\u0111", "d"), ("\u0110", "D")):
        s = s.replace(a, b)
    return " ".join(s.split()).lower()


def norm(text):
    import unicodedata
    s = unicodedata.normalize("NFD", str(text or ""))
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")
    s = s.replace("đ", "d").replace("Đ", "D")
    return " ".join(s.lower().replace("(", " ").replace(")", " ").split())


class Template:
    """Khuon doc tu template.json (sinh bang tools/build_template.py)."""

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
        self.sst = list(data.get("sst") or [])
        self.sst_index = dict((s, i) for i, s in enumerate(self.sst))
        self._extra = []                 # chuoi moi chua co trong file mau
        self.first_row = int(data.get("first_row") or 4)
        self.data_rows = int(data.get("data_rows") or 27)
        self.totals_row = data.get("totals_row")
        self.name_col = int(data.get("name_col") or 0)
        self.name_xf = int(data.get("name_xf") or self.col_xf.get(self.name_col, 15))
        self.total_col = data.get("total_col")
        # cot TONG: QTY_THANH_TIEN (43,44 trong file mau) - bo sung dinh dang theo file mau
        if self.total_col is not None:
            self.total_col = int(self.total_col)
            self.col_xf.setdefault(self.total_col, 97)
            self.col_xf.setdefault(self.total_col + 1, 98)
        self.source_file = data.get("source_file") or ""
        self.day_columns = dict((int(k), v) for k, v in (data.get("day_columns") or {}).items())
        self.header_rows = dict((int(k), base64.b64decode(v))
                                for k, v in (data.get("header_rows") or {}).items())

    # --- thong tin khuon ------------------------------------------
    def columns(self):
        return sorted(self.col_xf)

    def max_col(self):
        cols = self.columns()
        return max(cols) if cols else 10

    def pairs(self):
        """Cac cap cot (so luong, thanh tien), gom ca cot TONG neu co."""
        cols = sorted(self.columns())
        if self.total_col is not None:
            cols = [c for c in cols if c not in (self.total_col, self.total_col + 1)]
        out = [(cols[i], cols[i + 1]) for i in range(0, len(cols) - 1, 2)]
        tp = self.total_pair()
        if tp and tp not in out:
            out.append(tp)
        return out

    def day_column_map(self):
        if self.day_columns:
            return [(c, c + 1) for c in sorted(self.day_columns)]
        return self.pairs()[:-1]

    def day_labels(self):
        return [self.day_columns[k] for k in sorted(self.day_columns)]

    def total_pair(self):
        """Cap cot TONG: 2 cot co cong thuc trong dong du lieu (file mau: AR=43, AS=44)."""
        if self.total_col is not None:
            return (self.total_col, self.total_col + 1)
        day_starts = set(self.day_columns) if self.day_columns else set()
        for pair in reversed(self.pairs()):
            if pair[0] not in day_starts:
                return pair
        return None

    def sst_id(self, text):
        """Chi so chuoi dung DUNG ten (khong phan biet hoa/thuong, khong bo dau ngoac).

        Bang chuoi dung chung: chi duoc dung lai khi ten trung khop; ten khac (vi du
        'Nuoc mam Nam Ngu 750ml' va 'Nuoc mam Nam Ngư (750ml)') phai ghi rieng, neu
        dung lai se ra sai ten tren bang quyet toan.
        """
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

    def register_extra(self, names):
        """Chuoi chua co trong file mau -> danh sach de bao nguoi dung."""
        self._extra = [n for n in names if self.sst_id(n) is None]
        return list(self._extra)

    def tail_fixed(self, frozen_at=None):
        """Record cuoi sheet: PANE doi thanh vung dong bang moi (0 = bo)."""
        out = b""
        pos = 0
        blob = self.tail
        while pos + 4 <= len(blob):
            rid, ln = struct.unpack("<HH", blob[pos:pos + 4])
            body = blob[pos + 4:pos + 4 + ln]
            if rid == PANE and frozen_at is not None:
                if frozen_at <= 0:
                    pos += 4 + ln
                    continue
                body = struct.pack("<HHHHH", 0, frozen_at, 0, frozen_at, 0x03)
            out += rec(rid, body)
            pos += 4 + ln
        return out


def _emit_mulrk(row, run):
    """Ghi 1 cum MULRK. run: [(col, xf, rk)] tang dan theo cot.

    MULRK chi bieu dien duoc cot LIEN TIEP (colFirst, colFirst+1, ...). Vi vay moi
    lan cot nhay (khong lien tiep) phai dong cum lai va mo cum moi - neu khong o
    sau se bi ghi lech sang cot khac. Truong cuoi cua record la so cot CUOI CUNG
    (colLast), khong phai xf - day la cho de viet nham nhat.
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


FORMULA = 0x0006
ARRAY = 0x0221


def _ranges_in_formula(body, cce):
    """Cac vung tham chieu dang ROW_RANGE (ptgRef 0x24/0x44/0x64) trong cc chua tung.

    Tra ve list (offset_tuong_doi_so_voi_dau_ptg, r_first, r_last); offset tro toi
    byte dau rwFirst cua toan vung (rwFirst=2 byte, rwLast=2 byte, col=2 byte).
    """
    out = []
    i = 0
    end = len(cce)
    while i < end:
        ptg = cce[i]
        if ptg in (0x24, 0x44, 0x64, 0x25, 0x45, 0x65, 0x3A, 0x3B, 0x3C,
                   0x5A, 0x5B, 0x5C):
            if i + 7 > end:
                break
            r_first, r_last = struct.unpack("<HH", cce[i + 1:i + 5])
            out.append((i + 1, r_first, r_last))
            i += 7
        elif ptg in (0x15, 0x16, 0x17, 0x18, 0x19, 0x1E, 0x1F, 0x1D):
            i += {0x15: 9, 0x16: 1, 0x17: 3, 0x18: 1, 0x19: 2, 0x1E: 1, 0x1F: 1,
                  0x1D: 1}[ptg]
        else:
            i += 1
    return out


def _rewrite_formula(body, new_last_row, calc=None):
    """Noi vung ROW_RANGE trong cong thuc den dong du lieu cuoi cung.

    calc: dict {(row, col): value}; neu co thi tinh lai ket qua luu san
    (Excel/LibreOffice doc ket qua nay truoc khi ban recalc).
    """
    body = bytearray(body)
    cce = struct.unpack("<H", body[20:22])[0]
    cce_start = 22
    ranges = _ranges_in_formula(body, body[cce_start:cce_start + cce])
    for (off, r_first, r_last) in ranges:
        if r_last != new_last_row:
            abs_off = cce_start + off + 2          # +2: bo qua rwFirst
            struct.pack_into("<H", body, abs_off, new_last_row)
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


def _shift_row(blob, new_row, new_last_row=None, calc=None):
    """Doi so hang cua cum record dong tong; noi vung SUM neu bang dai hon khuon."""
    out = b""
    pos = 0
    while pos + 4 <= len(blob):
        rid, ln = struct.unpack("<HH", blob[pos:pos + 4])
        body = bytearray(blob[pos + 4:pos + 4 + ln])
        if rid in (ROW, MULRK, NUMBER, LABELSST, 0x0201, 0x00BE, FORMULA) and len(body) >= 2:
            struct.pack_into("<H", body, 0, new_row)
        if rid == FORMULA and new_last_row is not None:
            body = bytearray(_rewrite_formula(bytes(body), new_last_row, calc))
        if rid == ARRAY and len(body) >= 6 and new_last_row is not None:
            head = bytes(body[:4])
            body = bytearray(head + _rewrite_formula(bytes(body[4:]), new_last_row, calc))
        out += rec(rid, bytes(body))
        pos += 4 + ln
    return out


def build_stream(tpl, rows, freeze=True):
    """rows: [{'name': str, 'values': {col: value}}] -> stream BIFF8 day du."""
    max_col = tpl.max_col()
    out = [tpl.globals, tpl.sheet_header]
    nrows = tpl.first_row + len(rows) + 2
    out.append(rec(DIMENSIONS, struct.pack("<IIHHH", 0, nrows, 0, max_col + 1, 0)))

    for hr in sorted(tpl.header_rows):
        raw = tpl.header_rows[hr]
        if raw:
            out.append(rec(ROW, raw))
    out.append(tpl.title_blob)

    for i, item in enumerate(rows):
        r = tpl.first_row + i
        cells = []
        name = item.get("name")
        sid = tpl.sst_id(name) if name else None
        if sid is not None:
            cells.append(("sst", rec(LABELSST, struct.pack("<HHHI", r, tpl.name_col,
                                                          tpl.name_xf, sid))))
        elif name:
            # ten chua co trong file mau: ghi duoi dang chuoi Unicode truc tiep (LABEL)
            # de khong phai sua bang chuoi dung chung
            cells.append(("label", rec(LABEL, struct.pack("<HHH", r, tpl.name_col,
                                                         tpl.name_xf) + _unicode_field(name))))
        cells = [c[1] if isinstance(c, tuple) else c for c in cells]
        vals = []
        for col in sorted(item.get("values") or {}):
            v = item["values"][col]
            if v is None:
                continue
            vals.append((col, tpl.col_xf.get(col, 15), float(v)))
        out.append(rec(ROW, struct.pack("<HHHHHHHH", r, 0, max_col, 0x01B3, 0, 0, 0, 1)))
        out.extend(cells)
        out.append(_mulrk_row(r, vals))

    if tpl.totals_blob:
        # dong cuoi cua vung du lieu: first_row + len(rows) - 1
        last_data = tpl.first_row + len(rows) - 1
        calc = {}
        for i, item in enumerate(rows):
            for col, val in (item.get("values") or {}).items():
                calc[(tpl.first_row + i, col)] = val
        out.append(_shift_row(tpl.totals_blob, tpl.first_row + len(rows), last_data, calc))
    out.append(tpl.tail_fixed(tpl.first_row if freeze else 0))
    return b"".join(out)


def write_xls(path, tpl, rows, freeze=True):
    import _ole
    stream = build_stream(tpl, rows, freeze)
    tmp = path + ".tmp"
    _ole.write_ole_workbook(tmp, stream)
    os.replace(tmp, path)
    return path
