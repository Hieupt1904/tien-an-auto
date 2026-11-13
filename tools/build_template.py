"""Tao khuon (template.json) tu file .xls mau. Chay lai moi khi thay file mau.

Khuon gom:
  globals       : stream truoc phan sheet (font, number format, bang chuoi dung chung)
  sheet_header  : BOF..COLINFO cua sheet (do rong cot, dinh dang vung)
  title_blob    : ROW header + cac o cua hang tieu de 0..3 (ten cot, ngay, dinh dang)
  totals_blob   : ROW header + cac o dong tong (co cong thuc)
  col_xf        : dinh dang (xf) dung cho tung cot khi ghi so lieu
  name_xf       : dinh dang o ten mat hang
  day_columns   : cot so luong cua tung ngay + nhan "Ngay dd-mm" in o dong tieu de
  total_col     : cot dau tien cua cap TONG
  sst           : bang chuoi (de dung lai ten mat hang co san)
  tail          : cac record ket thuc sheet (PANE, SETUP, WINDOW2, EOF...)
"""
import base64
import json
import os
import struct
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(os.path.dirname(HERE), "src"))
import _xls  # noqa: E402

ROW = 0x0208
FORMULA = 0x0006
MULRK = 0x00BD
NUMBER = 0x0203
LABELSST = 0x00FD
BLANK = 0x0201
MULBLANK = 0x00BE
RK = 0x027E
TITLE_ROWS = (0, 1, 2, 3)
CELL_IDS = (0x00FD, 0x0203, 0x00BD, 0x00BE, 0x0201, 0x0204, 0x0006, 0x04BC)


def enc(blob):
    return base64.b64encode(blob).decode("ascii")


def cell_rows(records, sheet):
    """Gom record theo hang: ROW header + cac o cung hang (o so nam sau day cac ROW)."""
    rows = {}
    pos = 0
    for rid, body in records:
        size = 4 + len(body)
        if rid == ROW:
            r = struct.unpack("<H", body[:2])[0]
            rows.setdefault(r, []).append((rid, body))
        elif rid in CELL_IDS and len(body) >= 2:
            r = struct.unpack("<H", body[:2])[0]
            rows.setdefault(r, []).append((rid, body))
        pos += size
    return rows


def blob_of(rows, hrs):
    out = b""
    for hr in hrs:
        for rid, body in rows.get(hr, []):
            out += struct.pack("<HH", rid, len(body)) + body
    return out


def main(src, out):
    stream = _xls.read_ole_stream(open(src, "rb").read())
    recs = _xls._records(stream)
    pos = seen = 0
    glob_end = None
    for rid, body in recs:
        if rid == 0x0809:
            seen += 1
            if seen == 2:
                glob_end = pos
                break
        pos += 4 + len(body)
    if glob_end is None:
        raise SystemExit("khong tim thay sheet trong file mau")
    globals_blob = stream[:glob_end]
    sheet = stream[glob_end:]
    srecs = _xls._records(sheet)

    # header: tu BOF den truoc ROW dau tien
    pos = 0
    for rid, body in srecs:
        if rid == ROW:
            break
        pos += 4 + len(body)
    sheet_header = sheet[:pos]

    rows = cell_rows(srecs, sheet)
    max_data = max([r for r in rows if r < 3000] or [0])

    # dong tong: hang co cong thuc va khong co ten vat tu o cot A
    totals_row = None
    for hr in sorted(rows):
        if hr < 4 or hr > max_data:
            continue
        has_f = any(rid in (0x0006, 0x04BC) for rid, _b in rows[hr])
        has_name = any(rid == 0x00FD and len(b) >= 4 and struct.unpack("<H", b[2:4])[0] == 0
                       for rid, b in rows[hr])
        if has_f and not has_name:
            totals_row = hr
    data_rows = (totals_row - 1) if totals_row else max_data

    # bang chuoi
    blocks, started = [], False
    for rid, body in _xls._records(globals_blob):
        if rid == 0x00FC:
            started = True
            blocks.append(body)
        elif rid == 0x003C and started:
            blocks.append(body)
        elif started:
            break
    sst = _xls._Sst(blocks).parse() if blocks else []

    # cot ngay: o dong tieu de hang 2 co chu "Ngay"
    day_cols = {}
    for rid, body in rows.get(2, []):
        if rid == 0x00FD and len(body) >= 10:
            _r, c, _xf, idx = struct.unpack("<HHHI", body[:10])
            txt = sst[idx] if idx < len(sst) else ""
            if "ng" in txt.lower() and any(ch.isdigit() for ch in txt):
                day_cols[c] = txt

    # Dinh dang (XF) cho tung cot: lay tu DONG DU LIEU DAU TIEN (hang 4) vi do
    # la hang co du moi kieu o (so luong, thanh tien, cot trong); hang sau chi
    # nhap mot phan cot nen khong du tin cay.
    col_xf = {}
    name_xf = 15
    hero = rows.get(4, [])
    for rid, body in hero:
        if rid == 0x00BD:                                  # MULRK: cot 1..n
            c0 = struct.unpack("<H", body[2:4])[0]
            c_last = struct.unpack("<H", body[-2:])[0]
            for k in range(c0, c_last + 1):
                col_xf[k] = struct.unpack("<H", body[4 + (k - c0) * 6:6 + (k - c0) * 6])[0]
        elif rid == 0x00BE:                                # MULBLANK: cot trong
            c0 = struct.unpack("<H", body[2:4])[0]
            c_last = struct.unpack("<H", body[-2:])[0]
            xf = struct.unpack("<H", body[4:6])[0]
            for k in range(c0, c_last + 1):
                col_xf[k] = xf
        elif rid in (NUMBER, RK) and len(body) >= 6:
            c, xf = struct.unpack("<HH", body[2:6])
            col_xf[c] = xf
        elif rid in (LABELSST, BLANK) and len(body) >= 6:
            c, xf = struct.unpack("<HH", body[2:6])
            col_xf[c] = xf
            if c == 0:
                name_xf = xf
    # cot ma dong 4 khong co (hang du lieu ngan hon bang): lay tu cac hang sau,
    # van theo thu tu hang gan nhat co du lieu.
    for hr in sorted(rows):
        if hr <= 4 or (totals_row and hr >= totals_row):
            continue
        for rid, body in rows[hr]:
            if rid == 0x00BD:
                c0 = struct.unpack("<H", body[2:4])[0]
                c_last = struct.unpack("<H", body[-2:])[0]
                for k in range(c0, c_last + 1):
                    col_xf.setdefault(k, struct.unpack("<H", body[4 + (k - c0) * 6:6 + (k - c0) * 6])[0])
            elif rid in (NUMBER, RK) and len(body) >= 6:
                c, xf = struct.unpack("<HH", body[2:6])
                col_xf.setdefault(c, xf)
    # cap cot TONG: chi lay XF cho cac cot BEN NGOAI bang ngay (dong tong o cac
    # cot ngay dung chung dinh dang voi cot du lieu, khong duoc ghi de).
    if totals_row:
        for rid, body in rows[totals_row]:
            if rid == FORMULA and len(body) >= 6:
                c, xf = struct.unpack("<HH", body[2:6])
                if c not in day_cols and c not in col_xf:
                    col_xf[c] = xf

    # xf la chi so trong file mau -> doi sang chi so xf cua CHINH khuon
    # (khuon xep them 5 xf mac dinh luc build_stream: 0..4, du lieu bat dau tu 5)
    # O day ta chi can giu nguyen xf vi globals giu nguyen ca bang xf cua file mau.
    # cot TONG: cot co cong thuc trong dong du lieu dau tien
    total_col = None
    for hr in sorted(rows):
        if hr < 4 or hr > data_rows:
            continue
        for rid, body in rows[hr]:
            if rid == FORMULA and len(body) >= 6:
                c = struct.unpack("<H", body[2:4])[0]
                if c not in day_cols:
                    total_col = c
                    break
        if total_col is not None:
            break
    if total_col is None and day_cols:
        total_col = max(day_cols) + 2

    # tail: tu PANE cuoi cung den EOF
    tail_start = None
    for i, (rid, body) in enumerate(srecs):
        if rid == 0x0041:
            tail_start = i
    tail = b""
    if tail_start is not None:
        for rid, body in srecs[tail_start:]:
            tail += struct.pack("<HH", rid, len(body)) + body
    if not tail.endswith(b"\x0a\x00\x00\x00"):
        tail += b"\x0a\x00\x00\x00"

    data = {
        "globals": enc(globals_blob),
        "sheet_header": enc(sheet_header),
        "title_blob": enc(blob_of(rows, TITLE_ROWS)),
        "title_blob_has_rows": True,
        "totals_blob": enc(blob_of(rows, [totals_row]) if totals_row else b""),
        "col_xf": dict((str(k), v) for k, v in sorted(col_xf.items())),
        "name_xf": name_xf,
        "day_columns": dict((str(k), v) for k, v in sorted(day_cols.items())),
        "total_col": total_col,
        "sst": sst,
        "first_row": 4,
        "data_rows": data_rows - 3,
        "totals_row": totals_row,
        "name_col": 0,
        "source_file": os.path.abspath(src),
        "tail": enc(tail),
    }
    json.dump(data, open(out, "w", encoding="utf-8"), ensure_ascii=False)
    print("cot so lieu:", [(k, v) for k, v in sorted(col_xf.items()) if k <= 12])
    print("ngay:", len(day_cols), "| cot TONG:", total_col,
          "| dong du lieu toi da:", data["data_rows"], "| dong tong:", totals_row)
    print("sst:", len(sst), "->", out)


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2] if len(sys.argv) > 2 else "src/template.json")
