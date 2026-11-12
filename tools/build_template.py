"""Tao khuon (template.json) tu file .xls mau - chay lai khi thay mau moi.

Khuon gom:
  globals      : toan bo stream truoc phan sheet (font, number format, bang chuoi)
  sheet_header : BOF..COLINFO cua sheet, giu nguyen do rong cot va dinh dang
  title_blob   : record cac o tieu de (hang 0..3, gom ten cot va ngay)
  totals_blob  : record dong tong (hang cuoi bang, co cong thuc)
  col_xf       : dinh dang (xf) dung cho tung cot so lieu
  sst          : bang chuoi (de tra ten mat hang co san)
  tail         : cac record ket thuc sheet
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
TE_IDS = (0x00FD, 0x0203, 0x00BD, 0x00BE, 0x0201, 0x0204, 0x0006, 0x04BC)
TITLE_ROWS = (0, 1, 2, 3)


def enc(blob):
    return base64.b64encode(blob).decode("ascii")


def xf_of(rid, body):
    if rid in (0x0203, 0x00FD, 0x0201) and len(body) >= 6:
        return struct.unpack("<H", body[4:6])[0]
    if rid == 0x00BD and len(body) >= 8:
        return struct.unpack("<H", body[4:6])[0]
    return None


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
    globals_blob = stream[:glob_end]
    sheet = stream[glob_end:]
    srecs = _xls._records(sheet)

    # header: truoc ROW dau tien
    pos = 0
    for rid, body in srecs:
        if rid == ROW:
            break
        pos += 4 + len(body)
    sheet_header = sheet[:pos]

    # gom record theo hang
    rows = {}
    for rid, body in srecs:
        if rid == ROW:
            r = struct.unpack("<H", body[:2])[0]
            rows.setdefault(r, []).append((rid, body))
        elif rid in TE_IDS and len(body) >= 2:
            r = struct.unpack("<H", body[:2])[0]
            rows.setdefault(r, []).append((rid, body))

    def blob_of(hrs):
        out = b""
        for hr in hrs:
            for rid, body in rows.get(hr, []):
                out += struct.pack("<HH", rid, len(body)) + body
        return out

    # dong tong: hang co cong thuc, khong co ten o cot A, nam cuoi
    totals_row = None
    for hr in sorted(rows):
        if hr < 4:
            continue
        has_formula = any(rid in (0x0006, 0x04BC) for rid, _b in rows[hr])
        has_name = any(rid == 0x00FD and len(b) >= 4 and struct.unpack("<H", b[2:4])[0] == 0
                       for rid, b in rows[hr])
        if has_formula and not has_name:
            totals_row = hr
    max_data_row = (totals_row - 1) if totals_row else max(rows)

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

    # dinh dang tung cot: uu tien dong du lieu dau tien co nhieu o nhat
    col_xf = {}
    best = None
    for hr in sorted(rows):
        if hr < 4 or hr >= max_data_row + 1:
            continue
        n = sum(1 for rid, _b in rows[hr] if rid in (0x00BD, 0x0203))
        if best is None or n > best[1]:
            best = (hr, n)
    if best:
        for rid, body in rows[best[0]]:
            if rid == 0x00BD:
                r0, c0 = struct.unpack("<HH", body[:4])
                k = (len(body) - 6) // 6
                for i in range(k):
                    xf = struct.unpack("<H", body[4 + i * 6:6 + i * 6])[0]
                    col_xf.setdefault(c0 + i, xf)
            elif rid == 0x0203 and len(body) >= 6:
                c, xf = struct.unpack("<HH", body[2:6])
                col_xf.setdefault(c, xf)
    # cot 'TONG' (2 cot cuoi) lay tu dong tong neu co
    if totals_row:
        for rid, body in rows[totals_row]:
            if rid == 0x00BD:
                r0, c0 = struct.unpack("<HH", body[:4])
                k = (len(body) - 6) // 6
                for i in range(k):
                    xf = struct.unpack("<H", body[4 + i * 6:6 + i * 6])[0]
                    col_xf[c0 + i] = xf

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

    # ngay cua tung cap cot: doc o tieu de hang 2 (nhan "Ngày xx-xx")
    sst_norm = dict((i, s) for i, s in enumerate(sst))
    day_cols = {}
    for rid, body in rows.get(2, []):
        if rid == 0x00FD and len(body) >= 10:
            _r, c, _xf, idx = struct.unpack("<HHHI", body[:10])
            txt = sst[idx] if idx < len(sst) else ""
            if "ngày" in txt.lower() or "ngay" in txt.lower():
                day_cols[str(c)] = txt
    day_cols_map = set(int(k) for k in day_cols)
    header_rows = {}
    for hr in TITLE_ROWS:
        for rid, body in rows.get(hr, []):
            if rid == ROW:
                header_rows[str(hr)] = base64.b64encode(body).decode("ascii")
    # dinh dang o ten mat hang (cot A cua dong du lieu)
    name_xf = 15
    if best:
        for rid, body in rows[best[0]]:
            if rid == 0x00FD and len(body) >= 10:
                _r, c, xf, _idx = struct.unpack("<HHHI", body[:10])
                if c == 0:
                    name_xf = xf
    # cot TONG: cap cot sau khi het cot ngay (lay tu dong tieu de)
    day_sort = sorted(int(k) for k in day_cols)
    total_start = (day_sort[-1] + 2) if day_sort else None

    # cot TONG thuc te: cot co cong thuc trong mot dong du lieu bat ky
    total_col = None
    for hr in sorted(rows):
        if hr < 4 or hr > max_data_row:
            continue
        for rid, body in rows[hr]:
            if rid == 0x0006 and len(body) >= 4:
                c = struct.unpack("<H", body[2:4])[0]
                if c not in day_cols_map:
                    total_col = c
                    break
        if total_col is not None:
            break
    if total_col is None:
        total_col = total_start
    # dinh dang (xf) cua cap cot TONG: lay tu record cong thuc
    if total_col is not None:
        for hr in sorted(rows):
            for rid, body in rows[hr]:
                if rid == 0x0006 and len(body) >= 6:
                    c, xf = struct.unpack("<HH", body[2:6])
                    if c in (total_col, total_col + 1):
                        col_xf[c] = xf

    data = {
        "globals": enc(globals_blob),
        "name_xf": name_xf,
        "total_col": total_col,
        "day_columns": day_cols,
        "header_rows": header_rows,
        "sheet_header": enc(sheet_header),
        "title_blob": enc(blob_of(TITLE_ROWS)),
        "totals_blob": enc(blob_of([totals_row]) if totals_row else b""),
        "col_xf": dict((str(k), v) for k, v in sorted(col_xf.items())),
        "sst": sst,
        "first_row": 4,
        "data_rows": max_data_row - 3,
        "totals_row": totals_row,
        "totals_rel": (totals_row - 4) if totals_row else 27,
        "name_col": 0,
        "source_file": os.path.abspath(src),
        "tail": enc(tail),
    }
    json.dump(data, open(out, "w", encoding="utf-8"), ensure_ascii=False)
    print("cot:", sorted(col_xf.items()))
    print("dong du lieu toi da:", data["data_rows"], "| dong tong:", totals_row)
    print("sst:", len(sst), "->", out)


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2] if len(sys.argv) > 2 else "src/template.json")
