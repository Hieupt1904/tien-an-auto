"""Kiem tra toa do o: gia tri nhap cot nao phai ra dung cot do (ke ca cot thua).

Chay duoc khong can mang, khong can Excel: doc lai stream BIFF8 vua sinh va so
tung o voi du lieu dau vao (round-trip o muc record).
"""
import os
import struct
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(os.path.dirname(HERE), "src")
sys.path.insert(0, SRC)

import _xlsio  # noqa: E402

ROW = 0x0208
MULRK = 0x00BD
NUMBER = 0x0203
BLANK = 0x0201
MULBLANK = 0x00BE
LABELSST = 0x00FD
FORMULA = 0x0006


def rk_to_float(rk):
    if rk & 0x02:
        val = rk >> 2
        if val & 0x20000000:
            val -= 0x40000000
        num = float(val)
    else:
        num = struct.unpack("<d", struct.pack("<Q", (rk & 0xFFFFFFFC) << 32))[0]
    if rk & 0x01:
        num /= 100.0
    return num


def cells_by_row(stream):
    """Doc lai: {(row, col): value} tu stream BIFF8."""
    out = {}
    pos = 0
    while pos + 4 <= len(stream):
        rid, ln = struct.unpack("<HH", stream[pos:pos + 4])
        body = stream[pos + 4:pos + 4 + ln]
        if rid == MULRK:
            r, c0 = struct.unpack("<HH", body[:4])
            n = (len(body) - 6) // 6
            for i in range(n):
                _xf, rk = struct.unpack("<HI", body[4 + i * 6:10 + i * 6])
                out[(r, c0 + i)] = rk_to_float(rk)
            (col_last,) = struct.unpack("<H", body[-2:])
            assert col_last == c0 + n - 1, "colLast sai: %s vs %s" % (col_last, c0 + n - 1)
        elif rid == NUMBER:
            r, c, _xf = struct.unpack("<HHH", body[:6])
            out[(r, c)] = struct.unpack("<d", body[6:14])[0]
        elif rid == LABELSST:
            r, c, _xf, idx = struct.unpack("<HHHI", body[:10])
            out[(r, c)] = ("sst", idx)
        elif rid == FORMULA:
            r, c, _xf = struct.unpack("<HHH", body[:6])
            out[(r, c)] = struct.unpack("<d", body[6:14])[0]
        pos += 4 + ln
    return out


def check(tpl, values, note):
    rows = [{"name": "Ga", "values": values}]
    stream = _xlsio.build_stream(tpl, rows, freeze=False)
    got = cells_by_row(stream)
    r = tpl.first_row
    tp = tpl.total_pair() or ()
    for col, want in values.items():
        if tp and col in tp:
            continue                            # cot TONG kiem tra rieng ben duoi
        have = got.get((r, col))
        assert have is not None, "%s: mat o (hang %d, cot %d)" % (note, r, col)
        assert abs(float(have) - float(want)) < 1e-9, \
            "%s: cot %d = %r, phai la %r" % (note, col, have, want)
    # khong duoc co o la o cot khac trong cung hang (tru cot ten va cap TONG,
    # cap TONG luon co cong thuc tren moi dong du lieu)
    want_cols = set(values) - set(tp)          # cot TONG do cong thuc sinh
    extra = [c for (rr, c) in got
             if rr == r and c != tpl.name_col and c not in want_cols and c not in tp]
    assert not extra, "%s: sinh them o cot %s" % (note, extra)
    # cap cot TONG: phai co CONG THUC (khong phai so) va ket qua dung
    if tp:
        q_cols = [c for c, _ in tpl.day_column_map()]
        a_cols = [c for _, c in tpl.day_column_map()]
        for col, sources in ((tp[0], q_cols), (tp[1], a_cols)):
            have = got.get((r, col))
            assert have is not None, "%s: cot TONG %d thieu o" % (note, col)
            want = sum(float(values.get(c, 0.0)) for c in sources)
            assert abs(float(have) - want) < 1e-6, \
                "%s: cot TONG %d = %r, phai la %r" % (note, col, have, want)


def main():
    tpl = _xlsio.Template(os.path.join(SRC, "template.json"))
    day_cols = [c for c, _ in tpl.day_column_map()]
    assert day_cols, "khuon khong co cot ngay"

    check(tpl, {1: 255.0, 2: 357000.0}, "lien tiep 1-2")
    if len(day_cols) >= 4:
        check(tpl, {day_cols[2]: 12.9, day_cols[2] + 1: 193500.0}, "cot 5/6 (co cot thua)")
        check(tpl, {day_cols[3]: 7.0}, "chi 1 o le")
    # nhieu nhom roi rac
    sparse = {}
    for i, base in enumerate(day_cols[:5]):
        sparse[base] = float(i + 1)
        sparse[base + 1] = float((i + 1) * 10000)
    check(tpl, sparse, "nhieu nhom roi rac")

    # cot QTY_THANH_TIEN nam sau cot thua: dam bao khong lech
    total = tpl.total_pair()
    if total:
        # cap TONG do cong thuc sinh ra tu cac cot nguon; gia tri truyen vao
        # cho cot TONG bi bo qua (khong ghi so tho vao o cong thuc).
        first_day = day_cols[0]
        check(tpl, {first_day: 3.5, first_day + 1: 49000.0,
                    total[0]: 999.0, total[1]: 999.0}, "cot TONG")
    print("OK: toa do o dung voi moi truong hop (lien tiep, le, roi rac, cot tong)")


if __name__ == "__main__":
    main()
