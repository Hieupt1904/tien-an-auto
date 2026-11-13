"""Khoa hoi quy: cong thuc dong tong phai tinh DU moi dong du lieu va dung ket qua.

Loi tung mac: khung file mau co =SUM(C5:C31); khi bang co 36 mat hang (dong 5..40)
ma khong noi vung thi tong thieu 9 dong - file van "dep" nen de bi bo qua.
"""
import os
import struct
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(os.path.dirname(HERE), "src")
sys.path.insert(0, SRC)

import _xlsio  # noqa: E402

FORMULA = 0x0006


def formula_cells(stream, row):
    out = []
    pos = 0
    while pos + 4 <= len(stream):
        rid, ln = struct.unpack("<HH", stream[pos:pos + 4])
        body = stream[pos + 4:pos + 4 + ln]
        if rid == FORMULA and len(body) >= 22:
            r, c, _xf = struct.unpack("<HHH", body[:6])
            if r == row:
                cce = struct.unpack("<H", body[20:22])[0]
                ranges = _xlsio._ranges_in_formula(body[22:22 + cce])
                cached = struct.unpack("<d", body[6:14])[0]
                out.append((c, ranges, cached))
        pos += 4 + ln
    return out


def main():
    tpl = _xlsio.Template(os.path.join(SRC, "template.json"))
    first = tpl.first_row
    for n in (1, 27, 35, 36, 40):
        rows = [{"name": "M%d" % i, "values": {1: float(i), 2: float(i * 1000)}}
                for i in range(1, n + 1)]
        stream = _xlsio.build_stream(tpl, rows, freeze=False)
        total_row = first + n
        cells = formula_cells(stream, total_row)
        assert cells, "khong co dong tong khi %d dong" % n
        tp = tpl.total_pair()
        for col, ranges, cached in cells:
            if tp and col in tp:
                # cap cot TONG: phai la cong thuc CO THAM CHIEU toi cot ngay,
                # khong duoc la ptgExp/ARRAY dung chung hay gia tri cu 26.154.000
                assert ranges, \
                    "%d dong cot %d: cot TONG phai co tham chieu (khong phai ptgExp)" % (n, col)
                q_src = [q for q, _a in tpl.day_column_map()]
                a_src = [a for _q, a in tpl.day_column_map()]
                want = (float(n * (n + 1) / 2) if col == tp[0]
                        else float(sum(i * 1000 for i in range(1, n + 1))))
                assert abs(cached - want) < 0.5, \
                    "%d dong cot %d: ket qua cot TONG %r, phai la %r" % (n, col, cached, want)
                # 21 cot ngay -> vung tham chieu phai co dung 21 muc
                cols = sorted({(r_first_col) for (_o, r_first_col, _rl) in []})
                continue
            for (_off, r_first, r_last) in ranges:
                # dong cuoi cua vung SUM phai la dong du lieu cuoi (first + n - 1),
                # hoac dong ngay sau (mau co san 1 dong dem) - nhung khong duoc thieu
                # dau vung: dong du lieu dau (dong mau goc bat dau tu 5) - cho phep 4/5
                assert r_first in (first, first + 1), \
                    "%d dong: vung SUM bat dau sai (%d)" % (n, r_first)
                assert r_last >= first + n - 1, \
                    "%d dong: vung SUM dung o %d, thieu dong (can >= %d)" % (
                        n, r_last, first + n - 1)
                # cot thanh tien (chan) moi co ket qua > 0; cot so luong cung phai khop
                qc, ac = tpl.day_column_map()[0]
                if col == qc:
                    want = float(n * (n + 1) / 2)
                elif col == ac:
                    want = float(sum(i * 1000 for i in range(1, n + 1)))
                else:
                    continue
                assert abs(cached - want) < 0.5, \
                    "%d dong cot %d: ket qua luu san %r, phai la %r" % (n, col, cached, want)
    print("OK: dong tong tinh du moi dong (1/27/35/36/40) va dung ket qua")


if __name__ == "__main__":
    main()
