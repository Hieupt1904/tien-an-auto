"""Khoa hoi quy: cau truc BIFF8 phai dung chuan Excel, khong chi "LibreOffice doc duoc".

LibreOffice de tinh; Excel thi bao "File error: data may have been lost" va
thoat ra khi gap 2 loi sau (da tung mac):
  1. ROW.colMac = 0 trong khi hang co o -> Excel khong biet hang rong den dau.
     BIFF8: rw(2), colMic(2), colMac(2) voi colMac la SO COT, tuc last_col + 1.
  2. Hai record DIMENSIONS trong cung sheet -> Excel tu choi file.
"""
import os
import struct
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(os.path.dirname(HERE), "src")
sys.path.insert(0, SRC)

import _xlsio  # noqa: E402

ROW = 0x0208
DIMENSIONS = 0x0200
CELL_IDS = (0x00BD, 0x00BE, 0x0201, 0x0203, 0x00FD, 0x0204, 0x0006, 0x027E)


def walk(stream):
    pos = 0
    while pos + 4 <= len(stream):
        rid, ln = struct.unpack("<HH", stream[pos:pos + 4])
        yield rid, stream[pos + 4:pos + 4 + ln]
        pos += 4 + ln


def cell_ref(rid, body):
    r = struct.unpack("<H", body[:2])[0]
    if rid == 0x00BD:                       # MULRK: colLast o 2 byte cuoi
        c = struct.unpack("<H", body[-2:])[0]
        c0 = struct.unpack("<H", body[2:4])[0]
        return r, c0, c
    if rid == 0x00BE:                       # MULBLANK
        return r, struct.unpack("<H", body[2:4])[0], struct.unpack("<H", body[-2:])[0]
    return r, struct.unpack("<H", body[2:4])[0], struct.unpack("<H", body[2:4])[0]


def main():
    tpl = _xlsio.Template(os.path.join(SRC, "template.json"))
    # nhieu so dong khac nhau de bat loi o ca bang ngan va dai
    for n in (1, 27, 31):
        rows = [{"name": "M%d" % i, "values": {1: float(i), 2: float(i * 1000)}}
                for i in range(1, n + 1)]
        stream = _xlsio.build_stream(tpl, rows, freeze=False)

        dims = []
        colmac = {}
        bad = []
        max_row = 0
        max_col = 0
        for rid, body in walk(stream):
            if rid == DIMENSIONS:
                dims.append(struct.unpack("<IIHHH", body[:14]))
            elif rid == ROW:
                r, _cmic, cmac = struct.unpack("<HHH", body[:6])
                colmac[r] = cmac
            elif rid in CELL_IDS and len(body) >= 6:
                r, c0, c1 = cell_ref(rid, body)
                max_row = max(max_row, r)
                max_col = max(max_col, c1)
                cmac = colmac.get(r)
                assert cmac is not None, "o (%d,%d) khong co ROW header" % (r, c0)
                if c1 >= cmac:
                    bad.append((r, c0, c1, cmac))

        # 1. dung MOT DIMENSIONS
        assert len(dims) == 1, "%d dong: co %d record DIMENSIONS (phai la 1)" % (n, len(dims))

        # 2. moi o nam trong colMac cua hang (Excel kiem cai nay)
        assert not bad, \
            "%d dong: o nam ngoai colMac (rw, c0, c1, colMac): %s" % (n, bad[:5])

        # 3. DIMENSIONS phai bao dung vung co du lieu
        r1, r2, c1d, c2d, _ = dims[0]
        assert r2 >= max_row + 1, \
            "%d dong: DIMENSIONS rwMic=%d < dong cuoi %d" % (n, r2, max_row)
        assert c2d >= max_col + 1, \
            "%d dong: DIMENSIONS colMic=%d < cot cuoi %d" % (n, c2d, max_col)

    print("OK: BIFF8 dung chuan Excel (1 DIMENSIONS, colMac phu moi o, "
          "vung DIMENSIONS dung)")


if __name__ == "__main__":
    main()
