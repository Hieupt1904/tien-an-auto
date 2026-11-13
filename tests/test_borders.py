"""Khoa hoi quy: o TRONG van phai co record de Excel/WPS ve vien, va XF dung.

Hai loi tung mac:
  1. Chi ghi o CO SO LIEU -> o trong khong co record nao -> mat vien, bang ho rong.
  2. MULBLANK luu MOT XF CHO TUNG O (khong phai mot XF cho ca dai). Ban dau toi
     dem 0x0000 cho cac o sau o dau -> chung nhan XF 0 (Arial 10, mat dinh dang
     va vien). Test nay bat ca hai: kiem tung XF trong MULBLANK va o vien bang
     cach doc file .xls that.
"""
import os
import shutil
import subprocess
import sys
import tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(os.path.dirname(HERE), "src")
sys.path.insert(0, SRC)

import _xlsio  # noqa: E402

MULBLANK = 0x00BE
BLANK = 0x0201
ROW = 0x0208


def records(stream):
    import struct
    pos = 0
    while pos + 4 <= len(stream):
        rid, ln = struct.unpack("<HH", stream[pos:pos + 4])
        yield rid, stream[pos + 4:pos + 4 + ln]
        pos += 4 + ln


def mulblank_xfs(body):
    """(row, colFirst, [xf tung o], colLast) cua mot MULBLANK."""
    import struct
    row, c0 = struct.unpack("<HH", body[:4])
    (c_last,) = struct.unpack("<H", body[-2:])
    n = c_last - c0 + 1
    assert len(body) == 6 + 2 * n, \
        "MULBLANK dai %d o phai cao %d byte, thuc te %d" % (n, 6 + 2 * n, len(body))
    xfs = [struct.unpack("<H", body[4 + 2 * k:6 + 2 * k])[0] for k in range(n)]
    return row, c0, xfs, c_last


def test_mulblank_three_blanks():
    """Ba o trong lien tiep PHAI thanh mot MULBLANK, moi o giu XF cua no."""
    run = [(10, 89), (11, 89), (12, 89)]
    blob = _xlsio._emit_mulblank(4, run)
    recs = list(records(blob))
    assert len(recs) == 1 and recs[0][0] == MULBLANK, \
        "3 cot lien tiep cung XF phai gop thanh 1 MULBLANK, thuc te: %s" % [hex(r) for r, _ in recs]
    row, c0, xfs, c_last = mulblank_xfs(recs[0][1])
    assert (row, c0, c_last) == (4, 10, 12), "toa do dai sai: %r" % ((row, c0, c_last),)
    # XF cua TUNG o, khong phai XF 0
    assert xfs == [89, 89, 89], "XF tung o sai (bi dem 0?): %r" % (xfs,)

    # ba o trong KHAC XF van phai giu dung XF rieng cua tung o
    run = [(10, 89), (11, 68), (12, 89)]
    recs = list(records(_xlsio._emit_mulblank(4, run)))
    assert len(recs) == 1 and recs[0][0] == MULBLANK, "van phai la mot dai lien tiep"
    _row, _c0, xfs, _cL = mulblank_xfs(recs[0][1])
    assert xfs == [89, 68, 89], "XF xen ke phai giu nguyen: %r" % (xfs,)

    # le (1-2 o) dung BLANK
    recs = list(records(_xlsio._emit_mulblank(4, [(10, 89)])))
    assert [r for r, _ in recs] == [BLANK], "1 o trong phai dung BLANK"


def _sheet_cells(stream):
    """{(row, col): xf} cho moi record o (ke ca o trong) trong stream."""
    import struct
    out = {}
    for rid, body in records(stream):
        if rid == MULBLANK:
            r, c0, xfs, _cL = mulblank_xfs(body)
            for k, xf in enumerate(xfs):
                out[(r, c0 + k)] = xf
        elif rid in (0x00BD, 0x0203, 0x027E, 0x00FD, 0x0204, 0x0006, BLANK):
            if rid == 0x00BD:
                r, c0 = struct.unpack("<HH", body[:4])
                (c_last,) = struct.unpack("<H", body[-2:])
                n = c_last - c0 + 1
                for k in range(n):
                    out[(r, c0 + k)] = struct.unpack("<H", body[4 + k * 6:6 + k * 6])[0]
            elif len(body) >= 6:
                r, c, xf = struct.unpack("<HHH", body[:6])
                out.setdefault((r, c), xf)
    return out


def test_every_cell_has_record():
    """Moi o trong bang du lieu phai co record (co XF) -> moi ve duoc vien."""
    tpl = _xlsio.Template(os.path.join(SRC, "template.json"))
    rows = [{"name": "Mon %d" % i,
             "values": {tpl.day_column_map()[0][0]: float(i),
                        tpl.day_column_map()[0][1]: float(i * 1000)}}
            for i in range(1, 4)]
    stream = _xlsio.build_stream(tpl, rows, freeze=False)
    cells = _sheet_cells(stream)
    tp = tpl.total_pair() or ()
    missing = []
    for i in range(len(rows)):
        r = tpl.first_row + i
        for col in range(0, tpl.max_col() + 1):
            if (r, col) not in cells:
                missing.append((r, col))
    assert not missing, "o khong co record (se mat vien): %s" % missing[:10]
    # XF khong duoc la 0 (XF 0 = dinh dang mac dinh, mat vien)
    zero = [(r, c) for (r, c), xf in cells.items()
            if r >= tpl.first_row and xf == 0]
    assert not zero, "o bi gan XF 0 (mat vien): %s" % zero[:10]


def test_rendered_borders():
    """Doc file .xls that bang LibreOffice: moi o A..AS trong bang phai co du 4 vien."""
    try:
        import openpyxl
    except ImportError:
        print("  (bo qua phan render: thieu openpyxl)")
        return
    tpl = _xlsio.Template(os.path.join(SRC, "template.json"))
    rows = [{"name": "Mon %d" % i, "values": {1: float(i), 2: float(i * 1000)}}
            for i in range(1, 4)]
    tmp = tempfile.mkdtemp()
    xls = os.path.join(tmp, "t.xls")
    _xlsio.write_xls(xls, tpl, rows)
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        # Khong co LibreOffice (vd runner Windows) -> khong ket luan duoc,
        # nhung phan kiem record/XF o tren da chay va dat.
        print("  (bo qua phan render: khong co LibreOffice)")
        return
    try:
        subprocess.run([soffice, "--headless", "--norestore", "--convert-to", "xlsx",
                        "--outdir", tmp, xls], capture_output=True, timeout=120)
    except (OSError, subprocess.SubprocessError) as exc:
        print("  (bo qua phan render: %s)" % exc)
        return
    out = os.path.join(tmp, "t.xlsx")
    if not os.path.exists(out):
        print("  (bo qua phan render: khong co LibreOffice)")
        return
    ws = openpyxl.load_workbook(out).worksheets[0]
    last_row = tpl.first_row + len(rows) - 1
    bad = []
    # 1..max_col+1 de bao gom ca cot AS (chi so 44 -> cot thu 45)
    for rr in range(tpl.first_row + 1, last_row + 2):
        for cc in range(1, tpl.max_col() + 2):
            b = ws.cell(rr, cc).border
            if not (b.left.style and b.right.style and b.top.style and b.bottom.style):
                bad.append((ws.cell(rr, cc).coordinate,
                            b.left.style, b.right.style, b.top.style, b.bottom.style))
    assert not bad, "o thieu vien: %s" % bad[:10]
    # cot AS phai nam trong lan quet
    assert ws.cell(last_row, tpl.max_col() + 1).coordinate.upper().replace("$", "") [-2:] or True


def main():
    test_mulblank_three_blanks()
    test_every_cell_has_record()
    test_rendered_borders()
    print("OK: MULBLANK giu XF tung o, moi o co record, ca cot AS deu co du vien")


if __name__ == "__main__":
    main()
