"""Khoa hoi quy: src/template.json phai dung y nguyen khi sinh lai tu file mau.

Vi sao can: bang XF (dinh dang o) trong template.json tung bi sua tay de chu so
hien dung, va tools/build_template.py khong sinh lai duoc cac gia tri do. Lan
sau ai chay lai build_template la dinh dang lech quay ve ma khong ai biet.

Bai nay chay lai build_template tu file mau va so:
  * col_xf (dinh dang tung cot)
  * day_columns / total_col / name_xf / sst
Neu build_template bi doi ma khong cap nhat template.json (hoac nguoc lai),
bai se truot va noi ro cot nao lech.
"""
import json
import os
import subprocess
import sys
import tempfile

for _s in (sys.stdout, sys.stderr):
    try:
        _s.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
SRC = os.path.join(ROOT, "src")
sys.path.insert(0, SRC)

TEMPLATE = os.path.join(SRC, "template.json")
BUILDER = os.path.join(ROOT, "tools", "build_template.py")


def _invariants(cur):
    """Cac bat bien doc lap voi file mau: kiem tra TRUOC khi co the bo qua."""
    cx = cur["col_xf"]
    days = sorted(int(k) for k in cur["day_columns"])
    assert days, "khuon khong co cot ngay"
    qty, amt = cx.get("1"), cx.get("2")
    assert qty is not None, "thieu XF cho cot so luong (cot 1)"
    assert amt is not None, "thieu XF cho cot thanh tien (cot 2)"
    assert qty != amt, "cot so luong va thanh tien khong duoc cung mot XF"
    for dcol in days:
        assert str(dcol) in cx, "cot so luong %d thieu XF" % dcol
        assert str(dcol + 1) in cx, "cot thanh tien %d thieu XF" % (dcol + 1)
        assert cx[str(dcol)] == qty, \
            "cot so luong %d dung XF %s, phai la %s" % (dcol, cx[str(dcol)], qty)
        assert cx[str(dcol + 1)] == amt, \
            "cot thanh tien %d dung XF %s, phai la %s" % (dcol + 1, cx[str(dcol + 1)], amt)
    assert all(v != 0 for v in cx.values()), \
        "co cot bi gan XF 0 (mat dinh dang/vien)"
    tp = cur.get("total_col")
    assert tp is not None, "thieu cot TONG"
    assert str(tp) in cx and str(tp + 1) in cx, "cap cot TONG thieu XF"
    assert cur.get("first_row") is not None and cur.get("totals_row") is not None
    return days


def main():
    if not os.path.exists(TEMPLATE):
        raise AssertionError("thieu src/template.json")
    cur = json.load(open(TEMPLATE, encoding="utf-8"))

    # (1) LUON kiem bat bien doc lap - khong phu thuoc file mau
    days = _invariants(cur)

    # (2) tai sinh tu file mau (chi khi co file mau)
    src_file = cur.get("source_file") or ""
    if not src_file or not os.path.exists(src_file):
        print("CANH BAO: KHONG kiem duoc tai sinh - thieu file mau goc.")
        print("  source_file = %s" % (src_file or "(trong)"))
        print("  Da kiem xong bat bien doc lap: %d cot, %d ngay, XF le/chan dung."
              % (len(cur["col_xf"]), len(days)))
        print("  Muon kiem day du: dat file mau goc dung duong dan tren roi chay lai.")
        return

    out = os.path.join(tempfile.mkdtemp(), "template.json")
    r = subprocess.run([sys.executable, BUILDER, src_file, out],
                       capture_output=True, text=True)
    assert r.returncode == 0, "build_template loi:\n%s" % (r.stderr or r.stdout)
    new = json.load(open(out, encoding="utf-8"))

    keys = set(cur["col_xf"]) | set(new["col_xf"])
    diff = [(int(k), cur["col_xf"].get(k), new["col_xf"].get(k))
            for k in keys if cur["col_xf"].get(k) != new["col_xf"].get(k)]
    diff.sort()
    assert not diff, (
        "template.json lech voi build_template o dinh dang cot (cot, dang dung, sinh lai):\n"
        + "\n".join("  cot %d: %s vs %s" % d for d in diff[:20])
        + ("\n  ... va %d cot nua" % (len(diff) - 20) if len(diff) > 20 else "")
        + "\n-> chay lai: python tools/build_template.py \"%s\" src/template.json"
        % src_file)

    for key in ("day_columns", "total_col", "name_xf", "first_row", "totals_row",
                "name_col", "title_blob_has_rows"):
        assert cur.get(key) == new.get(key), \
            "%s lech: %r vs %r" % (key, cur.get(key), new.get(key))
    assert cur.get("sst") == new.get("sst"), "bang chuoi (sst) lech"

    print("OK: template.json dung y nguyen khi sinh lai tu file mau "
          "(%d cot, %d ngay)" % (len(cur["col_xf"]), len(days)))


if __name__ == "__main__":
    main()
