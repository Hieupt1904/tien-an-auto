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

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
SRC = os.path.join(ROOT, "src")
sys.path.insert(0, SRC)

TEMPLATE = os.path.join(SRC, "template.json")
BUILDER = os.path.join(ROOT, "tools", "build_template.py")


def main():
    if not os.path.exists(TEMPLATE):
        print("BO QUA: chua co template.json")
        return
    cur = json.load(open(TEMPLATE, encoding="utf-8"))
    src_file = cur.get("source_file") or ""
    if not src_file or not os.path.exists(src_file):
        print("BO QUA: khong tim thay file mau goc (%s)" % (src_file or "?"))
        return

    out = os.path.join(tempfile.mkdtemp(), "template.json")
    r = subprocess.run([sys.executable, BUILDER, src_file, out],
                       capture_output=True, text=True)
    assert r.returncode == 0, "build_template loi:\n%s" % (r.stderr or r.stdout)
    new = json.load(open(out, encoding="utf-8"))

    # 1. dinh dang tung cot
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

    # 2. cac thong tin khac phai khop
    for key in ("day_columns", "total_col", "name_xf", "first_row", "totals_row",
                "name_col", "title_blob_has_rows"):
        assert cur.get(key) == new.get(key), \
            "%s lech: %r vs %r" % (key, cur.get(key), new.get(key))
    assert cur.get("sst") == new.get("sst"), "bang chuoi (sst) lech"

    # 3. dinh dang cot ngay phai nhat quan: le = so luong, chan = thanh tien
    cx = cur["col_xf"]
    qty, amt = cx.get("1"), cx.get("2")
    assert qty is not None and amt is not None, "thieu XF cho cot so luong / thanh tien"
    for dcol in sorted(int(k) for k in cur["day_columns"]):
        assert cx.get(str(dcol)) == qty, \
            "cot so luong %d dung XF %s, phai la %s" % (dcol, cx.get(str(dcol)), qty)
        assert cx.get(str(dcol + 1)) == amt, \
            "cot thanh tien %d dung XF %s, phai la %s" % (dcol + 1, cx.get(str(dcol + 1)), amt)
    print("OK: template.json dung y nguyen khi sinh lai tu file mau "
          "(%d cot, %d ngay)" % (len(cx), len(cur["day_columns"])))


if __name__ == "__main__":
    main()
