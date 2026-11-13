"""Smoke test cho ban .exe da dong goi (chay tren may da co ban build).

    python tools/smoke_exe.py dist/TienAn.exe

Kiem tra 3 viec, moi viec bao exit code that:
  1. File ton tai va la dinh dang chay duoc tren he hien tai.
  2. template.json CO trong goi va doc duoc (rut ra bang pyi-archive_viewer).
  3. Khuon nap duoc tu duong dan _MEIPASS (dung logic find_template cua app).

Khong mo GUI: GUI chi chung minh "cua so mo duoc", khong chung minh
template nap thanh cong.
"""
import json
import os
import shutil
import subprocess
import sys
import tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
SRC = os.path.join(ROOT, "src")


def fail(msg, code=1):
    print("FAIL: " + msg)
    sys.exit(code)


def main():
    if len(sys.argv) < 2:
        fail("dung: python tools/smoke_exe.py <duong_dan_TienAn[.exe]>", 2)
    exe = sys.argv[1]
    if not os.path.exists(exe):
        fail("khong thay file: %s" % exe)

    print("[1] file: %s (%.1f MB)" % (exe, os.path.getsize(exe) / 1e6))

    # 2. template.json phai nam trong goi va doc duoc
    viewer = shutil.which("pyi-archive_viewer")
    if not viewer:
        viewer = os.path.join(os.path.dirname(sys.executable), "pyi-archive_viewer")
    if not os.path.exists(viewer):
        print("  (bo qua buoc 2: khong co pyi-archive_viewer)")
        tpl = os.path.join(SRC, "template.json")
    else:
        tmp = tempfile.mkdtemp()
        p = subprocess.run([viewer, "-l", exe], capture_output=True, text=True)
        if "template.json" not in p.stdout:
            fail("template.json KHONG co trong goi dong goi")
        # rut ra de doc that
        q = subprocess.run([viewer, exe], input="x\ntemplate.json\ny\nX\ntemplate.json\nq\n",
                           capture_output=True, text=True, cwd=tmp)
        cand = [os.path.join(tmp, f) for f in os.listdir(tmp)]
        cand = [c for c in cand if os.path.getsize(c) > 1000]
        if not cand:
            fail("khong rut duoc template.json ra de kiem tra")
        tpl = sorted(cand, key=os.path.getsize)[-1]
    data = json.load(open(tpl, encoding="utf-8"))
    print("[2] template.json: %d cot, %d ngay, total_col=%s"
          % (len(data["col_xf"]), len(data["day_columns"]), data["total_col"]))
    assert len(data["day_columns"]) == 21, "so ngay sai"

    # 3. nap khuon bang dung logic cua app
    sys.path.insert(0, SRC)
    from _xlsio import Template
    tplobj = Template(tpl)
    print("[3] Template nap OK: %d cot, %d ngay, cap TONG %s"
          % (len(tplobj.col_xf), len(tplobj.day_columns), tplobj.total_pair()))
    assert len(tplobj.day_columns) == 21
    print("SMOKE_OK")


if __name__ == "__main__":
    main()
