"""Doi chieu file quyet toan voi so lieu API (chay khi co mang).

    python tools/verify_report.py "Bao cao/Quyet toan tien an Thang 11-2026.xls"

Voi tung cot ngay trong file, goi lai API cua ngay do va so:
  * danh sach mat hang / so luong / thanh tien tung dong
  * tong tien cua cot
Bao "OK" hoac liet ke tung o lech.
"""
import argparse
import datetime
import os
import re
import subprocess
import sys
import tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(os.path.dirname(HERE), "src")
sys.path.insert(0, SRC)

import _api       # noqa: E402
import _pull      # noqa: E402
import _session   # noqa: E402


def to_xlsx(path):
    out = tempfile.mkdtemp()
    subprocess.run(["soffice", "--headless", "--norestore", "--convert-to", "xlsx",
                    "--outdir", out, path], capture_output=True)
    base = os.path.splitext(os.path.basename(path))[0] + ".xlsx"
    full = os.path.join(out, base)
    if not os.path.exists(full):
        raise SystemExit("Khong chuyen duoc sang xlsx (can LibreOffice)")
    return full


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("report")
    ap.add_argument("--user", default=None)
    ap.add_argument("--password", default=None)
    ap.add_argument("--warehouse", default="2")
    args = ap.parse_args()

    import openpyxl
    import json
    cfg_path = os.path.join(os.path.dirname(HERE), "config.json")
    cfg = {}
    if os.path.exists(cfg_path):
        cfg = json.load(open(cfg_path, encoding="utf-8"))
    user = args.user or cfg.get("username")
    pwd = args.password or cfg.get("password")
    if not user or not pwd:
        raise SystemExit("Thieu tai khoan: dung --user/--password hoac config.json")

    sh = openpyxl.load_workbook(to_xlsx(args.report)).active
    header = {}
    for c in range(1, sh.max_column + 1):
        v = sh.cell(row=3, column=c).value
        if isinstance(v, str) and v.startswith("Ngày"):
            m = re.search(r"(\d{1,2})[-/](\d{1,2})", v)
            if m:
                header[c] = (int(m.group(1)), int(m.group(2)))
    import _report as _rep

    def matching_rows(name):
        """Dong co cung ten sau khi bo dau (phan biet hoa/thuong, giu dau cau)."""
        k = _rep.key(name)
        return [r for n, r in names.items() if _rep.key(n) == k]

    names = {}
    for r in range(5, sh.max_row + 1):
        n = sh.cell(row=r, column=1).value
        if n:
            names[str(n)] = r
    print("File: %s | %d dong | %d cot ngay" % (args.report, len(names), len(header)))

    api = _api.TienAnApi(user, pwd)
    saved = _session.load(os.path.join(os.path.dirname(HERE), "session.json"), user)
    if saved:
        _session.apply(api, saved)
    try:
        api.menu_report(1)
    except Exception:
        api.login(max_attempts=80)

    year = None
    for (d, m) in sorted(header.values()):
        year = year or datetime.date.today().year
    bad_total = 0
    bad_cells = 0
    for col, (d, m) in sorted(header.items()):
        date_text = "%02d/%02d/%04d" % (d, m, year)
        try:
            day = _pull.extract_day(api.marketbill(date_text, args.warehouse))
        except Exception as exc:
            print("  %s  khong lay duoc: %s" % (date_text, exc))
            continue
        col_sum = 0.0
        for name, info in day["foods"].items():
            cand = matching_rows(name)
            r = cand[0] if len(cand) == 1 else None
            if r is None:
                print("  %s  thieu dong: %r" % (date_text, name))
                bad_cells += 1
                continue
            q = sh.cell(row=r, column=col).value
            a = sh.cell(row=r, column=col + 1).value
            if q is None or abs(float(q) - info["qty"]) > 1e-6 or \
                    abs(float(a) - info["amount"]) > 1e-6:
                print("  %s  lech: %r file=(%s, %s) api=(%s, %s)"
                      % (date_text, name, q, a, info["qty"], info["amount"]))
                bad_cells += 1
            col_sum += info["amount"]
        file_sum = sum(v for v in (sh.cell(row=r, column=col + 1).value
                                   for r in range(5, sh.max_row + 1))
                       if isinstance(v, (int, float)))
        mark = "OK" if abs(file_sum - col_sum) < 1 else "LECH"
        if mark != "OK":
            bad_total += 1
        print("  %s  file=%12d  api=%12d  %s" % (date_text, int(file_sum), int(col_sum), mark))
    print("Ket qua: %d cot lech tong, %d o lech" % (bad_total, bad_cells))
    return 0 if bad_total == 0 and bad_cells == 0 else 1


if __name__ == "__main__":
    sys.exit(main())
