"""Khoa hoi quy: xuat thang KHAC thang cua khuon (vd khuon thang 11, chon thang 10).

Loi tung mac: khuon mau chi co cot danh so thang 11 ("Ngay 02-11" ... "Ngay 30/11").
Chon thang 10 thi khong cot nao khop -> toan bo du lieu bi bo, file ra rong.
"""
import datetime
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(os.path.dirname(HERE), "src")
sys.path.insert(0, SRC)

import _report  # noqa: E402
import _xlsio   # noqa: E402


class Pull:
    def __init__(self, days):
        self.days = days
        self.warnings = []


def make_days(month, count):
    return [{"date": datetime.date(2026, month, d),
             "foods": {"Ga": {"qty": float(d), "price": 10000.0,
                              "amount": float(d) * 10000.0}}}
            for d in range(1, count + 1)]


def main():
    tpl = _xlsio.Template(os.path.join(SRC, "template.json"))
    pairs = tpl.day_column_map()
    capacity = len(pairs)
    assert capacity, "khuon khong co cot ngay"

    # 1. Thang cung voi khuon: van dung nhan "Ngay dd-mm" nhu cu
    same = _report.build_rows(Pull(make_days(11, 3)), tpl)
    rows, used = same
    assert used, "thang 11: khong ghi duoc ngay nao"

    # 2. Thang KHAC khuon: phai van ghi duoc du lieu, khong duoc rong
    for month in (10, 12, 1):
        rows, used = _report.build_rows(Pull(make_days(month, 3)), tpl)
        assert rows, "thang %d: khong co dong nao (du lieu bi bo)" % month
        assert len(used) == 3, \
            "thang %d: chi ghi %d/3 ngay" % (month, len(used))
        vals = rows[0]["values"]
        assert any(v for k, v in vals.items() if k not in (tpl.total_pair() or ())), \
            "thang %d: dong du lieu khong co so nao" % month

    # 3. Nhieu ngay hon so cot: ghi toi da capacity, khong tràn cot
    rows, used = _report.build_rows(Pull(make_days(10, capacity + 5)), tpl)
    assert len(used) == capacity, \
        "chi duoc ghi %d ngay (so cot khuon), thuc te %d" % (capacity, len(used))
    allowed = set(tpl.col_xf)               # moi cot khuon co dinh dang
    for col in rows[0]["values"]:
        assert col in allowed, "ghi ra cot ngoai bang: %d" % col

    print("OK: thang khac khuon van xuat duoc du lieu, khong tran cot")


if __name__ == "__main__":
    main()
