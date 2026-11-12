"""Khoa hoi quy: du lieu da luu -> file .xls phai khop tung o (khong can mang).

Snapshot nam trong tests/snapshot/ (JSON do tools/save_snapshot.py tao).
"""
import json
import os
import subprocess
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(os.path.dirname(HERE), "src")
sys.path.insert(0, SRC)
sys.path.insert(0, HERE)

import _report      # noqa: E402
import _xlsio       # noqa: E402
import test_coords  # noqa: E402

SNAPSHOT = os.path.join(HERE, "snapshot", "days.json")


def load_days(path):
    import datetime
    data = json.load(open(path, encoding="utf-8"))
    out = []
    for d in data["days"]:
        y, m, dd = [int(x) for x in d["date"][:10].split("-")]
        item = dict(d)
        item["date"] = datetime.date(y, m, dd)
        out.append(item)
    return out


def main():
    if not os.path.exists(SNAPSHOT):
        print("BO QUA: chua co snapshot")
        return
    tpl = _xlsio.Template(os.path.join(SRC, "template.json"))
    pull = _report.__dict__  # khong dung
    class P:
        pass
    p = P()
    p.days = load_days(SNAPSHOT)
    rows, used = _report.build_rows(p, tpl)
    stream = _xlsio.build_stream(tpl, rows, freeze=False)
    got = test_coords.cells_by_row(stream)

    checked = 0
    for i, row in enumerate(rows):
        r = tpl.first_row + i
        for col, want in row["values"].items():
            have = got.get((r, col))
            assert have is not None, "thieu o hang %d cot %d (%s)" % (r, col, row["name"])
            assert abs(float(have) - float(want)) < 1e-6, \
                "hang %d cot %d: %r != %r" % (r, col, have, want)
            checked += 1
    # tong tung ngay trong file phai bang tong API cua ngay do
    # (cot duoc gan theo NHAN NGAY in o dong tieu de, khong theo thu tu)
    pairs = tpl.day_column_map()
    labels = tpl.day_labels()
    by_label = {}
    for label, pair in zip(labels, pairs):
        md = _report.parse_day_label(label)
        if md:
            by_label[md] = pair
    for day in p.days:
        if not day.get("foods"):
            continue
        pair = by_label.get((day["date"].month, day["date"].day))
        if pair is None:
            continue
        acol = pair[1]
        api_sum = sum(v["amount"] for v in day["foods"].values())
        file_sum = sum(got.get((tpl.first_row + i, acol), 0) for i in range(len(rows)))
        assert abs(file_sum - api_sum) < 1, \
            "ngay %s: cot thanh tien %s != API %s" % (day["date"], file_sum, api_sum)
    print("OK: %d o va tong tung ngay khop snapshot API" % checked)


if __name__ == "__main__":
    main()
