"""Kiem tra nhanh khong can mang: khuon mau -> sinh file -> doc lai."""
import datetime
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(os.path.dirname(HERE), "src")
sys.path.insert(0, SRC)

import _pull        # noqa: E402
import _report      # noqa: E402
import _xlsio       # noqa: E402


def fake_day(day, foods):
    return {
        "date": datetime.date(2026, 11, day),
        "foods": dict((n, {"qty": q, "price": p, "amount": round(q * p)})
                      for n, q, p in foods),
        "total": 0,
        "services": [],
    }


def main():
    tpl = _xlsio.Template(os.path.join(SRC, "template.json"))
    assert tpl.columns(), "khuon khong co cot so lieu"
    assert tpl.day_column_map(), "khuon khong co cot ngay"

    pull = _pull.PullResult()
    pull.days = [
        fake_day(2, [("Ga", 255, 1400), ("Gạo tám", 21.1, 24000)]),
        fake_day(3, [("Ga", 260, 1400), ("Bắp cải", 12.9, 15000)]),
    ]
    rows, days = _report.build_rows(pull, tpl)
    assert len(days) == 2, days
    assert len(rows) == 3, rows
    ga = [r for r in rows if r["name"] == "Ga"][0]
    qcol, acol = tpl.day_column_map()[0]
    assert ga["values"][qcol] == 255, ga
    assert ga["values"][acol] == 357000, ga
    out = os.path.join(HERE, "ket_qua_thu.xls")
    _xlsio.write_xls(out, tpl, rows)
    assert os.path.getsize(out) > 10000
    os.remove(out)
    print("OK: doc khuon, sinh bao cao tu du lieu gia lap")


if __name__ == "__main__":
    main()
