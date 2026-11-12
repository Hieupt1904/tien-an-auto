"""Ghep du lieu phieu ke cho theo ngay thanh bang quyet toan (theo khuon file mau).

Nguyen tac:
  * Moi cap cot trong khuon la mot NGAY, gan theo NHAN NGAY in tren dong tieu de
    (Ngày 02-11, Ngày 03-11, ...). Ngay khong co phieu thi de trong cot cua no.
  * Ten mat hang so khop CHINH XAC (chi bo dau tieng Viet). Hai ten khac nhau
    tren he thong la hai mat hang khac nhau.
"""
import datetime
import re

_SPECIAL = {"\u0111": "d", "\u0110": "D"}


def fold(text):
    """Bo dau tieng Viet."""
    import unicodedata
    s = unicodedata.normalize("NFD", str(text or ""))
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")
    for a, b in _SPECIAL.items():
        s = s.replace(a, b)
    return s


def key(text):
    """Khoa so khop chinh xac: giu khoang trang, dau ngoac; chi bo dau tieng Viet."""
    return " ".join(fold(text).split()).lower()


def parse_day_label(label, default_month=None):
    """'Ngày 02-11' -> ngay 2 thang 11 (khong co nam)."""
    m = re.search(r"(\d{1,2})\s*[-/]\s*(\d{1,2})", str(label or ""))
    if not m:
        return None
    return (int(m.group(2)), int(m.group(1)))       # (thang, ngay)


def build_rows(pull, template, day_from=1, day_to=31):
    """Tra ve (rows, used_days) cho _xlsio.write_xls.

    rows: [{'name': ten, 'values': {cot: gia tri}}]
    used_days: cac ngay da ghi vao file (de bao cao lai).
    """
    # ngay -> du lieu
    by_day = {}
    for day in pull.days:
        d = day["date"]
        if not day.get("foods"):
            continue
        if not (day_from <= d.day <= day_to):
            continue
        by_day[d] = day

    # cot ngay trong khuon, gan theo nhan "Ngày dd-mm"
    slots = []
    for i, (label, pair) in enumerate(zip(template.day_labels(), template.day_column_map())):
        md = parse_day_label(label)
        if md is None:
            continue
        slots.append((md[0], md[1], pair))      # (thang, ngay, (qcol, acol))

    # gom mat hang
    table = {}
    order = []
    for d, day in sorted(by_day.items()):
        for name, info in day["foods"].items():
            k = key(name)
            if k not in table:
                table[k] = {"name": name, "days": {}}
                order.append(k)
            prev = table[k]["days"].get(d)
            if prev:
                prev["qty"] += info["qty"]
                prev["amount"] += info["amount"]
            else:
                table[k]["days"][d] = dict(info)

    order.sort(key=lambda k: (-len(table[k]["days"]),
                              -sum(v["amount"] for v in table[k]["days"].values()),
                              table[k]["name"].lower()))

    total_pair = template.total_pair()
    used = set()
    rows = []
    for k in order:
        item = table[k]
        values = {}
        qty_total = 0.0
        amount_total = 0.0
        for (mm, dd, (qcol, acol)) in slots:
            day = None
            for d in item["days"]:
                if d.day == dd and (d.month == mm or mm is None):
                    day = d
                    break
            if day is None:
                continue
            info = item["days"][day]
            values[qcol] = round(info["qty"], 6)
            values[acol] = round(info["amount"])
            used.add(day)
            qty_total += info["qty"]
            amount_total += info["amount"]
        if total_pair:
            values[total_pair[0]] = round(qty_total, 6)
            values[total_pair[1]] = round(amount_total)
        rows.append({"name": item["name"], "values": values})
    return rows, sorted(used)


def title(pull, days):
    if days:
        return "THANH TOÁN TIỀN ĂN THÁNG %d" % days[0].month
    return "THANH TOÁN TIỀN ĂN"
