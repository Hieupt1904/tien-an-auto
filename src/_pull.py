"""Keo du lieu phieu ke cho theo ngay, chuan hoa ve bang 'ngay x mat hang'."""
import json
import time


class PullResult:
    def __init__(self):
        self.days = []            # list dict: date, groups, foods {name: {qty, amount, price}}, service
        self.warnings = []
        self.names = []           # thu tu mat hang xuat hien
        self.month = None
        self.year = None

    def all_names(self, prefer_sheet_names=None):
        """Thu tu cot: uu tien ten co trong file mau, sau do theo thu tu xuat hien."""
        seen = []
        lookup = {}
        if prefer_sheet_names:
            for n in prefer_sheet_names:
                key = norm(n)
                if key not in lookup:
                    lookup[key] = len(seen)
                    seen.append(n)
        extra = []
        for d in self.days:
            for name in d["foods"]:
                key = norm(name)
                if key not in lookup:
                    lookup[key] = len(seen) + len(extra)
                    extra.append(name)
        return seen + extra


def norm(text):
    """Chuan hoa ten de so khop (bo dau, bo khoang trang thua)."""
    import unicodedata
    s = unicodedata.normalize("NFD", str(text or ""))
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")
    s = s.replace("đ", "d").replace("Đ", "D")
    s = " ".join(s.lower().replace("(", " ").replace(")", " ").split())
    return s


def _clean_cash_book(raw):
    if isinstance(raw, dict):
        return raw
    if not raw or (isinstance(raw, str) and raw.strip() in ("", "{}", "null")):
        return {}
    try:
        return json.loads(raw)
    except ValueError:
        return {}


def extract_day(data):
    """Tu payload marketbill -> thong tin 1 ngay."""
    cb = _clean_cash_book(data.get("cash_book"))
    markets = cb.get("markets") or {}
    foods = {}
    total = 0
    per_group = {}
    for wh_id, wh in markets.items():
        for key, f in (wh.get("foods") or {}).items():
            name = (f.get("name") or "").strip()
            if not name:
                continue
            qty = f.get("thucmuatheodvt")
            if qty in (None, ""):
                qty = 0
            price = f.get("price") or 0
            try:
                qty = float(qty)
            except (TypeError, ValueError):
                qty = 0.0
            try:
                price = float(price)
            except (TypeError, ValueError):
                price = 0.0
            amount = round(qty * price)
            if name in foods and qty == 0:
                continue
            foods[name] = {"qty": qty, "price": price, "amount": amount,
                           "warehouse": wh_id, "measure_id": f.get("measure_id")}
            total += amount
        groups = (wh.get("total") or {}).get("groups") or {}
        for g, v in groups.items():
            per_group[g] = per_group.get(g, 0) + (v or 0)
    services = cb.get("services") or {}
    svc_rows = []
    svc_total = 0
    for key, s in services.items():
        name = (s.get("name") or "").strip()
        qty = s.get("sotre") or 0
        price = s.get("price") or 0
        amount = s.get("thanhtien")
        if amount is None:
            amount = qty * price
        svc_rows.append({"name": name, "qty": qty, "price": price, "amount": amount})
        svc_total += amount
    dachi = (cb.get("dachitrongngay") or {})
    return {
        "foods": foods,
        "total_food": total,
        "total_service": svc_total,
        "total": total + svc_total,
        "total_from_server": dachi.get("total"),
        "groups": data.get("groups") or {},
        "services": svc_rows,
    }


def pull_days(api, month, year, days, warehouses="2", log=None):
    """days: list datetime.date. Tra ve PullResult."""
    log = log or (lambda m: None)
    res = PullResult()
    res.month = month
    res.year = year
    for d in days:
        text = d.strftime("%d/%m/%Y")
        data = None
        last_err = None
        for attempt in range(3):
            try:
                data = api.marketbill(text, warehouses)
                break
            except Exception as exc:  # noqa: BLE001 - bao loi cho nguoi dung
                last_err = exc
                time.sleep(1.0 + attempt)
        if data is None:
            res.warnings.append("Khong lay duoc ngay %s: %s" % (text, last_err))
            log("  ! loi ngay %s" % text)
            continue
        day = extract_day(data)
        day["date"] = d
        if not day["foods"] and not day["services"]:
            res.warnings.append("Ngay %s chua co phieu ke cho" % text)
            log("  - ngay %s chua co du lieu" % text)
        else:
            log("  + ngay %s: %d mat hang, tong %s" % (text, len(day["foods"]), fmt_money(day["total"])))
        res.days.append(day)
    res.days.sort(key=lambda x: x["date"])
    return res


def fmt_money(v):
    try:
        return "{:,.0f}".format(float(v or 0)).replace(",", ".")
    except (TypeError, ValueError):
        return str(v)
