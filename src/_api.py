"""Client goi API qlmn.vn (PMS mam non) - da reverse tu session TIENAN."""
import json
import re
import time
import urllib.parse

from _http import Http, HttpError, SessionExpired
from _ocr import read_captcha

BASE = "https://qlmn.vn"
LOGIN_URL = BASE + "/doing/admin/User/login"
CAPTCHA_URL = BASE + "/captcha"
MENU_LIST_URL = BASE + "/doing/dinhduong/menu_report/list"
MARKETBILL_URL = BASE + "/dinhduong/marketbill/marketbill"
CASH_SHOW_URL = BASE + "/doing/dinhduong/cash_book/cash_show"

UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/134.0.0.0 Safari/537.36")


class LoginFailed(Exception):
    pass


class TienAnApi:
    def __init__(self, username, password, timeout=45, log=None):
        self.username = username
        self.password = password
        self.http = Http(UA, timeout=timeout)
        self.csrf = None
        self.log = log or (lambda msg: None)
        self.last_captcha = ""
        self.last_captcha_conf = 0.0
        self.attempts = 0

    # ------------------------------------------------------------ ha tang
    def _headers(self, referer=None, ajax=True):
        h = {"Origin": BASE}
        if referer:
            h["Referer"] = referer
        if ajax:
            h["X-Requested-With"] = "XMLHttpRequest"
        if self.csrf:
            h["X-CSRF-TOKEN"] = self.csrf
        return h

    def _bootstrap_home(self):
        """Mo trang chu de lay cookie + token CSRF."""
        status, _, html = self.http.get(BASE + "/", {"Accept": "text/html,*/*"})
        m = re.search(r'id="csrf-token"\s+content="([^"]+)"', html)
        if m:
            self.csrf = m.group(1)
        return html

    def _refresh_csrf(self):
        try:
            self._bootstrap_home()
        except HttpError:
            pass

    # ------------------------------------------------------------ captcha
    def fetch_captcha(self):
        """Tai anh ma bao ve (bytes PNG)."""
        return self._fetch_binary(CAPTCHA_URL + "?sid=%d" % int(time.time() * 1000))

    def _fetch_binary(self, url):
        import urllib.request
        req = urllib.request.Request(url, headers={
            "User-Agent": UA,
            "Accept": "image/png,image/*,*/*",
            "Accept-Encoding": "identity",
            "Cookie": "; ".join("%s=%s" % (k, v) for k, v in self.http.cookies.items()),
            "Referer": BASE + "/",
        })
        try:
            with urllib.request.urlopen(req, timeout=self.http.timeout) as resp:
                data = resp.read()
        except Exception as exc:
            raise HttpError(0, "Khong tai duoc captcha: %s" % exc, url)
        if not data.startswith(b"\x89PNG"):
            raise HttpError(0, "Du lieu captcha khong hop le", url)
        return data

    # --------------------------------------------------------------- login
    def login(self, max_attempts=25, captcha_solver=None, progress=None):
        """Dang nhap, tu doc captcha va thu lai khi sai.

        captcha_solver: ham(png_bytes) -> (chuoi, do tin cay). Mac dinh dung _ocr.
        progress: ham(attempt, text, confidence, result) de bao tien do UI.
        """
        solver = captcha_solver or read_captcha
        self._bootstrap_home()
        last_error = ""
        for attempt in range(1, max_attempts + 1):
            self.attempts = attempt
            try:
                png = self._fetch_binary(CAPTCHA_URL + "?sid=%d" % int(time.time() * 1000))
                text, conf = solver(png)
                if not text:
                    continue
                self.last_captcha = text
                self.last_captcha_conf = conf
                payload = {
                    "username": self.username,
                    "password": self.password,
                    "captcha": text,
                }
                resp = self.http.post_json(
                    LOGIN_URL, payload,
                    self._headers(BASE + "/"),
                    retries=1,
                )
                result = (resp or {}).get("result")
                if progress:
                    progress(attempt, text, conf, result or "")
                if result == "success":
                    self._refresh_csrf()
                    return True
                err = resp.get("errors") or resp.get("error") or ""
                if isinstance(err, dict):
                    err = "; ".join(
                        "%s" % (" / ".join(v) if isinstance(v, list) else v)
                        for v in err.values()
                    )
                last_error = str(err)
                if "Mã bảo vệ" in last_error or "captcha" in last_error.lower():
                    # sai captcha -> doi cookie/token roi lap lai
                    self._refresh_csrf()
                    continue
                raise LoginFailed(last_error or "Dang nhap that bai")
            except HttpError as exc:
                last_error = str(exc)
                if exc.status in (419, 0):
                    self._refresh_csrf()
                time.sleep(0.8)
                continue
        raise LoginFailed("Khong dang nhap duoc sau %d lan: %s" % (max_attempts, last_error))

    # --------------------------------------------------------------- data
    def _check_session(self, payload):
        if isinstance(payload, dict) and payload.get("result") == "logouted":
            raise SessionExpired("Phien dang nhap da het")
        if isinstance(payload, dict) and payload.get("error") == "Phiên làm việc hết hạn":
            raise SessionExpired("Phien dang nhap da het")
        return payload

    def _call(self, fn, *args, **kwargs):
        """Goi API, tu dang nhap lai khi phien het han (toi da 2 lan)."""
        for attempt in range(3):
            try:
                return fn(*args, **kwargs)
            except SessionExpired:
                if attempt >= 2:
                    raise
                self.relogin()
            except HttpError as exc:
                raise
        raise SessionExpired("Phien dang nhap khong dung duoc")

    def relogin(self):
        self.log("Phien dang nhap het han - dang nhap lai...")
        self.http.cookies.clear()
        self.login(max_attempts=60, progress=self._login_progress)
        self.log("Da dang nhap lai.")

    def _login_progress(self, attempt, text, conf, result):
        if attempt == 1 or attempt % 10 == 0:
            self.log("Dang dang nhap... (lan %d)" % attempt)

    def parse_date(self, text):
        parts = str(text).split("/")
        if len(parts) == 3:
            return "%04d-%02d-%02d" % (int(parts[2]), int(parts[1]), int(parts[0]))
        return str(text)

    def menu_report(self, month):
        payload = self._call(self._menu_report_raw, month)
        return payload

    def _menu_report_raw(self, month):
        payload = self.http.post_json(
            MENU_LIST_URL, {"month": str(month)},
            self._headers(BASE + "/single/dinhduong/menu_report/list/"),
            retries=1,
        )
        return self._check_session(payload)

    def marketbill(self, date_ddmmyyyy, warehouse_ids="2", report_type=1):
        return self._call(self._marketbill_raw, date_ddmmyyyy, warehouse_ids, report_type)

    def _marketbill_raw(self, date_ddmmyyyy, warehouse_ids="2", report_type=1):
        """Lay phieu ke cho cua mot ngay. Tra ve dict data."""
        q = urllib.parse.urlencode({
            "date": date_ddmmyyyy,
            "warehouse_ids": warehouse_ids or "",
            "dataType": "json",
            "type": str(report_type),
            "async": "false",
            "_online": "",
        })
        payload = self.http.get_json(
            MARKETBILL_URL + "?" + q,
            self._headers(BASE + "/single/dinhduong/menu_report/list/"),
        )
        payload = self._check_session(payload)
        if not isinstance(payload, dict) or "data" not in payload:
            raise HttpError(0, "Du lieu phieu ke cho khong hop le", MARKETBILL_URL)
        return payload.get("data") or {}

    def cash_show(self, date_ddmmyyyy):
        return self._call(self._cash_show_raw, date_ddmmyyyy)

    def _cash_show_raw(self, date_ddmmyyyy):
        payload = self.http.post_json(
            CASH_SHOW_URL, {"date": date_ddmmyyyy},
            self._headers(BASE + "/single/dinhduong/menu_report/list/"),
        )
        return self._check_session(payload)

    def warehouses(self, month):
        data = self.menu_report(month).get("data") or {}
        wh = data.get("warehouses") or {}
        if isinstance(wh, dict):
            return [str(k) for k in wh.keys()]
        return [str(x) for x in (wh or [])]

    def workdays(self, month, year=None):
        """Cac ngay co phieu ke cho trong thang - doc tu 'marketbill' cua menu_report."""
        import datetime
        data = self.menu_report(month).get("data") or {}
        cal = data.get("calender") or {}
        year = int(cal.get("year") or year or time.localtime().tm_year)
        days = []
        for key, item in (data.get("marketbill") or {}).items():
            try:
                parts = str(item.get("date") or key).replace("-", "/").split("/")
                if len(parts) == 3:
                    d = datetime.date(int(parts[2]), int(parts[1]), int(parts[0]))
                else:
                    y, m, dd = [int(x) for x in key.split("-")]
                    d = datetime.date(y, m, dd)
            except (TypeError, ValueError):
                continue
            if d.month != month:
                continue
            days.append(d)
        if not days:
            # du phong: doc lich thang
            rows = cal.get("rows") or {}
            for _wk, items in rows.items():
                for item in items:
                    if not item.get("selected"):
                        continue
                    try:
                        d = datetime.date(year, int(item["month"]), int(item["day"]))
                    except (KeyError, ValueError):
                        continue
                    if d.month == month:
                        days.append(d)
        days = sorted(set(days))
        return days
