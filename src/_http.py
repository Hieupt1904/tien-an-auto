"""HTTP client toi gian cho qlmn.vn - chi dung thu vien chuan cua Python."""
import gzip
import json
import ssl
import time
import zlib
import urllib.error
import urllib.parse
import urllib.request


class SessionExpired(Exception):
    """Phien dang nhap khong con gia tri - can dang nhap lai."""


class HttpError(Exception):
    def __init__(self, status, body, url):
        self.status = status
        self.body = body
        self.url = url
        super().__init__("HTTP %s %s" % (status, url))


class NoRedirect(urllib.request.HTTPRedirectHandler):
    def redirect_request(self, req, fp, code, msg, headers, newurl):
        return None


class Http:
    """Session HTTP giu cookie, ho tro gzip, khong tu dong theo redirect."""

    def __init__(self, user_agent, timeout=45):
        self.ua = user_agent
        self.timeout = timeout
        self.cookies = {}
        # header gui kem ngay tu request dau (giu phien dang nhap tu file)
        self.extra_headers = {}
        self._opener = urllib.request.build_opener(
            urllib.request.HTTPCookieProcessor(_DictCookies(self.cookies)),
            NoRedirect(),
            urllib.request.HTTPSHandler(context=ssl.create_default_context()),
        )

    def _decode(self, raw, headers):
        enc = (headers.get("Content-Encoding") or "").lower()
        if "gzip" in enc:
            try:
                raw = gzip.decompress(raw)
            except Exception:
                pass
        elif "deflate" in enc:
            try:
                raw = zlib.decompress(raw, -zlib.MAX_WBITS)
            except Exception:
                pass
        charset = "utf-8"
        for part in (headers.get("Content-Type") or "").split(";"):
            part = part.strip()
            if part.lower().startswith("charset="):
                charset = part.split("=", 1)[1].strip().strip('"') or "utf-8"
        try:
            return raw.decode(charset, "replace")
        except LookupError:
            return raw.decode("utf-8", "replace")

    def request(self, url, data=None, headers=None, method=None):
        hdr = {
            "User-Agent": self.ua,
            "Accept": "application/json, text/javascript, */*; q=0.01",
            "Accept-Language": "vi-VN,vi;q=0.9,en;q=0.8",
            "Accept-Encoding": "gzip, deflate",
        }
        if data is not None:
            hdr["Content-Type"] = "application/x-www-form-urlencoded; charset=UTF-8"
        if self.extra_headers:
            hdr.update(self.extra_headers)
        if headers:
            hdr.update({k: v for k, v in headers.items() if v is not None})
        body = None
        if data is not None:
            body = urllib.parse.urlencode(data).encode("utf-8") if isinstance(data, dict) else data
        req = urllib.request.Request(url, data=body, headers=hdr, method=method)
        try:
            resp = self._opener.open(req, timeout=self.timeout)
            raw = resp.read()
            status = resp.status
            rheaders = dict(resp.headers.items())
        except urllib.error.HTTPError as exc:
            raw = exc.read() or b""
            status = exc.code
            rheaders = dict((exc.headers or {}).items())
        except urllib.error.URLError as exc:
            raise HttpError(0, "Khong ket noi duoc may chu (%s)" % exc.reason, url)
        except Exception as exc:
            raise HttpError(0, "Loi mang: %s" % exc, url)
        return status, rheaders, self._decode(raw, rheaders)

    def get(self, url, headers=None):
        return self.request(url, None, headers, "GET")

    def post(self, url, data, headers=None):
        return self.request(url, data, headers, "POST")

    def get_json(self, url, headers=None, retries=3):
        return self._json(lambda: self.get(url, headers), retries)

    def post_json(self, url, data, headers=None, retries=3):
        return self._json(lambda: self.post(url, data, headers), retries)

    def _json(self, call, retries):
        last = None
        for attempt in range(retries):
            try:
                status, _, text = call()
                if status in (419, 401, 403):
                    raise SessionExpired("Phien dang nhap da het (HTTP %s)" % status)
                return json.loads(text)
            except (ValueError, HttpError) as exc:
                last = exc
                if attempt + 1 < retries:
                    time.sleep(0.8 * (attempt + 1))
        raise last if last is not None else RuntimeError("Khong lay duoc du lieu JSON")


class _DictCookies:
    """Cau noi de urllib ghi nhan cookie vao dictionnary cua session."""

    def __init__(self, store):
        self.store = store

    def extract_cookies(self, response, request=None):
        for hname, hval in response.headers.items():
            if hname.lower() != "set-cookie":
                continue
            first = hval.split(";")[0]
            if "=" not in first:
                continue
            name, value = first.split("=", 1)
            name, value = name.strip(), value.strip()
            if not name:
                continue
            if value in ("", "deleted"):
                self.store.pop(name, None)
            else:
                self.store[name] = value

    def add_cookie_header(self, request):
        if not self.store:
            return
        request.add_unredirected_header(
            "Cookie", "; ".join("%s=%s" % (k, v) for k, v in self.store.items())
        )
