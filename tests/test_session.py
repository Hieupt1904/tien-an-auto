import os
import sys
import tempfile
sys.path.insert(0, os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "src"))
import _api
import _session

api = _api.TienAnApi("u", "p")
api.http.cookies.update({"laravel_session": "abc", "pms_route": "x"})
api.csrf = "tok"
path = os.path.join(tempfile.mkdtemp(), "session_test.json")
_session.save(path, api, "u")
data = _session.load(path, "u")
assert data, "khong doc duoc session"
api2 = _api.TienAnApi("u", "p")
assert _session.apply(api2, data)
assert api2.http.cookies.get("laravel_session") == "abc"
assert api2.csrf == "tok"
assert _session.load(path, "khac") is None
os.remove(path)
print("OK: luu/doc phien dang nhap")
