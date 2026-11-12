"""Bo nho phien dang nhap: luu cookie + token de lan sau khoi phai dang nhap lai.

File session.json nam trong thu muc cua app (0700 tren Linux / an tren Windows
khong can). Cookie qlmn.vn song khoang 2 gio, du cho vai lan bam nut lien tiep.
"""
import json
import os
import time


def save(path, api, user):
    if not api or not api.http.cookies:
        return
    data = {
        "user": user,
        "cookies": dict(api.http.cookies),
        "csrf": api.csrf or "",
        "agent": api.http.ua,
        "saved_at": int(time.time()),
    }
    try:
        tmp = path + ".tmp"
        with open(tmp, "w", encoding="utf-8") as fh:
            json.dump(data, fh, ensure_ascii=False)
        os.replace(tmp, path)
        try:
            os.chmod(path, 0o600)
        except Exception:
            pass
    except Exception:
        pass


def load(path, user):
    if not os.path.exists(path):
        return None
    try:
        with open(path, "r", encoding="utf-8") as fh:
            data = json.load(fh)
    except Exception:
        return None
    if data.get("user") != user:
        return None
    # cookie chi song ~2 gio; qua 90 phut thi bo
    if time.time() - float(data.get("saved_at") or 0) > 90 * 60:
        return None
    return data


def apply(api, data):
    if not data:
        return False
    api.http.cookies.clear()
    api.http.cookies.update(data.get("cookies") or {})
    api.csrf = data.get("csrf") or None
    agent = data.get("agent")
    if agent:
        api.http.ua = agent
    return True
