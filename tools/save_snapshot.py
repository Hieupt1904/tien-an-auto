"""Luu so lieu API cua 1 thang de kiem thu lai (khong can mang).

    python tools/save_snapshot.py --month 11 --year 2026
"""
import argparse
import json
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, os.path.join(ROOT, "src"))

import _api       # noqa: E402
import _pull      # noqa: E402
import _session   # noqa: E402


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--month", type=int, required=True)
    ap.add_argument("--year", type=int, required=True)
    ap.add_argument("--warehouse", default="2")
    ap.add_argument("--out", default=os.path.join(ROOT, "tests", "snapshot", "days.json"))
    args = ap.parse_args()

    cfg = {}
    cfg_path = os.path.join(ROOT, "config.json")
    if os.path.exists(cfg_path):
        cfg = json.load(open(cfg_path, encoding="utf-8"))
    api = _api.TienAnApi(cfg.get("username"), cfg.get("password"))
    saved = _session.load(os.path.join(ROOT, "session.json"), api.username)
    if saved:
        _session.apply(api, saved)
    try:
        api.menu_report(args.month)
    except Exception:
        api.login(max_attempts=80)

    out = {"month": args.month, "year": args.year, "days": []}
    for day in api.workdays(args.month, args.year):
        try:
            data = api.marketbill(day.strftime("%d/%m/%Y"), args.warehouse)
        except Exception as exc:
            print("  bo qua %s: %s" % (day, exc))
            continue
        item = _pull.extract_day(data)
        item["date"] = day.isoformat()
        out["days"].append(item)
        print("  %s: %d mat hang, %d" % (day.strftime("%d/%m"), len(item["foods"]),
                                         item["total"]))
    os.makedirs(os.path.dirname(args.out), exist_ok=True)
    json.dump(out, open(args.out, "w", encoding="utf-8"), ensure_ascii=False)
    print("->", args.out)


if __name__ == "__main__":
    main()
