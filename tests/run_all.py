"""Chay tat ca kiem tra khong can mang."""
import os
import subprocess
import sys

# Windows console mac dinh la cp1252 -> in tieng Viet se UnicodeEncodeError.
for _stream in (sys.stdout, sys.stderr):
    try:
        _stream.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

HERE = os.path.dirname(os.path.abspath(__file__))
tests = sorted(f for f in os.listdir(HERE) if f.startswith("test_") and f.endswith(".py"))
fail = 0
for name in tests:
    print("== " + name)
    env = dict(os.environ, PYTHONIOENCODING="utf-8")
    r = subprocess.run([sys.executable, os.path.join(HERE, name)], env=env)
    if r.returncode != 0:
        fail += 1
print("\n%d/%d bai dat" % (len(tests) - fail, len(tests)))
sys.exit(1 if fail else 0)
