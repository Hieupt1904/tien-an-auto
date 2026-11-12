"""Chay tat ca kiem tra khong can mang."""
import os
import subprocess
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
tests = sorted(f for f in os.listdir(HERE) if f.startswith("test_") and f.endswith(".py"))
fail = 0
for name in tests:
    print("== " + name)
    r = subprocess.run([sys.executable, os.path.join(HERE, name)])
    if r.returncode != 0:
        fail += 1
print("\n%d/%d bai dat" % (len(tests) - fail, len(tests)))
sys.exit(1 if fail else 0)
