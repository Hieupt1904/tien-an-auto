"""Diem vao chinh: chay bang python (hoac dong goi .pyz)."""
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.dirname(HERE)
for p in (SRC, HERE):
    if p not in sys.path:
        sys.path.insert(0, p)

import main  # noqa: E402

if __name__ == "__main__":
    main.main()
