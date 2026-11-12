"""Ghi nhat ky chung cho ung dung."""
import datetime
import os
import time

LOG_DIR = "Nhật ký"


class Logger:
    def __init__(self, base_dir):
        self.base = base_dir
        self.also_print = False

    def _path(self):
        folder = os.path.join(self.base, LOG_DIR)
        if not os.path.isdir(folder):
            os.makedirs(folder)
        return os.path.join(folder, datetime.date.today().strftime("%Y-%m") + ".log")

    def __call__(self, text):
        line = "[%s] %s" % (time.strftime("%d/%m %H:%M:%S"), text)
        if self.also_print:
            print(line)
        try:
            with open(self._path(), "a", encoding="utf-8") as fh:
                fh.write(line + "\n")
        except Exception:
            pass
