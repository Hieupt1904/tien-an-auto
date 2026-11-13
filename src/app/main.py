"""TienAn Auto - tu dong keo so lieu tien an va lap file quyet toan.

Mot cua so, bon nut. Khong can cai them thu vien nao (chi Python 3 co san).
"""
import datetime
import json
import os
import queue
import subprocess
import sys
import threading
import tkinter as tk
from tkinter import messagebox, ttk

HERE = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.dirname(HERE)
for p in (SRC, HERE):
    if p not in sys.path:
        sys.path.insert(0, p)

import _api        # noqa: E402
import _log        # noqa: E402
import _pull       # noqa: E402
import _report     # noqa: E402
import _session    # noqa: E402
import _xlsio      # noqa: E402

APP_TITLE = "Tiền ăn – tự động cập nhật"
CFG_NAME = "config.json"
TPL_NAME = "template.json"
SESSION_NAME = "session.json"
LOG_DIR = "Nhật ký"
OUT_DIR = "Báo cáo"        # thu muc chua file ket qua (ten khong dau de Windows
                            # va git deu xu ly on dinh)


def app_dir():
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.dirname(HERE))


def load_config(path):
    try:
        with open(path, "r", encoding="utf-8") as fh:
            return json.load(fh)
    except Exception:
        return {}


def save_config(path, cfg):
    try:
        with open(path, "w", encoding="utf-8") as fh:
            json.dump(cfg, fh, ensure_ascii=False, indent=1)
    except Exception:
        pass


def find_template(base):
    for cand in (os.path.join(base, TPL_NAME), os.path.join(SRC, TPL_NAME),
                 os.path.join(getattr(sys, "_MEIPASS", SRC), TPL_NAME)):
        if os.path.exists(cand):
            return cand
    return None


class App:
    def __init__(self, root):
        self.root = root
        self.base = app_dir()
        self.cfg_path = os.path.join(self.base, CFG_NAME)
        self.session_path = os.path.join(self.base, SESSION_NAME)
        self.cfg = load_config(self.cfg_path)
        self.logger = _log.Logger(self.base)
        self.queue = queue.Queue()
        self.worker = None

        root.title(APP_TITLE)
        root.geometry("660x540")
        root.minsize(580, 500)
        root.columnconfigure(0, weight=1)

        style = ttk.Style()
        for theme in ("vista", "winnative", "clam"):
            try:
                style.theme_use(theme)
                break
            except tk.TclError:
                continue
        style.configure("Big.TButton", font=("Segoe UI", 12, "bold"), padding=16)

        head = ttk.Frame(root, padding=(18, 16, 18, 4))
        head.grid(row=0, column=0, sticky="ew")
        self.lbl_user = ttk.Label(head, text="", font=("Segoe UI", 10))
        self.lbl_user.pack(anchor="w")
        ttk.Label(head, text="Bấm một nút để lấy số liệu và lập file quyết toán.",
                  font=("Segoe UI", 9), foreground="#666").pack(anchor="w", pady=(2, 0))

        body = ttk.Frame(root, padding=(18, 8, 18, 8))
        body.grid(row=1, column=0, sticky="nsew")
        root.rowconfigure(1, weight=1)

        self.btn_run = ttk.Button(body, text="▶   CẬP NHẬT DỮ LIỆU", style="Big.TButton",
                                  command=self.on_run)
        self.btn_run.pack(fill="x", ipady=8)

        self.progress = ttk.Progressbar(body, mode="determinate", maximum=100)
        self.progress.pack(fill="x", pady=(14, 6))

        self.box = tk.Frame(body, bg="#eef4ff", bd=0)
        self.box.pack(fill="x")
        self.lbl_status = tk.Label(self.box, text="Sẵn sàng.", font=("Segoe UI", 11),
                                   bg="#eef4ff", fg="#123", anchor="w", justify="left",
                                   wraplength=560, padx=12, pady=10)
        self.lbl_status.pack(fill="x")

        row = ttk.Frame(body)
        row.pack(fill="x", pady=(12, 0))
        self.btn_open = ttk.Button(row, text="📂  Mở thư mục báo cáo", command=self.on_open_dir)
        self.btn_open.pack(side="left", expand=True, fill="x", padx=(0, 6))
        self.btn_cfg = ttk.Button(row, text="⚙  Cài đặt", command=self.on_settings)
        self.btn_cfg.pack(side="left", expand=True, fill="x", padx=(0, 6))
        self.btn_log = ttk.Button(row, text="📜  Nhật ký", command=self.on_log)
        self.btn_log.pack(side="left", expand=True, fill="x")

        foot = ttk.Frame(root, padding=(18, 4, 18, 12))
        foot.grid(row=2, column=0, sticky="ew")
        self.lbl_last = ttk.Label(foot, text="", font=("Segoe UI", 9), foreground="#555")
        self.lbl_last.pack(anchor="w")
        self.txt_log = tk.Text(foot, height=5, font=("Consolas", 9), wrap="word",
                               bg="#fafafa", fg="#333", relief="flat", bd=1, state="disabled")
        self.txt_log.pack(fill="x", pady=(6, 0))

        self.refresh_labels()
        self.root.after(150, self.pump)

    # ------------------------------------------------------------- hien thi
    def ui_log(self, text):
        try:
            self.txt_log.configure(state="normal")
            self.txt_log.insert("end", datetime.datetime.now().strftime("%H:%M ") + text + "\n")
            lines = int(self.txt_log.index("end-1c").split(".")[0])
            if lines > 80:
                self.txt_log.delete("1.0", "%d.0" % (lines - 80))
            self.txt_log.see("end")
            self.txt_log.configure(state="disabled")
        except Exception:
            pass

    def log(self, text):
        self.queue.put(("log", text))

    def status(self, text, percent=None):
        self.queue.put(("status", text, percent))
        try:
            self.lbl_status.config(text=text)
            if percent is not None:
                self.progress.config(value=percent)
        except Exception:
            pass

    def refresh_labels(self):
        user = self.cfg.get("username") or "(chưa nhập)"
        self.lbl_user.config(text="Tài khoản: %s" % user)
        last = self.cfg.get("last_file")
        if last and os.path.exists(last):
            self.lbl_last.config(text="File gần nhất: %s" % last)
        else:
            self.lbl_last.config(text="Chưa có file báo cáo nào.")

    def set_busy(self, busy):
        state = "disabled" if busy else "normal"
        for b in (self.btn_run, self.btn_open, self.btn_cfg, self.btn_log):
            try:
                b.config(state=state)
            except tk.TclError:
                pass

    def pump(self):
        try:
            while True:
                item = self.queue.get_nowait()
                kind = item[0]
                if kind == "log":
                    self.ui_log(item[1])
                    self.logger.also_print = False
                    self.logger(item[1])
                elif kind == "status":
                    self.lbl_status.config(text=item[1])
                    if len(item) > 2 and item[2] is not None:
                        self.progress.config(value=item[2])
                elif kind == "done":
                    self.set_busy(False)
                    self.progress.config(value=100)
                    self.lbl_status.config(text=item[1])
                    self.worker = None
                    self.refresh_labels()
                    messagebox.showinfo(APP_TITLE, item[1])
                elif kind == "error":
                    self.set_busy(False)
                    self.progress.config(value=0)
                    self.lbl_status.config(text=item[1])
                    self.worker = None
                    messagebox.showerror(APP_TITLE, item[1])
        except queue.Empty:
            pass
        self.root.after(150, self.pump)

    # ----------------------------------------------------------------- nut
    def on_run(self):
        if self.worker and self.worker.is_alive():
            return
        if not self.cfg.get("username") or not self.cfg.get("password"):
            messagebox.showwarning(APP_TITLE,
                                   "Lần đầu dùng: bấm Cài đặt để nhập tài khoản qlmn.vn.")
            self.on_settings()
            return
        if not find_template(self.base):
            messagebox.showerror(APP_TITLE, "Thiếu file template.json cạnh chương trình.")
            return
        self.set_busy(True)
        self.progress.config(value=2)
        self.status("Đang đăng nhập...", 3)
        self.worker = threading.Thread(target=self.run_job, daemon=True)
        self.worker.start()

    def open_folder(self, name):
        folder = os.path.join(self.base, name)
        if not os.path.isdir(folder):
            os.makedirs(folder)
        try:
            os.startfile(folder)
        except AttributeError:
            subprocess.Popen(["xdg-open", folder])
        except Exception:
            messagebox.showinfo(APP_TITLE, folder)

    def on_open_dir(self):
        self.open_folder(OUT_DIR)

    def on_log(self):
        self.open_folder(LOG_DIR)

    def on_settings(self):
        SettingsDialog(self)

    # ------------------------------------------------------------ cong viec
    def run_job(self):
        try:
            self.run_steps()
        except Exception as exc:
            self.log("Lỗi: %s" % exc)
            self.queue.put(("error", "Không cập nhật được.\n\n%s" % exc, False))

    def run_steps(self):
        tpl = _xlsio.Template(find_template(self.base))
        api = _api.TienAnApi(self.cfg["username"], self.cfg["password"], log=self.log)
        self.status("Đang đăng nhập...", 5)

        resumed = False
        saved = _session.load(self.session_path, self.cfg["username"])
        if saved and _session.apply(api, saved):
            try:
                api.menu_report(1)
                resumed = True
                self.log("Dùng lại phiên đăng nhập trước đó.")
            except Exception:
                resumed = False
        if not resumed:
            api.login(max_attempts=80, progress=self.on_login)
            self.log("Đăng nhập xong sau %d lần thử mã bảo vệ." % api.attempts)
        _session.save(self.session_path, api, self.cfg["username"])

        today = datetime.date.today()
        month = int(self.cfg.get("month") or today.month)
        year = int(self.cfg.get("year") or today.year)
        self.status("Đang lấy danh sách ngày...", 10)
        days = api.workdays(month, year) or [today]
        self.log("Tháng %02d/%d: %d ngày có phiếu." % (month, year, len(days)))

        pulled = _pull.PullResult()
        for i, day in enumerate(days):
            self.status("Đang lấy dữ liệu ngày %s..." % day.strftime("%d/%m"),
                        10 + int(70 * i / max(1, len(days))))
            text = day.strftime("%d/%m/%Y")
            data = None
            for attempt in range(3):
                try:
                    data = api.marketbill(text, str(self.cfg.get("warehouse") or "2"))
                    break
                except _api.SessionExpired:
                    api.relogin()
                except Exception as exc:
                    self.log("  lỗi ngày %s: %s" % (text, exc))
            if data is None:
                pulled.warnings.append("Không lấy được ngày %s" % text)
                continue
            item = _pull.extract_day(data)
            item["date"] = day
            pulled.days.append(item)

        self.status("Đang lập file quyết toán...", 85)
        rows, used_days = _report.build_rows(pulled, tpl,
                                            int(self.cfg.get("day_from") or 1),
                                            int(self.cfg.get("day_to") or 31))
        if not rows:
            raise RuntimeError("Tháng này chưa có dữ liệu phiếu kê chợ.")

        out_dir = os.path.join(self.base, OUT_DIR)
        os.makedirs(out_dir, exist_ok=True)
        out_path = os.path.join(out_dir, "Quyet toan tien an Thang %02d-%d.xls" % (month, year))
        _xlsio.write_xls(out_path, tpl, rows)

        self.cfg["last_file"] = out_path
        self.cfg["last_run"] = datetime.datetime.now().strftime("%d/%m/%Y %H:%M")
        save_config(self.cfg_path, self.cfg)
        _session.save(self.session_path, api, self.cfg["username"])
        for warn in pulled.warnings:
            self.log("  ! " + warn)
        self.log("Đã ghi %s (%d mặt hàng)." % (out_path, len(rows)))
        msg = ("Đã cập nhật xong %d ngày, %d mặt hàng.\n\nFile: %s"
               % (len(used_days), len(rows), out_path))
        self.queue.put(("done", msg, True))

    def on_login(self, attempt, text, conf, result):
        if attempt == 1:
            self.status("Đang đăng nhập...", 5)
        if attempt % 5 == 0:
            self.log("  thử mã bảo vệ lần %d" % attempt)


class SettingsDialog:
    def __init__(self, app):
        self.app = app
        self.win = tk.Toplevel(app.root)
        self.win.title("Cài đặt")
        self.win.geometry("430x340")
        self.win.transient(app.root)
        self.win.grab_set()
        cfg = app.cfg

        frm = ttk.Frame(self.win, padding=14)
        frm.pack(fill="both", expand=True)
        ttk.Label(frm, text="Tài khoản qlmn.vn", font=("Segoe UI", 10, "bold")).grid(
            row=0, column=0, columnspan=2, sticky="w", pady=(0, 8))

        self.var_user = tk.StringVar(value=cfg.get("username", ""))
        self.var_pass = tk.StringVar(value=cfg.get("password", ""))
        self.var_month = tk.StringVar(value=str(cfg.get("month", "")))
        self.var_year = tk.StringVar(value=str(cfg.get("year", "")))
        self.var_wh = tk.StringVar(value=str(cfg.get("warehouse", "2")))
        self.var_from = tk.StringVar(value=str(cfg.get("day_from", 1)))
        self.var_to = tk.StringVar(value=str(cfg.get("day_to", 31)))

        rows = [("Tên đăng nhập", self.var_user, False),
                ("Mật khẩu", self.var_pass, True),
                ("Tháng", self.var_month, False),
                ("Năm", self.var_year, False),
                ("Kho (2 = kho trưa)", self.var_wh, False),
                ("Lấy từ ngày", self.var_from, False),
                ("Đến ngày", self.var_to, False)]
        for i, (label, var, secret) in enumerate(rows, start=1):
            ttk.Label(frm, text=label).grid(row=i, column=0, sticky="w", pady=4)
            ent = ttk.Entry(frm, textvariable=var, width=26,
                            show="•" if secret else "")
            ent.grid(row=i, column=1, sticky="w", pady=4)

        ttk.Label(frm, text="Tháng/năm để trống = tháng hiện tại.",
                  foreground="#666").grid(row=9, column=0, columnspan=2,
                                          sticky="w", pady=(10, 0))
        btns = ttk.Frame(frm)
        btns.grid(row=10, column=0, columnspan=2, sticky="e", pady=(16, 0))
        ttk.Button(btns, text="Hủy", command=self.win.destroy).pack(side="right", padx=6)
        ttk.Button(btns, text="Lưu", command=self.save).pack(side="right")

    def save(self):
        cfg = self.app.cfg
        cfg["username"] = self.var_user.get().strip()
        cfg["password"] = self.var_pass.get()
        cfg["month"] = self.var_month.get().strip()
        cfg["year"] = self.var_year.get().strip()
        cfg["warehouse"] = self.var_wh.get().strip() or "2"
        cfg["day_from"] = self.var_from.get().strip() or "1"
        cfg["day_to"] = self.var_to.get().strip() or "31"
        save_config(self.app.cfg_path, cfg)
        self.app.refresh_labels()
        self.win.destroy()


def main():
    root = tk.Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
