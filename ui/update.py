from pathlib import Path
from queue import Empty, Queue
import sys
from threading import Thread
import tkinter as tk
from tkinter import messagebox, ttk

import updater


def show(app):
    if app.busy:
        messagebox.showinfo("Cập nhật", "Hãy chờ quét xong các tệp rồi kiểm tra cập nhật.")
        return
    if not getattr(sys, "frozen", False) or sys.platform != "win32":
        messagebox.showinfo("Cập nhật", "Tính năng này dùng cho bản Nhapdon.exe trên Windows.")
        return
    target = Path(sys.executable)
    win = tk.Toplevel(app.root)
    win.title("Cập nhật Nhapdon")
    win.transient(app.root)
    win.grab_set()
    status = tk.StringVar(value=f"Đang kiểm tra GitHub… Phiên bản hiện tại: {updater.VERSION}")
    ttk.Label(win, textvariable=status, padding=18).pack()
    button = ttk.Button(win, text="Đóng", command=win.destroy)
    button.pack(pady=(0, 12))
    events = Queue()

    def run(action, callback):
        def worker():
            try:
                events.put((callback, action(), None))
            except Exception as exc:
                events.put((callback, None, exc))
        Thread(target=worker, daemon=True).start()
        poll()

    def poll():
        try:
            callback, result, error = events.get_nowait()
        except Empty:
            win.after(150, poll)
            return
        if error:
            status.set(f"Cập nhật chưa thành công: {error}")
            win.protocol("WM_DELETE_WINDOW", win.destroy)
            button.configure(text="Đóng", command=win.destroy, state="normal")
        else:
            callback(result)

    def ready(path):
        try:
            updater.install_on_exit(path, target)
        except OSError as exc:
            path.unlink(missing_ok=True)
            messagebox.showerror("Cập nhật", str(exc), parent=win)
            win.destroy()
            return
        app.root.destroy()

    def found(release):
        if not release:
            status.set(f"Bạn đang dùng bản mới nhất ({updater.VERSION}).")
            return
        status.set(f"Có bản {release[0]} • {release[2]/1048576:.1f} MB")
        def start():
            if app.rows and not messagebox.askyesno("Khởi động lại",
                    "Các dòng đang xem sẽ đóng khi cập nhật. Bạn đã xuất hoặc ghi các đơn cần giữ và muốn tiếp tục?",
                    parent=win):
                return
            button.configure(state="disabled")
            win.protocol("WM_DELETE_WINDOW", lambda: None)
            status.set("Đang tải và kiểm tra tệp. Nhapdon sẽ khởi động lại khi hoàn tất.")
            run(lambda: updater.download(release, target.parent), ready)
        button.configure(text="Tải và khởi động lại", command=start)
    run(updater.latest, found)
