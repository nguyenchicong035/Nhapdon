import os
from datetime import date
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
from openpyxl import Workbook

from inventory.ledger import Ledger, period, snapshot_rows


def show(app):
    win = tk.Toplevel(app.root)
    win.title("Tồn kho dự tính • toàn bộ kho")
    data = Path(os.environ.get("LOCALAPPDATA") or Path.home()) / "Nhapdon"
    data.mkdir(parents=True, exist_ok=True)
    ledger = Ledger(data / "ton_kho.sqlite3")
    labels = ("Mã hàng","Tên hàng","ĐVT","Tồn đầu kỳ","Đã bán trong kỳ","Tồn cuối kỳ")
    state = {}
    bar = ttk.Frame(win, padding=8)
    bar.pack(fill="x")

    def safe(action):
        def run():
            try:
                action()
            except Exception as exc:
                messagebox.showerror("Tồn kho", str(exc), parent=win)
        return run

    def clear(*_):
        state.clear()
        grid.delete(*grid.get_children())
        note.set("Chọn kỳ rồi nhấn Xem báo cáo.")

    def snapshot():
        path = filedialog.askopenfilename(parent=win,title="Danh mục vật tư có cột Số lượng tồn",filetypes=[("Excel","*.xlsx")])
        if not path:
            return
        stamp = simpledialog.askstring("Thời điểm chốt", "Số tồn trong tệp là số dư CUỐI ngày nào? (DD/MM/YYYY)",parent=win)
        if stamp:
            entries = snapshot_rows(path)
            ledger.snapshot(stamp, entries)
            clear()
            note.set(f"Đã lưu số dư cuối ngày {stamp} cho {len(entries)} mã hàng.")

    def record():
        if app.busy or not app.rows:
            raise ValueError("Cần quét xong các đơn hàng và kiểm tra mã, ngày trước khi ghi.")
        if messagebox.askyesno("Ghi sổ bán","Ghi toàn bộ đơn trong bảng quét, kể cả dòng đang ẩn bởi bộ lọc?",parent=win):
            new, skipped = ledger.record(app.rows.values())
            clear()
            note.set(f"Đã ghi {new} đơn mới. Bỏ qua {skipped} đơn đã ghi giống hệt.")

    def report():
        start,end = period(selected.get(), mode.get())
        stamp,rows = ledger.report(start,end)
        clear()
        state.update(start=start,end=end,stamp=stamp,rows=rows)
        for row in rows:
            grid.insert("", "end", values=row, tags=("negative",) if row[-1]<0 else ())
        note.set(f"Kỳ {start:%d/%m/%Y} – {end:%d/%m/%Y} • Số dư gốc cuối ngày {stamp} • {len(rows)} mã")

    def export():
        if not state:
            raise ValueError("Hãy xem báo cáo trước khi xuất.")
        path = filedialog.asksaveasfilename(parent=win,defaultextension=".xlsx",filetypes=[("Excel","*.xlsx")])
        if path:
            book=Workbook()
            sheet=book.active
            sheet.title="Tồn kho"
            sheet.append([note.get()])
            sheet.append(labels)
            for row in state["rows"]:
                sheet.append(row)
            for col,width in zip("ABCDEF",(30,60,12,17,20,17)):
                sheet.column_dimensions[col].width=width
            book.save(path)

    for title,action in [("Nạp tồn MISA…",snapshot),("Ghi đơn đã duyệt",record)]:
        ttk.Button(bar,text=title,command=safe(action)).pack(side="left")
    mode=tk.StringVar(value="Ngày")
    selected=tk.StringVar(value=date.today().strftime("%d/%m/%Y"))
    ttk.Combobox(bar,textvariable=mode,values=("Ngày","Tuần","Tháng"),state="readonly",width=7).pack(side="left",padx=5)
    ttk.Entry(bar,textvariable=selected,width=12).pack(side="left")
    for title,action in [("Xem báo cáo",report),("Xuất Excel…",export)]:
        ttk.Button(bar,text=title,command=safe(action)).pack(side="left")
    note=tk.StringVar(value="Nạp số dư cuối ngày và ghi các đơn đã duyệt để tính tồn.")
    ttk.Label(win,textvariable=note,padding=8).pack(anchor="w")
    grid=ttk.Treeview(win,columns=labels,show="headings",height=18)
    for title,width in zip(labels,(180,360,60,100,120,100)):
        grid.heading(title,text=title)
        grid.column(title,width=width)
    grid.tag_configure("negative",foreground="#ac2638")
    scroll=ttk.Scrollbar(win,orient="vertical",command=grid.yview)
    grid.configure(yscrollcommand=scroll.set)
    scroll.pack(side="right",fill="y")
    grid.pack(fill="both",expand=True)
    mode.trace_add("write",clear)
    selected.trace_add("write",clear)
