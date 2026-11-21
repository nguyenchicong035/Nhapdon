import json
import os
import shutil
import sys
import queue
import re
from tempfile import NamedTemporaryFile
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, simpledialog, ttk

from openpyxl import load_workbook
from parsers import excel, pdf
from parsers.common import FIELDS, LABELS, norm
from exporters.misa import detect, overflow, write
from rules.engine import apply, accessories, set_prices_to_1000
from ui.inventory import show as show_inventory

def config_directory():
    bundled = Path(__file__).resolve().parents[1] / "config"
    if not getattr(sys, "frozen", False):
        return bundled
    target = Path(os.environ["LOCALAPPDATA"]) / "Nhapdon" / "config"
    target.mkdir(parents=True, exist_ok=True)
    for source in bundled.iterdir():
        if source.is_file() and not (target / source.name).exists():
            shutil.copy2(source, target / source.name)
    return target


CONFIG = config_directory() / "rules_config.json"


class TemplateDialog(simpledialog.Dialog):
    def __init__(self, parent, book):
        self.book = book
        super().__init__(parent, "Ánh xạ mẫu MISA")

    def body(self, frame):
        self.sheet = tk.StringVar(value=self.book.active.title)
        self.start = tk.StringVar()
        ttk.Label(frame, text="Sheet nhập liệu").grid(row=0, column=0)
        box = ttk.Combobox(frame, textvariable=self.sheet, values=self.book.sheetnames, state="readonly")
        box.grid(row=0, column=1)
        box.bind("<<ComboboxSelected>>", self.detect)
        ttk.Label(frame, text="Dòng dữ liệu đầu tiên").grid(row=1, column=0)
        ttk.Entry(frame, textvariable=self.start).grid(row=1, column=1)
        ttk.Label(frame, text="Nhập chữ cái cột Excel (A, B, …), để trống để bỏ qua").grid(row=2, columnspan=2)
        self.fields = {}
        for i, k in enumerate(FIELDS[:-1], 3):
            ttk.Label(frame, text=LABELS[k]).grid(row=i, column=0, sticky="w")
            self.fields[k] = tk.StringVar()
            ttk.Entry(frame, textvariable=self.fields[k]).grid(row=i, column=1)
        self.detect()

    def detect(self, *_):
        row, mapping = detect(self.book[self.sheet.get()])
        self.start.set(row+1)
        for k, var in self.fields.items():
            var.set(mapping.get(k, ""))

    def validate(self):
        mapping = {k: v.get().strip().upper() for k, v in self.fields.items() if v.get().strip()}
        if not self.start.get().isdigit() or any(not re.fullmatch("[A-Z]{1,3}", c) for c in mapping.values()):
            messagebox.showerror("Ánh xạ", "Kiểm tra số dòng và chữ cái cột.", parent=self)
            return False
        self.result = self.sheet.get(), int(self.start.get()), mapping
        return True


class CatalogChoice(simpledialog.Dialog):
    def __init__(self, parent, source, row, catalog):
        self.source, self.row, self.catalog = source, row, catalog
        super().__init__(parent, "Xác nhận mã MISA cho SKU sàn")

    def body(self, frame):
        ttk.Label(frame, text=f"SKU sàn: {self.source}").pack(anchor="w")
        ttk.Label(frame, text=self.row.get("name","") + " · " + self.row.get("variant",""),
                  wraplength=560).pack(anchor="w")
        ttk.Label(frame, text="Chọn mã áp dụng cho mọi dòng cùng SKU và các lần quét sau:").pack(anchor="w")
        self.choices = {f"{k}  |  {v['name']} ({v['unit']})": k for k,v in self.catalog.items()}
        self.box = ttk.Combobox(frame, width=78, values=list(self.choices), state="readonly")
        self.box.pack()
        return self.box

    def validate(self):
        self.result = self.choices.get(self.box.get())
        return self.result is not None


class App:
    def __init__(self, root):
        self.root = root
        root.title("Nhập đơn · Shopee / TikTok → MISA AMIS")
        root.geometry("1220x740")
        self.cfg = json.loads(CONFIG.read_text(encoding="utf-8"))
        self.rows = {}
        self.seen = set()
        self.generated = {}
        self.events = queue.Queue()
        self.busy = False
        bar = ttk.Frame(root, padding=8)
        bar.pack(fill="x")
        self.import_btn = ttk.Button(bar, text="Thêm PDF / Excel", command=self.pick)
        self.import_btn.pack(side="left")
        ttk.Button(bar, text="Tồn kho…", command=lambda: show_inventory(self)).pack(side="left")
        ttk.Label(bar, text="   Kho mặc định ").pack(side="left")
        self.store = ttk.Combobox(bar, values=list(self.cfg["warehouses"]), state="readonly", width=16)
        self.store.current(0)
        self.store.pack(side="left")
        self.auto = tk.BooleanVar(value=True)
        ttk.Checkbutton(bar, text="Nhận diện kho từ tên tệp", variable=self.auto).pack(side="left")
        ttk.Button(bar, text="Xuất theo mẫu MISA…", command=lambda: self.export(True)).pack(side="right")
        ttk.Button(bar, text="Xuất bảng đối soát…", command=lambda: self.export(False)).pack(side="right")
        self.status = ttk.Treeview(root, columns=("file", "status"), show="headings", height=5)
        for col, title in [("file", "Tệp nguồn"), ("status", "Trạng thái")]:
            self.status.heading(col, text=title)
        self.status.pack(fill="x", padx=8)
        actions = ttk.Frame(root, padding=8)
        actions.pack(fill="x")
        ttk.Label(actions, text="Lọc dòng").pack(side="left")
        self.search = tk.StringVar()
        ttk.Entry(actions, textvariable=self.search).pack(side="left", padx=8)
        self.search.trace_add("write", lambda *_: self.refresh())
        self.missing = tk.BooleanVar()
        self.missing_button = ttk.Checkbutton(actions, text="Thiếu mã", variable=self.missing, command=self.refresh)
        self.missing_button.pack(side="left")
        ttk.Button(actions, text="Ghép mã MISA…", command=self.resolve).pack(side="left")
        ttk.Button(actions, text="Đơn giá 1.000 đ", command=self.bulk_price).pack(side="left")
        ttk.Label(actions, text="Nhấp đúp vào ô để sửa").pack(side="left")
        ttk.Button(actions, text="Xóa dòng đã chọn", command=self.remove).pack(side="right")
        frame = ttk.Frame(root)
        frame.pack(fill="both", expand=True, padx=8, pady=6)
        self.grid = ttk.Treeview(frame, columns=FIELDS, show="headings")
        for k in FIELDS:
            self.grid.heading(k, text=LABELS[k])
            self.grid.column(k, width=220 if k in ("name", "source") else 110, stretch=False)
        for orient, side in [("horizontal", "bottom"), ("vertical", "right")]:
            scroll = ttk.Scrollbar(frame, orient=orient, command=self.grid.xview if side=="bottom" else self.grid.yview)
            scroll.pack(side=side, fill="x" if side=="bottom" else "y")
            self.grid.configure(**{("x" if side=="bottom" else "y")+"scrollcommand": scroll.set})
        self.grid.pack(fill="both", expand=True)
        self.grid.bind("<Double-1>", self.edit)

    def pick(self):
        paths = filedialog.askopenfilenames(filetypes=[("Đơn hàng", "*.pdf *.xlsx *.xls")])
        tasks = []
        for path in paths:
            if path in self.seen:
                continue
            self.seen.add(path)
            token = self.status.insert("", "end", values=(Path(path).name, "Chờ xử lý"))
            tasks.append((path, token))
        if not tasks:
            return
        self.busy = True
        self.import_btn.state(["disabled"])
        store, auto = self.store.get(), self.auto.get()

        def worker():
            for path, token in tasks:
                self.events.put((token, path, "Đang đọc", None))
                try:
                    items = pdf.parse(path, self.cfg["ocr_languages"]) if path.lower().endswith(".pdf") else excel.parse(path)
                    rows = apply(items, path, self.cfg, store, auto)
                    self.events.put((token, path, f"Đã đọc {len(rows)} dòng", rows))
                except Exception as exc:
                    self.events.put((token, path, f"Lỗi: {exc}", False))
            self.events.put(None)
        threading.Thread(target=worker, daemon=True).start()
        self.poll()

    def poll(self):
        while not self.events.empty():
            event = self.events.get()
            if event is None:
                self.busy = False
                self.import_btn.state(["!disabled"])
                self.refresh()
                return
            token, path, status, rows = event
            self.status.item(token, values=(Path(path).name, status))
            if rows is False:
                self.seen.discard(path)
            elif rows:
                for i, r in enumerate(rows):
                    self.rows[f"{token}-{i}"] = r
                for adapter in accessories(rows, self.cfg["adapter"]):
                    order = adapter["order"]
                    if order not in self.generated:
                        key = f"extra-{token}-{len(self.generated)}"
                        self.generated[order] = key
                        self.rows[key] = adapter
                    elif self.cfg["adapter"]["mode"] == "per_unit":
                        old = self.rows.get(self.generated[order])
                        if old:
                            old["qty"] = int(old["qty"]) + adapter["qty"]
        self.root.after(60, self.poll)

    def refresh(self):
        self.grid.delete(*self.grid.get_children())
        query = norm(self.search.get())
        missing = sum(not r.get("order") or not r.get("sku") for r in self.rows.values())
        self.missing_button.configure(text=f"Thiếu mã ({missing})")
        for key, r in self.rows.items():
            if self.missing.get() and r.get("order") and r.get("sku"):
                continue
            values = [r.get(k, "") for k in FIELDS]
            if query in norm(" ".join(map(str, values))):
                self.grid.insert("", "end", iid=key, values=values)

    def resolve(self):
        if self.busy:
            return
        selected = self.grid.selection()
        row = self.rows[selected[0]] if selected else None
        if not row:
            row = next((r for r in self.rows.values() if not r.get("sku")), None)
        if row is None:
            messagebox.showinfo("Ghép mã", "Toàn bộ dòng đã có mã hàng.")
            return
        key = row.get("_mapping_key")
        if not key:
            messagebox.showinfo("Ghép mã", "Dòng thiếu SKU nguồn. Nhấp đúp ô Mã hàng để điền riêng cho dòng này.")
            return
        code = CatalogChoice(self.root, key, row, self.cfg.get("catalog", {})).result
        if not code:
            return
        try:
            cfg = json.loads(CONFIG.read_text(encoding="utf-8"))
            cfg["sku_to_misa"][key] = code
            cfg["review_skus"] = [s for s in cfg.get("review_skus", []) if s != key]
            with NamedTemporaryFile(mode="w", encoding="utf-8", dir=CONFIG.parent, delete=False) as tmp:
                json.dump(cfg, tmp, ensure_ascii=False, indent=2)
            os.replace(tmp.name, CONFIG)
        except (OSError, ValueError) as exc:
            messagebox.showerror("Không lưu được ánh xạ", str(exc))
            return
        self.cfg = cfg
        for item in self.rows.values():
            if item.get("_mapping_key") == key:
                item.update(sku=code, **cfg["catalog"][code])
        self.refresh()

    def edit(self, event):
        key = self.grid.identify_row(event.y)
        col = self.grid.identify_column(event.x)
        if not key or not col:
            return
        field = FIELDS[int(col[1:])-1]
        value = simpledialog.askstring("Sửa dòng", LABELS[field], initialvalue=self.rows[key].get(field, ""))
        if value is not None:
            self.rows[key][field] = value
            if field in ("qty", "price"):
                self.rows[key].pop("amount", None)
            self.refresh()

    def bulk_price(self):
        if self.busy or not self.rows:
            messagebox.showinfo("Đơn giá", "Hãy quét xong các tệp trước khi đặt đơn giá.")
            return
        if not messagebox.askyesno("Đặt đơn giá", f"Đặt đơn giá 1.000 đồng cho toàn bộ {len(self.rows)} dòng, kể cả dòng đang ẩn, và tính lại thành tiền?"):
            return
        try:
            set_prices_to_1000(self.rows.values())
        except ValueError as exc:
            messagebox.showerror("Kiểm tra số lượng", str(exc))
            return
        self.refresh()

    def remove(self):
        keys = self.grid.selection()
        if keys and messagebox.askyesno("Xóa dòng", f"Xóa {len(keys)} dòng khỏi bảng xuất?"):
            for key in keys:
                del self.rows[key]
            self.refresh()

    def export(self, use_template):
        if self.busy or not self.rows:
            messagebox.showinfo("Xuất Excel", "Cần có dữ liệu và hoàn tất quét trước khi xuất.")
            return
        rows = sorted(self.rows.values(), key=lambda r: str(r.get("order", "")))
        if use_template and any(not r.get("order") or not r.get("sku") for r in rows):
            self.missing.set(True)
            self.search.set("")
            self.refresh()
            messagebox.showinfo("Cần đối chiếu", "Đã lọc các dòng thiếu mã. Dùng Ghép mã MISA cho SKU sàn hoặc nhấp đúp để điền mã đơn từ phiếu gốc. Xuất bảng đối soát sẽ lấy toàn bộ dòng.")
            return
        large = overflow(rows)
        if large and not messagebox.askyesno("Vượt 500 dòng", f"Chứng từ vượt 500 dòng: {', '.join(large)}. Vẫn xuất?"):
            return
        options = {}
        try:
            if use_template:
                path = filedialog.askopenfilename(title="Chọn mẫu nhập đơn đặt hàng từ MISA", initialdir=str(CONFIG.parent), initialfile="Mau_MISA_don_dat_hang.xlsx", filetypes=[("Excel", "*.xlsx")])
                if not path:
                    return
                book = load_workbook(path)
                try:
                    result = TemplateDialog(self.root, book).result
                finally:
                    book.close()
                if not result:
                    return
                sheet, start, mapping = result
                options = dict(template=path, sheet_name=sheet, start=start, mapping=mapping)
            target = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")])
            if target:
                if str(Path(target).resolve()) in {str(Path(p).resolve()) for p in self.seen}:
                    raise ValueError("Hãy chọn tên khác tệp nguồn.")
                write(target, rows, **options)
                messagebox.showinfo("Hoàn tất", f"Đã xuất {len(rows)} dòng.")
        except Exception as exc:
            messagebox.showerror("Chưa xuất được", str(exc))


def main():
    root = tk.Tk()
    App(root)
    if "--self-test" in sys.argv:
        import pymupdf
        import pypdfium2
        book = load_workbook(CONFIG.parent / "Mau_MISA_don_dat_hang.xlsx", read_only=True)
        assert "Don dat hang" in book.sheetnames
        book.close()
        root.update_idletasks()
        root.destroy()
        return
    root.mainloop()
