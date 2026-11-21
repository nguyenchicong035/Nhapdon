"""Local stock snapshots and idempotent recording of reviewed orders."""
import calendar
import sqlite3
from collections import Counter, defaultdict
from contextlib import closing
from datetime import date, datetime, timedelta
from decimal import Decimal
from openpyxl import load_workbook
from parsers.common import clean, norm, quantity


def day(value):
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d %H:%M:%S"):
        try:
            return datetime.strptime(str(value).strip(), fmt).date()
        except ValueError:
            continue
    raise ValueError(f"Ngày không hợp lệ: {value}. Dùng DD/MM/YYYY.")


def period(value, mode):
    d = day(value)
    if mode == "Tuần":
        start = d - timedelta(days=d.weekday())
        return start, start + timedelta(days=6)
    if mode == "Tháng":
        return d.replace(day=1), d.replace(day=calendar.monthrange(d.year, d.month)[1])
    return d, d


def snapshot_rows(path):
    book = load_workbook(path, read_only=True, data_only=True)
    try:
        for sheet in book:
            columns = None
            result = {}
            for row in sheet.values:
                if columns is None:
                    h = {norm(v): i for i, v in enumerate(row)}
                    if {"ma", "ten", "don vi tinh chinh", "so luong ton"} <= h.keys():
                        columns = h
                    continue
                sku = clean(row[columns["ma"]])
                if not sku:
                    continue
                if sku in result:
                    raise ValueError(f"Mã trùng trong danh mục: {sku}")
                raw = row[columns["so luong ton"]]
                try:
                    balance = Decimal(str(raw))
                    if not balance.is_finite():
                        raise ValueError()
                except Exception as exc:
                    raise ValueError(f"Số tồn của {sku} không hợp lệ: {raw}") from exc
                result[sku] = (clean(row[columns["ten"]]), clean(row[columns["don vi tinh chinh"]]), str(balance))
            if result:
                return result
        raise ValueError("Không tìm thấy bảng Mã, Tên, Đơn vị tính chính, Số lượng tồn.")
    finally:
        book.close()


class Ledger:
    def __init__(self, path):
        self.path = path
        with closing(sqlite3.connect(path)) as db:
            db.executescript("""
                CREATE TABLE IF NOT EXISTS stock (
                    stamp TEXT, sku TEXT, name TEXT, unit TEXT, balance TEXT,
                    PRIMARY KEY(stamp, sku));
                CREATE TABLE IF NOT EXISTS sales (
                    doc TEXT, sku TEXT, stamp TEXT, qty INTEGER,
                    PRIMARY KEY(doc, sku));
                CREATE INDEX IF NOT EXISTS sales_date ON sales(stamp);
            """)

    def snapshot(self, stamp, entries):
        stamp = day(stamp).isoformat()
        with closing(sqlite3.connect(self.path)) as db, db:
            db.execute("BEGIN IMMEDIATE")
            old = {sku:(name,unit,bal) for sku,name,unit,bal in db.execute(
                "SELECT sku,name,unit,balance FROM stock WHERE stamp=?", (stamp,))}
            if old and old != entries:
                raise ValueError("Đã có bộ số dư khác cùng ngày chốt. Kiểm tra lại ngày của danh mục.")
            if not old:
                db.executemany("INSERT INTO stock VALUES (?,?,?,?,?)",
                               [(stamp,sku,*detail) for sku,detail in entries.items()])

    def record(self, rows):
        grouped = defaultdict(Counter)
        dates = {}
        with closing(sqlite3.connect(self.path)) as db, db:
            db.execute("BEGIN IMMEDIATE")
            known = {r[0] for r in db.execute("SELECT DISTINCT sku FROM stock")}
            if not known:
                raise ValueError("Hãy nạp danh mục tồn đầu tiên.")
            for row in rows:
                doc, sku = clean(row.get("order")), clean(row.get("sku"))
                if not doc or sku not in known:
                    raise ValueError(f"Đơn {doc}: cần mã hàng đã đối chiếu danh mục ({sku or 'đang trống'}).")
                stamp = day(row.get("date")).isoformat()
                if dates.setdefault(doc, stamp) != stamp:
                    raise ValueError(f"Đơn {doc} có nhiều ngày.")
                grouped[doc][sku] += quantity(row["qty"])
            added = 0
            for doc, counts in grouped.items():
                old = list(db.execute("SELECT sku,stamp,qty FROM sales WHERE doc=?", (doc,)))
                if old:
                    if {s:(d,q) for s,d,q in old} != {s:(dates[doc],q) for s,q in counts.items()}:
                        raise ValueError(f"Đơn {doc} đã ghi với dữ liệu khác. Lô này chưa được ghi.")
                    continue
                db.executemany("INSERT INTO sales VALUES (?,?,?,?)",
                               [(doc,s,dates[doc],q) for s,q in counts.items()])
                added += 1
            return added, len(grouped)-added

    def report(self, start, end):
        with closing(sqlite3.connect(self.path)) as db:
            stamp = db.execute("SELECT MAX(stamp) FROM stock WHERE stamp < ?", (start.isoformat(),)).fetchone()[0]
            if not stamp:
                raise ValueError("Cần số dư cuối ngày trước ngày bắt đầu kỳ báo cáo.")
            sold = {s:(before,during) for s,before,during in db.execute("""
                SELECT sku, SUM(CASE WHEN stamp < ? THEN qty ELSE 0 END),
                       SUM(CASE WHEN stamp >= ? THEN qty ELSE 0 END)
                FROM sales WHERE stamp > ? AND stamp <= ? GROUP BY sku
            """, (start.isoformat(),start.isoformat(),stamp,end.isoformat()))}
            rows = []
            for sku,name,unit,balance in db.execute("SELECT sku,name,unit,balance FROM stock WHERE stamp=? ORDER BY sku",(stamp,)):
                before,during = sold.pop(sku,(0,0))
                opening = Decimal(balance)-before
                rows.append((sku,name,unit,opening,during,opening-during))
            if sold:
                raise ValueError("Các mã đã bán thiếu số dư trong danh mục đã chọn: "+", ".join(sold))
            return stamp,rows
