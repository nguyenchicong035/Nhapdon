from collections import Counter
from copy import copy
from decimal import Decimal
from pathlib import Path
from tempfile import NamedTemporaryFile
import os

from openpyxl import Workbook, load_workbook
from openpyxl.formula.translate import Translator
from openpyxl.utils import get_column_letter

from parsers.common import FIELDS, LABELS, headers, quantity, money


def detect(sheet):
    best = (0, 1, {})
    for row in sheet.iter_rows(max_row=min(sheet.max_row, 120)):
        mapping = headers([c.value for c in row])
        score = len(set(FIELDS) & mapping.keys())
        if score > best[0]:
            best = score, row[0].row, mapping
    return best[1], {k: get_column_letter(i+1) for k, i in best[2].items() if k in FIELDS}


def overflow(rows):
    return [order for order, n in Counter(r.get("order") for r in rows).items() if n > 500]


def write(path, rows, template=None, sheet_name=None, start=2, mapping=None):
    if template and Path(path).resolve() == Path(template).resolve():
        raise ValueError("Hãy lưu kết quả sang tệp khác tệp mẫu.")
    book = load_workbook(template) if template else Workbook()
    try:
        if template:
            sheet = book[sheet_name]
        else:
            sheet = book.active
            sheet.title = "Đơn đặt hàng"
            sheet.append([LABELS[k] for k in FIELDS[:-1]])
            mapping = {k: get_column_letter(i+1) for i, k in enumerate(FIELDS[:-1])}
        if not {"order", "sku", "qty", "warehouse"}.issubset(mapping):
            raise ValueError("Cần ánh xạ ít nhất mã đơn, mã hàng, số lượng và mã kho.")
        if len(set(mapping.values())) != len(mapping):
            raise ValueError("Hai trường đang dùng cùng một cột.")
        if start < 2:
            raise ValueError("Dòng dữ liệu phải nằm sau tiêu đề.")
        # Only extend an empty data area or a single prototype row.
        if template and any(c.value is not None for row in sheet.iter_rows(min_row=start+1) for c in row):
            raise ValueError("Mẫu có dữ liệu bên dưới dòng bắt đầu. Hãy dùng bản sao mẫu trống.")
        width = sheet.max_column
        # Explicit constants in the prototype supply defaults such as customer or currency.
        defaults = {c.column: (c.value, c.data_type) for c in sheet[start]
                    if c.column_letter not in mapping.values() and c.value is not None}
        for offset, r in enumerate(rows):
            idx = start + offset
            for key in ("order", "sku", "warehouse"):
                if not str(r.get(key, "")).strip():
                    raise ValueError(f"Dòng {offset+1} thiếu {LABELS[key]}.")
            qty = quantity(r.get("qty"))
            values = dict(r, qty=qty)
            if r.get("price") not in (None, ""):
                values["price"] = money(r["price"])
                if values.get("amount") in (None, ""):
                    values["amount"] = values["price"] * qty
            if template and offset:
                for col in range(1, width+1):
                    proto, cell = sheet.cell(start, col), sheet.cell(idx, col)
                    cell._style = copy(proto._style)
                    if col in defaults and defaults[col][1] != "f":
                        cell.value, cell.data_type = defaults[col]
                    if proto.data_type == "f":
                        cell.value = Translator(proto.value, origin=proto.coordinate).translate_formula(cell.coordinate)
            for key, col in mapping.items():
                cell = sheet[f"{col}{idx}"]
                value = values.get(key, "")
                if value in ("", None) and cell.data_type == "f":
                    continue
                if key in ("amount", "tax_amount") and value not in ("", None):
                    value = money(value)
                if key == "tax" and "%" in cell.number_format and value not in ("", None):
                    text = str(value).strip()
                    number = Decimal(text.rstrip("%").strip().replace(",", "."))
                    value = number / 100 if text.endswith("%") or number > 1 else number
                cell.value = value
                if isinstance(value, str):
                    cell.data_type = "s"
        with NamedTemporaryFile(dir=Path(path).parent, suffix=".xlsx", delete=False) as temp:
            temporary = Path(temp.name)
        try:
            book.save(temporary)
            os.replace(temporary, path)
        finally:
            temporary.unlink(missing_ok=True)
    finally:
        book.close()
