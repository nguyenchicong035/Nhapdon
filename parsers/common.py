import re
from datetime import date
import unicodedata
from decimal import Decimal, InvalidOperation


def norm(value):
    text = unicodedata.normalize("NFD", str(value or "").casefold().replace("đ", "d"))
    return " ".join(re.findall(r"[a-z0-9]+", "".join(c for c in text if not unicodedata.combining(c))))


ALIASES = {
    "order": ("Mã đơn hàng", "Mã đơn", "Order ID", "Order SN", "Order number", "Số đơn hàng", "Số chứng từ", "Số đơn đặt hàng", "Mã đơn đặt hàng"),
    "name": ("Tên sản phẩm", "Tên hàng", "Tên hàng hóa", "Product Name", "Item Name", "Tên mặt hàng"),
    "variant": ("Phân loại hàng", "Tên phân loại hàng", "Variation Name", "Variation", "SKU Name", "Phân loại", "Variation Information"),
    "sku": ("SKU", "SKU sản phẩm", "Mã SKU", "SKU sản phẩm gốc", "Parent SKU", "Mã hàng"),
    "child": ("SKU phân loại", "SKU phân loại hàng", "Seller SKU", "Mã SKU người bán", "SKU Reference No", "Seller SKU ID"),
    "qty": ("Số lượng", "Quantity", "Qty", "Số lượng sản phẩm"),
    "unit": ("ĐVT", "Đơn vị tính", "Đơn vị"),
    "price": ("Đơn giá", "Giá bán", "SKU Unit Original Price", "Giá gốc"),
    "tax": ("Thuế suất", "Thuế suất GTGT", "% thuế GTGT", "VAT"),
    "image": ("Hình ảnh sản phẩm", "Hình ảnh", "Product Image"),
    "warehouse": ("Mã kho", "Warehouse Code"),
    "amount": ("Thành tiền",),
    "date": ("Ngày đơn hàng", "Ngày đặt hàng", "Ngày chứng từ", "Ngày tạo đơn", "Created Time", "Order Creation Date"),
    "customer": ("Mã khách hàng", "Mã đối tượng"),
    "buyer": ("Tên khách hàng", "Tên người mua", "Buyer Username", "Tên Người nhận", "Recipient"),
    "tax_amount": ("Tiền thuế", "Tiền thuế GTGT"),
    "source": ("Nguồn",),
}
LOOKUP = {norm(v): k for k, vs in ALIASES.items() for v in vs}
FIELDS = ["order", "sku", "name", "variant", "unit", "warehouse", "qty", "price", "amount", "tax", "tax_amount", "date", "customer", "buyer", "source"]
LABELS = dict(zip(FIELDS, ["Mã đơn hàng", "Mã hàng", "Tên hàng", "Phân loại", "ĐVT", "Mã kho", "Số lượng", "Đơn giá", "Thành tiền", "Thuế suất", "Tiền thuế", "Ngày đơn hàng", "Mã khách hàng", "Tên khách hàng", "Nguồn"]))


def headers(row):
    return {LOOKUP[norm(v)]: i for i, v in enumerate(row) if norm(v) in LOOKUP}


def is_header(cols):
    return {"order", "qty"}.issubset(cols) and bool({"name", "child", "sku"} & cols.keys())


def clean(value):
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def money(value):
    if isinstance(value, (int, float, Decimal)):
        return Decimal(str(value))
    text = re.sub(r"(?i)(vnd|vnđ|₫|đ)|\s+", "", str(value))
    if "," in text and "." in text:
        decimal_sep = "," if text.rfind(",") > text.rfind(".") else "."
        text = text.replace("." if decimal_sep == "," else ",", "").replace(decimal_sep, ".")
    elif re.fullmatch(r"-?[1-9]\d{0,2}([.,]\d{3})+", text):
        text = re.sub(r"[.,]", "", text)
    else:
        text = text.replace(",", ".")
    try:
        number = Decimal(text)
        if number.is_finite():
            return number
    except InvalidOperation:
        pass
    raise ValueError(f"Số tiền không hợp lệ: {value!r}")


def quantity(value):
    try:
        number = Decimal(str(value).strip().replace(",", "."))
        if number.is_finite() and number == number.to_integral_value() and number > 0:
            return int(number)
    except (InvalidOperation, ValueError):
        pass
    raise ValueError(f"Số lượng không hợp lệ: {value!r}")


def records(rows, location):
    cols = {}
    order = ""
    for n, values in enumerate(rows, 1):
        found = headers(values)
        if is_header(found):
            cols = found
            order = ""
            continue
        if not cols:
            continue
        item = {k: clean(values[i]) if i < len(values) else "" for k, i in cols.items()}
        for k in ("price", "amount", "tax_amount"):
            if item.get(k):
                item[k] = money(values[cols[k]])
        if "date" in cols and cols["date"] < len(values) and isinstance(values[cols["date"]], date):
            item["date"] = values[cols["date"]]
        if not any(item.get(k) for k in ("name", "sku", "child", "qty")):
            continue
        order = item.get("order") or order
        item["order"] = order
        item["sku"] = item.get("child") or item.get("sku", "")
        try:
            item["qty"] = quantity(item.get("qty"))
        except ValueError as exc:
            raise ValueError(f"{location}, dòng {n}: {exc}") from exc
        yield item
