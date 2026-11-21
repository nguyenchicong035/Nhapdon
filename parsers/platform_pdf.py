"""Order-level layout readers for Shopee order slips and TikTok picking sheets."""
from datetime import datetime
import re

import pymupdf

from .common import money, quantity


def cells(words, starts):
    values = [[] for _ in starts]
    for x, y, _, _, text, *_ in words:
        index = next((i for i in reversed(range(len(starts))) if x >= starts[i]-1.5), None)
        if index is not None:
            values[index].append(text)
    return [" ".join(parts) for parts in values]


def picking_catalog(page):
    words = page.get_text("words")
    footer = next((w[1] for w in words if w[4]=="Trang:"),page.rect.height)
    words = [w for w in words if w[1]<footer-2]
    header = next(w for w in words if w[4]=="#")
    line = sorted((w for w in words if abs(w[1]-header[1])<2), key=lambda w:w[0])
    starts = [w[0] for w in line if w[4] in ("#", "SKU", "Hình", "Tên", "Phân", "Số", "Order")]
    if len(starts)!=8:
        raise ValueError("Chưa đọc được bảng SKU Shopee.")
    anchors = [w[1] for w in words if abs(w[0]-starts[6])<2 and w[1]>header[1] and w[4].isdigit()]
    for top,end in zip(anchors,anchors[1:]+[page.rect.height]):
        _,parent,_,_,sku,_,qty,_ = cells([w for w in words if top-1<=w[1]<end-1],starts)
        yield parent.replace(" ",""), sku.replace(" ",""), quantity(qty)


def shopee(page, text, catalog):
    order = re.search(r"\b([0-9]{6}[A-Z0-9]{8})\s+package\s+\d+", text)
    words = page.get_text("words")
    header = next(w for w in words if w[4] == "#" and w[1] > next(v[1] for v in words if v[4]=="OrderSN:"))
    line = [w for w in words if abs(w[1]-header[1]) < 2]
    starts = [w[0] for w in line if w[4] in ("#", "SKU", "Tên", "Phân", "SL", "Đơn", "Thành")]
    if len(starts) != 8 or not order:
        raise ValueError("Bố cục chi tiết Shopee chưa nhận diện được.")
    starts.sort()
    anchors = [w[1] for w in words if abs(w[0]-starts[0]) < 2 and w[1]>header[1] and w[4].isdigit()]
    if not anchors:
        raise ValueError("Thiếu dòng chi tiết Shopee.")
    dates = set(re.findall(r"\b\d{2}-\d{2}-20\d{2}\b",text))
    date = next(iter(dates)).replace("-","/") if len(dates)==1 else ""
    rows = []
    for top, end in zip(anchors, anchors[1:]+[page.rect.height]):
        _, parent, name, sku, variant, qty, price, total = cells(
            [w for w in words if top-1 <= w[1] < end-1], starts)
        parent, sku = parent.replace(" ",""), sku.replace(" ","")
        if not sku or parent.replace("_","")==sku.replace("_","") and "_" not in sku:
            sku = parent
        sku = catalog.get(parent) or sku
        q, p, amount = quantity(qty), money(price), money(total)
        if p*q != amount:
            raise ValueError("Thành tiền Shopee chưa khớp với đơn giá và số lượng.")
        rows.append(dict(order=order[1],name=name,sku=sku,variant=variant,qty=q,price=p,amount=amount,date=date))
    return rows


def tiktok(page, picking):
    words = page.get_text("words")
    qty_header = next(w for w in words if w[4]=="Qty")
    line = [w for w in words if abs(w[1]-qty_header[1])<2]
    starts = [next(w[0] for w in line if w[4]==s) for s in ("Product","SKU","Seller","Qty")]
    if picking:
        starts.append(next(w[0] for w in line if w[4]=="Order"))
    footer = next((w[1] for w in words if w[4]=="Total:"),page.rect.height)
    body = [w for w in words if qty_header[1]+5<w[1]<footer-2]
    anchors = [w[1] for w in body if starts[3]-2<=w[0]< (starts[4]-2 if picking else page.rect.width) and w[4].isdigit()]
    rows = []
    for top,end in zip(anchors,anchors[1:]+[footer]):
        parts = cells([w for w in body if top-2<=w[1]<end-2],starts)
        qty = quantity(parts[3])
        ids = re.findall(r"\b\d{18}\b",parts[4] if picking else page.get_text().split("Order ID:")[-1])
        if not ids or len(ids)>1 and len(ids)!=qty:
            raise ValueError("Không thể phân bổ số lượng cho các mã đơn TikTok.")
        for order in ids:
            rows.append(dict(order=order,name=parts[0],variant=parts[1],sku=parts[2].replace(" ",""),
                             qty=qty if len(ids)==1 else 1))
    total = re.search(r"(?:Qty Total:|Item quantity:)\s*(\d+)",page.get_text())
    if not rows or not total or sum(r["qty"] for r in rows)!=int(total[1]):
        raise ValueError("Tổng số lượng TikTok chưa khớp.")
    return rows


def read(path):
    with pymupdf.open(path) as doc:
        first = doc[0].get_text()
        if "Phiếu xuất hàng" in first and len(doc)>1 and "OrderSN:" in doc[-1].get_text():
            rows = []
            catalog = {}
            expected = 0
            for p in doc:
                text = p.get_text()
                if "Order SN" in text and "OrderSN:" not in text:
                    for parent, sku, count in picking_catalog(p):
                        if parent in catalog and catalog[parent]!=sku:
                            raise ValueError(f"SKU gốc có nhiều mã phân loại: {parent}")
                        catalog[parent] = sku
                        expected += count
                elif "OrderSN:" in text:
                    rows.extend(shopee(p,text,catalog))
                elif "Order SN" not in text:
                    raise ValueError(f"Trang {p.number+1}: chưa nhận diện được loại phiếu.")
            if expected != sum(r["qty"] for r in rows):
                raise ValueError("Tổng chi tiết và phiếu xuất hàng Shopee chưa khớp.")
            return rows
        if first.startswith("Picking List"):
            return [r for p in doc for r in tiktok(p,True)]
        if "Thời gian đặt hàng:" in first and "Order ID:" in first:
            dates,rows = {},[]
            for p in doc:
                text = p.get_text()
                if "Product Name" in text:
                    rows.extend(tiktok(p,False))
                else:
                    order = re.search(r"Order ID:\s*(\d+)",text)
                    date = re.search(r"(20\d\d-\d\d-\d\d) \d\d:\d\d\s+Thời gian đặt hàng:",text)
                    if not order or not date:
                        raise ValueError(f"Trang {p.number+1}: thiếu mã đơn hoặc ngày đặt.")
                    dates[order[1]]=datetime.strptime(date[1],"%Y-%m-%d").strftime("%d/%m/%Y")
            if set(dates)!={r["order"] for r in rows}:
                raise ValueError("Mã đơn trên nhãn và phiếu chi tiết chưa khớp.")
            for r in rows:
                r["date"]=dates[r["order"]]
            return rows
