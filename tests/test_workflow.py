import json
from copy import deepcopy
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest

from openpyxl import Workbook, load_workbook
from parsers import excel, pdf
from exporters.misa import write, overflow, detect
from rules.engine import apply, accessories

CFG = json.loads((Path(__file__).resolve().parents[1]/"config/rules_config.json").read_text(encoding="utf8"))


class Workflow(unittest.TestCase):
    def test_excel_to_template(self):
        with TemporaryDirectory() as folder:
            source = Path(folder)/"Shopee HN.xlsx"
            book = Workbook()
            sheet = book.active
            sheet.append(["Shopee"])
            sheet.append(["Order ID", "Product Name", "Seller SKU", "Variation", "Quantity", "Giá bán"])
            sheet.append(["00123", "Máy đo huyết áp", "B2", "Combo 02", 4, 200])
            sheet.append([None, "Máy đo huyết áp", "M", "", 1, 150])
            book.save(source)
            cfg = deepcopy(CFG)
            cfg.update(combo_sku_mode="map", combo_to_unit={"B2":"B"}, sku_to_misa={"B":"MISA01"})
            rows = apply(excel.parse(source), source, cfg, "Hồ Chí Minh", True)
            extra = list(accessories(rows, cfg["adapter"]))
            self.assertEqual((len(extra), extra[0]["qty"]), (1, 1))
            self.assertEqual(list(accessories(rows, dict(cfg["adapter"], mode="per_unit")))[0]["qty"], 9)
            template = Path(folder)/"template.xlsx"
            book = Workbook()
            sheet = book.active
            sheet.title = "Chi tiết"
            sheet.append(["Mẫu kiểm thử"])
            sheet.append(["Mã đơn hàng", "Mã kho", "Mã hàng", "Số lượng", "Đơn giá", "Thành tiền", "Khách hàng"])
            sheet["G3"] = "KH01"
            book.create_sheet("Danh mục")
            book.save(template)
            header, mapping = detect(sheet)
            out = Path(folder)/"out.xlsx"
            write(out, rows+extra, template, sheet.title, header+1, mapping)
            result = load_workbook(out)
            self.assertEqual(result.sheetnames, ["Chi tiết", "Danh mục"])
            self.assertEqual(list(result.active.values)[2], ("00123", "HN", "MISA01", 8, 100, 800, "KH01"))
            self.assertEqual(result.active["C5"].value, "Adapto")
            self.assertEqual(result.active["G5"].value, "KH01")
            result.close()
            self.assertFalse(overflow(rows[:1]*500))
            self.assertEqual(overflow(rows[:1]*501), ["00123"])

    def test_positioned(self):
        words = [dict(x0=x, top=y, text=t) for x,y,t in [
            (10,10,"SKU"),(80,10,"Product"),(115,10,"Name"),(240,10,"Qty"),(300,10,"Order"),(335,10,"ID"),
            (10,30,"P01"),(80,30,"Product"),(240,30,"3"),(300,30,"0101"),(80,44,"wrapped")]]
        rows, layout = pdf.positioned(words, [], 1)
        self.assertEqual((rows[0]["name"], rows[0]["qty"], rows[0]["order"]), ("Product wrapped", 3, "0101"))
