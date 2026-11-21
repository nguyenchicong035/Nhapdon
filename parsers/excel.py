from pathlib import Path
from openpyxl import load_workbook
import xlrd
from .common import records


def parse(path):
    result = []
    if Path(path).suffix.lower() == ".xls":
        book = xlrd.open_workbook(path, on_demand=True)
        try:
            for sheet in book.sheets():
                result.extend(records((sheet.row_values(i) for i in range(sheet.nrows)), sheet.name))
        finally:
            book.release_resources()
    else:
        book = load_workbook(path, read_only=True, data_only=True)
        try:
            for sheet in book:
                result.extend(records(sheet.values, sheet.title))
        finally:
            book.close()
    if not result:
        raise ValueError("Không tìm thấy bảng đơn hàng có tiêu đề phù hợp.")
    return result
