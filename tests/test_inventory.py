from pathlib import Path
from tempfile import TemporaryDirectory
import unittest

from inventory.ledger import Ledger, period


class Stock(unittest.TestCase):
    def test_periods_and_duplicate_transactions(self):
        with TemporaryDirectory() as temp:
            path = Path(temp)/"stock.db"
            ledger = Ledger(path)
            ledger.snapshot("31/10/2026",{"SKU":("Gel","Gói","20")})
            def sale(doc, date, qty):
                return dict(order=doc,date=date,sku="SKU",qty=qty)
            batch = [sale("X","31/10/2026",9),sale("Y","01/11/2026",3),
                     sale("Z","21/11/2026",2),sale("Z","21/11/2026",1)]
            self.assertEqual(ledger.record(batch),(3,0))
            ledger = Ledger(path)
            self.assertEqual(ledger.record(batch),(0,3))
            with self.assertRaises(ValueError):
                ledger.record([sale("NEW","21/11/2026",1),sale("Z","21/11/2026",4)])
            for mode, expected in [("Ngày",(17,3,14)),("Tuần",(17,3,14)),("Tháng",(20,6,14))]:
                stamp, rows = ledger.report(*period("21/11/2026",mode))
                self.assertEqual(stamp,"2026-10-31")
                self.assertEqual(rows[0][3:],expected)
            ledger.snapshot("20/11/2026",{"SKU":("Gel","Gói","10")})
            self.assertEqual(ledger.report(*period("21/11/2026","Ngày"))[1][0][-1],7)
