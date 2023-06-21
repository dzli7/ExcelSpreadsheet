import context
from sheets.workbook import Workbook
from sheets.sheet import Sheet
from sheets.cellerror import CellError
from sheets.cellerrortype import CellErrorType

import unittest
from decimal import Decimal
from pyinstrument import Profiler

class SmallCircularStressTest(unittest.TestCase):
    def test_small_cycle_10(self):
        wb = Workbook()
        p = Profiler()

        _, name = wb.new_sheet()
        for i in range(10):
            wb.set_cell_contents(name, f"A{i}", f"=B{i}")
            wb.set_cell_contents(name, f"B{i}", f"=A{i} + B10")

        p.start()
        wb.set_cell_contents(name, "B10", "5")
        p.stop()

        p.print()

    def test_small_cycle_100(self):
        wb = Workbook()
        p = Profiler()

        _, name = wb.new_sheet()
        for i in range(100):
            wb.set_cell_contents(name, f"A{i}", f"=B{i}")
            wb.set_cell_contents(name, f"B{i}", f"=A{i} + B100")

        p.start()
        wb.set_cell_contents(name, "B100", "5")
        p.stop()

        p.print()

    def test_small_cycle_1000(self):
        wb = Workbook()
        p = Profiler()

        _, name = wb.new_sheet()
        for i in range(1000):
            wb.set_cell_contents(name, f"A{i}", f"=B{i}")
            wb.set_cell_contents(name, f"B{i}", f"=A{i} + B1000")

        p.start()
        wb.set_cell_contents(name, "B1000", "5")
        p.stop()

        p.print()



if __name__ == "__main__":
    unittest.main()