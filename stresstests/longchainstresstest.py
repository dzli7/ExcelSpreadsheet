import context
from sheets.workbook import Workbook
from sheets.sheet import Sheet
from sheets.cellerror import CellError
from sheets.cellerrortype import CellErrorType

import unittest
from decimal import Decimal
from pyinstrument import Profiler

class LongChainStressTest(unittest.TestCase):
    def test_chain_20(self):
        wb = Workbook()
        p = Profiler()

        _, name = wb.new_sheet()
        for i in range(2, 21):
            wb.set_cell_contents(name, f"A{i}", f"=A{i-1}")

        p.start()
        wb.set_cell_contents(name, "A1", "5")
        p.stop()

        p.print(show_all=True)

    def test_chain_100(self):
        wb = Workbook()
        p = Profiler()

        _, name = wb.new_sheet()
        for i in range(2, 101):
            wb.set_cell_contents(name, f"A{i}", f"=A{i-1}")

        p.start()
        wb.set_cell_contents(name, "A1", "5")
        p.stop()

        p.print(show_all=True)

    def test_chain_1000(self):
        wb = Workbook()
        p = Profiler()

        _, name = wb.new_sheet()
        for i in range(2, 1001):
            wb.set_cell_contents(name, f"A{i}", f"=A{i-1}")

        p.start()
        wb.set_cell_contents(name, "A1", "5")
        p.stop()

        p.print(show_all=True)


if __name__ == "__main__":
    unittest.main()