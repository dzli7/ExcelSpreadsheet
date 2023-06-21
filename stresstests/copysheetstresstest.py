import context
from sheets.workbook import Workbook
from sheets.sheet import Sheet
from sheets.cellerror import CellError
from sheets.cellerrortype import CellErrorType
from sheets.helper import column_string_from_index

import unittest
from decimal import Decimal
from pyinstrument import Profiler
from math import sqrt

class CopyCellStressTest(unittest.TestCase):
    def test_copy_sheet_20(self):
        wb = Workbook()
        p = Profiler()

        _, name = wb.new_sheet()
        N = 20
        M = int(sqrt(N))
        for j in range(1, M+1):
            for i in range(1, M+1):
                col = column_string_from_index(j)
                wb.set_cell_contents(name, f"{col}{i}", f"{i}")

        p.start()
        wb.copy_sheet(name)
        p.stop()

        p.print(show_all=True)

    def test_copy_sheet_100(self):
        wb = Workbook()
        p = Profiler()

        _, name = wb.new_sheet()
        N = 100
        M = int(sqrt(N))
        for j in range(1, M+1):
            for i in range(1, M+1):
                col = column_string_from_index(j)
                wb.set_cell_contents(name, f"{col}{i}", f"{i}")

        p.start()
        wb.copy_sheet(name)
        p.stop()

        p.print(show_all=True)

    def test_copy_sheet_2000(self):
        wb = Workbook()
        p = Profiler()

        _, name = wb.new_sheet()
        N = 2000
        M = int(sqrt(N))
        for j in range(1, M+1):
            for i in range(1, M+1):
                col = column_string_from_index(j)
                wb.set_cell_contents(name, f"{col}{i}", f"{i}")

        p.start()
        wb.copy_sheet(name)
        p.stop()

        with open('logs/project5-perf-log-copysheet.txt', 'w') as f:
            p.print(file=f, show_all=True)


if __name__ == "__main__":
    unittest.main()