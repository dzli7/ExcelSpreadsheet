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
    def test_copy_cell_row_20(self):
        wb = Workbook()
        p = Profiler()

        _, name = wb.new_sheet()
        N = 20
        for i in range(1, N+1):
            wb.set_cell_contents(name, f"A{i}", f"{i}")

        p.start()
        wb.copy_cells(name, 'A1', f'A{N}', 'B1')
        p.stop()

        p.print(show_all=True)

    def test_copy_cell_row_100(self):
        wb = Workbook()
        p = Profiler()

        _, name = wb.new_sheet()
        N = 100
        for i in range(1, N+1):
            wb.set_cell_contents(name, f"A{i}", f"{i}")

        p.start()
        wb.copy_cells(name, 'A1', f'A{N}', 'B1')
        p.stop()

        p.print(show_all=True)

    def test_copy_cell_2000(self):
        wb = Workbook()
        p = Profiler()

        _, name = wb.new_sheet()
        N = 2000
        M = int(sqrt(N))
        for j in range(1, M+1):
            for i in range(1, M+1):
                col = column_string_from_index(j)
                wb.set_cell_contents(name, f"{col}{i}", f"{i}")

        cmax = column_string_from_index(M)
        cmax1 = column_string_from_index(M+1)
        p.start()
        wb.copy_cells(name, 'A1', f'{cmax}{M}', f'{cmax1}1')
        p.stop()

        with open('logs/project5-perf-log-copycell.txt', 'w') as f:
            p.print(file=f, show_all=True)

    def test_copy_cell_diff_sheet_2000(self):
        wb = Workbook()
        p = Profiler()

        _, name = wb.new_sheet()
        N = 2000
        M = int(sqrt(N))
        for j in range(1, M+1):
            for i in range(1, M+1):
                col = column_string_from_index(j)
                wb.set_cell_contents(name, f"{col}{i}", f"{i}")

        cmax = column_string_from_index(M)
        cmax1 = column_string_from_index(M+1)
        _, name2 = wb.new_sheet()
        p.start()
        wb.copy_cells(name, 'A1', f'{cmax}{M}', f'{cmax1}1', to_sheet=name2)
        p.stop()

        with open('logs/project5-perf-log-copycell-diffsheet.txt', 'w') as f:
            p.print(file=f, show_all=True)


if __name__ == "__main__":
    unittest.main()