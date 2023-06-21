import context
from sheets.workbook import Workbook
from sheets.sheet import Sheet
from sheets.cellerror import CellError
from sheets.cellerrortype import CellErrorType
from sheets.helper import column_string_from_index

import unittest
from decimal import Decimal
from pyinstrument import Profiler
from pyinstrument.renderers import ConsoleRenderer

class PascalStressTest(unittest.TestCase):
    def test_pascal_30(self):
        wb = Workbook()
        p = Profiler()

        _, name = wb.new_sheet()

        N = 30
        for i in range(3, N+1):
            wb.set_cell_contents(name, f"A{i}", "1")
            col_str = column_string_from_index(i)
            wb.set_cell_contents(name, f"{col_str}1", "1")
        

        imax = N
        for j in range(2, N):
            i = 2
            while i < imax:
                col = column_string_from_index(j)
                col_up = column_string_from_index(j-1)

                left = wb.get_cell_value(name, f"{col}{i-1}")
                up = wb.get_cell_value(name, f"{col_up}{i}")

                # if left != 0 and up != 0:
                print(f"{col}{i} = {col}{i-1} + {col_up}{i}")
                wb.set_cell_contents(name, f"{col}{i}", f"={col}{i-1} + {col_up}{i}")

                i += 1
            imax -= 1

        p.start()
        wb.set_cell_contents(name, "A2", "1")
        wb.set_cell_contents(name, "B1", "1")
        p.stop()

        with open('logs/project5-perf-log-pascal.txt', 'w') as f:
            p.print(file=f, show_all=True)

    # def test_pascal_50(self):
    #     wb = Workbook()
    #     p = Profiler()

    #     _, name = wb.new_sheet()

    #     N = 50
    #     for i in range(3, N+1):
    #         wb.set_cell_contents(name, f"A{i}", "1")
    #         col_str = column_string_from_index(i)
    #         wb.set_cell_contents(name, f"{col_str}1", "1")
        

    #     imax = N
    #     for j in range(2, N):
    #         i = 2
    #         while i < imax:
    #             col = column_string_from_index(j)
    #             col_up = column_string_from_index(j-1)

    #             left = wb.get_cell_value(name, f"{col}{i-1}")
    #             up = wb.get_cell_value(name, f"{col_up}{i}")

    #             # if left != 0 and up != 0:
    #             # print(f"{col}{i} = {col}{i-1} + {col_up}{i}")
    #             wb.set_cell_contents(name, f"{col}{i}", f"={col}{i-1} + {col_up}{i}")

    #             i += 1
    #         imax -= 1

    #     p.start()
    #     wb.set_cell_contents(name, "A2", "1")
    #     wb.set_cell_contents(name, "B1", "1")
    #     p.stop()

    #     p.print(show_all=True)

    # def test_pascal_100(self):
    #     wb = Workbook()
    #     p = Profiler()

    #     _, name = wb.new_sheet()

    #     N = 100
    #     for i in range(3, N+1):
    #         wb.set_cell_contents(name, f"A{i}", "1")
    #         col_str = column_string_from_index(i)
    #         wb.set_cell_contents(name, f"{col_str}1", "1")
        

    #     imax = N
    #     for j in range(2, N):
    #         i = 2
    #         while i < imax:
    #             col = column_string_from_index(j)
    #             col_up = column_string_from_index(j-1)

    #             left = wb.get_cell_value(name, f"{col}{i-1}")
    #             up = wb.get_cell_value(name, f"{col_up}{i}")

    #             # if left != 0 and up != 0:
    #             # print(f"{col}{i} = {col}{i-1} + {col_up}{i}")
    #             wb.set_cell_contents(name, f"{col}{i}", f"={col}{i-1} + {col_up}{i}")

    #             i += 1
    #         imax -= 1

    #     p.start()
    #     wb.set_cell_contents(name, "A2", "1")
    #     wb.set_cell_contents(name, "B1", "1")
    #     p.stop()

    #     p.print(show_all=True)


if __name__ == "__main__":
    unittest.main()