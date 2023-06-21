import context
from sheets.workbook import Workbook
from sheets.sheet import Sheet
from sheets.cellerror import CellError
from sheets.cellerrortype import CellErrorType

import unittest
from decimal import Decimal
from pyinstrument import Profiler

class RenameStressTest(unittest.TestCase):
    # def test_rename_20(self):
    #     wb = Workbook()
    #     p = Profiler()

    #     _, name1 = wb.new_sheet()
    #     _, name2 = wb.new_sheet()
    #     for i in range(2, 21):
    #         wb.set_cell_contents(name1, f"A{i}", f"{i}")
    #         wb.set_cell_contents(name2, f"A{i}", f"={name1}!A{i}")
    #         assert (wb.get_cell_value(name1, f"A{i}") == Decimal(f"{i}"))
    #         assert (wb.get_cell_value(name2, f"A{i}") == Decimal(f"{i}"))

    #     p.start()
    #     wb.rename_sheet(name1, "July Totals")
    #     p.stop()

    #     for i in range(2, 21):
    #         assert (wb.get_cell_value(name2, f"A{i}") == Decimal(f"{i}"))

    #     p.print()

    # def test_rename_100(self):
    #     wb = Workbook()
    #     p = Profiler()

    #     _, name1 = wb.new_sheet()
    #     _, name2 = wb.new_sheet()
    #     for i in range(2, 101):
    #         wb.set_cell_contents(name1, f"A{i}", f"{i}")
    #         wb.set_cell_contents(name2, f"A{i}", f"={name1}!A{i}")
    #         assert (wb.get_cell_value(name1, f"A{i}") == Decimal(f"{i}"))
    #         assert (wb.get_cell_value(name2, f"A{i}") == Decimal(f"{i}"))

    #     p.start()
    #     wb.rename_sheet(name1, "July Totals")
    #     p.stop()

    #     for i in range(2, 101):
    #         assert (wb.get_cell_value(name2, f"A{i}") == Decimal(f"{i}"))

    #     p.print(show_all=True)

    def test_rename_2000(self):
        wb = Workbook()
        p = Profiler()

        _, name1 = wb.new_sheet()
        _, name2 = wb.new_sheet()
        for i in range(2, 2001):
            wb.set_cell_contents(name1, f"A{i}", f"{i}")
            wb.set_cell_contents(name2, f"A{i}", f"={name1}!A{i}")
            assert (wb.get_cell_value(name1, f"A{i}") == Decimal(f"{i}"))
            assert (wb.get_cell_value(name2, f"A{i}") == Decimal(f"{i}"))

        p.start()
        wb.rename_sheet(name1, "July Totals")
        p.stop()

        for i in range(2, 2001):
            assert (wb.get_cell_value(name2, f"A{i}") == Decimal(f"{i}"))

        with open('logs/project5-perf-log-rename.txt', 'w') as f:
            p.print(file=f, show_all=True)


if __name__ == "__main__":
    unittest.main()