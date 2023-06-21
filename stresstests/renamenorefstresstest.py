import context
from sheets.workbook import Workbook
from sheets.sheet import Sheet
from sheets.cellerror import CellError
from sheets.cellerrortype import CellErrorType

import unittest
from decimal import Decimal
from pyinstrument import Profiler

class RenameNoRefStressTest(unittest.TestCase):
    # def test_rename_noref_20(self):
    #     wb = Workbook()
    #     p = Profiler()

    #     _, name1 = wb.new_sheet()

    #     for i in range(2, 21):
    #         wb.set_cell_contents(name1, f"A{i}", f"{i}")
    #         wb.set_cell_contents(name1, f"B{i}", f"{i+1}")
    #         assert (wb.get_cell_value(name1, f"A{i}") == Decimal(f"{i}"))
    #         assert (wb.get_cell_value(name1, f"B{i}") == Decimal(f"{i+1}"))

    #     p.start()
    #     wb.rename_sheet(name1, "July Totals")
    #     p.stop()

    #     for i in range(2, 21):
    #         assert (wb.get_cell_value("July Totals", f"A{i}") == Decimal(f"{i}"))
    #         assert (wb.get_cell_value("July Totals", f"B{i}") == Decimal(f"{i+1}"))

    #     p.print()

    def test_rename_noref_100(self):
        wb = Workbook()
        p = Profiler()

        _, name1 = wb.new_sheet()

        for i in range(2, 101):
            wb.set_cell_contents(name1, f"A{i}", f"{i}")
            wb.set_cell_contents(name1, f"B{i}", f"{i+1}")
            assert (wb.get_cell_value(name1, f"A{i}") == Decimal(f"{i}"))
            assert (wb.get_cell_value(name1, f"B{i}") == Decimal(f"{i+1}"))

        p.start()
        wb.rename_sheet(name1, "July Totals")
        p.stop()

        for i in range(2, 101):
            assert (wb.get_cell_value("July Totals", f"A{i}") == Decimal(f"{i}"))
            assert (wb.get_cell_value("July Totals", f"B{i}") == Decimal(f"{i+1}"))

        p.print(show_all=True)

    def test_rename_noref_1000(self):
        wb = Workbook()
        p = Profiler()

        _, name1 = wb.new_sheet()

        for i in range(2, 1001):
            wb.set_cell_contents(name1, f"A{i}", f"{i}")
            wb.set_cell_contents(name1, f"B{i}", f"{i+1}")
            assert (wb.get_cell_value(name1, f"A{i}") == Decimal(f"{i}"))
            assert (wb.get_cell_value(name1, f"B{i}") == Decimal(f"{i+1}"))

        p.start()
        wb.rename_sheet(name1, "July Totals")
        p.stop()

        for i in range(2, 1001):
            assert (wb.get_cell_value("July Totals", f"A{i}") == Decimal(f"{i}"))
            assert (wb.get_cell_value("July Totals", f"B{i}") == Decimal(f"{i+1}"))

        with open('logs/project5-perf-log-rename-noref.txt', 'w') as f:
            p.print(file=f, show_all=True)


if __name__ == "__main__":
    unittest.main()