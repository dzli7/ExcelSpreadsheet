import context
from sheets.workbook import Workbook
from sheets.sheet import Sheet
from sheets.cellerror import CellError
from sheets.cellerrortype import CellErrorType

import unittest
from decimal import Decimal
from pyinstrument import Profiler

class LoadWorkbookStressTest(unittest.TestCase):
    def test_load_20(self):
        wb = Workbook()
        p = Profiler()

        _, name = wb.new_sheet()
        N = 20
        for i in range(1, N+1):
            wb.set_cell_contents(name, f"A{i}", f"{i}")

        with open("tests/test_json_objects/stress_20.json", 'w') as f:
            wb.save_workbook(f)

        p.start()
        wb2 = Workbook.load_workbook("tests/test_json_objects/stress_20.json")
        p.stop()

        p.print(show_all=True)

    def test_load_100(self):
        wb = Workbook()
        p = Profiler()

        _, name = wb.new_sheet()
        N = 100
        for i in range(1, N+1):
            wb.set_cell_contents(name, f"A{i}", f"{i}")

        with open("tests/test_json_objects/stress_100.json", 'w') as f:
            wb.save_workbook(f)

        p.start()
        wb2 = Workbook.load_workbook("tests/test_json_objects/stress_100.json")
        p.stop()

        p.print(show_all=True)

    def test_load_2000(self):
        wb = Workbook()
        p = Profiler()

        _, name = wb.new_sheet()
        N = 2000
        for i in range(1, N+1):
            wb.set_cell_contents(name, f"A{i}", f"{i}")

        with open("tests/test_json_objects/stress_1000.json", 'w') as f:
            wb.save_workbook(f)

        p.start()
        wb2 = Workbook.load_workbook("tests/test_json_objects/stress_1000.json")
        p.stop()

        with open('logs/project5-perf-log-load.txt', 'w') as f:
            p.print(file=f, show_all=True)



if __name__ == "__main__":
    unittest.main()