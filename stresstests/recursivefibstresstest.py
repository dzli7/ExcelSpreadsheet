import context
from sheets.workbook import Workbook
from sheets.sheet import Sheet
from sheets.cellerror import CellError
from sheets.cellerrortype import CellErrorType

import unittest
from decimal import Decimal
from pyinstrument import Profiler
from pyinstrument.renderers import ConsoleRenderer
import cProfile
from pstats import Stats

class RecursiveFibStressTest(unittest.TestCase):
    def test_fib_20(self):
        wb = Workbook()
        p = Profiler()

        _, name = wb.new_sheet()

        wb.set_cell_contents(name, "a2", "1")
        for i in range(3, 21):
            wb.set_cell_contents(name, f"A{i}", f"=A{i-1}+A{i-2}")

        p.start()
        wb.set_cell_contents(name, "a1", "1")
        p.stop()

        p.print(show_all=True)

    def test_fib_100(self):
        wb = Workbook()
        p = Profiler()

        _, name = wb.new_sheet()

        wb.set_cell_contents(name, "a2", "1")
        for i in range(3, 101):
            wb.set_cell_contents(name, f"A{i}", f"=A{i-1}+A{i-2}")

        p.start()
        wb.set_cell_contents(name, "a1", "1")
        p.stop()
        # session = p.stop()

        p.print(show_all=True)
        # profile_renderer = ConsoleRenderer(unicode=True, color=True, show_all=True)

        # print(profile_renderer.render(session))
    
    def test_fib_1000(self):
        wb = Workbook()
        p = Profiler()

        _, name = wb.new_sheet()

        wb.set_cell_contents(name, "a2", "1")
        for i in range(3, 501):
            wb.set_cell_contents(name, f"A{i}", f"=A{i-1}+A{i-2}")

        p.start()
        wb.set_cell_contents(name, "a1", "1")
        # p.stop()
        session = p.stop()

        # p.print(show_all=True)
        profile_renderer = ConsoleRenderer(unicode=True, color=True, show_all=True)

        print(profile_renderer.render(session))


if __name__ == "__main__":
    unittest.main()
    # profiler = cProfile.Profile()
    # profiler.enable()

    # test_fib_1000()

    # profiler.disable()
    # stats = Stats(profiler).sort_stats("cumtime")
    # stats.print_stats(200)