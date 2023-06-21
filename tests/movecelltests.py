import context
from sheets.workbook import Workbook
from sheets.sheet import Sheet
from sheets.cellerror import CellError
from sheets.cellerrortype import CellErrorType
import json
import io

import unittest
from decimal import Decimal

# Test Suite for Workbooks


class MoveCellTests(unittest.TestCase):
    def test_move_row_of_cells(self):
        wb = Workbook()
        _, name = wb.new_sheet()
        wb.set_cell_contents(name, "A1", "1")
        wb.set_cell_contents(name, "B1", "hi")
        wb.set_cell_contents(name, "C1", "'3")
        wb.set_cell_contents(name, "D1", "70")
        wb.move_cells(name, 'A1', 'D1', 'A2')
        assert wb.get_cell_value(name, "A2") == Decimal('1')
        assert wb.get_cell_value(name, "B2") == "hi"
        assert wb.get_cell_value(name, "C2") == Decimal('3')
        assert wb.get_cell_value(name, "D2") == Decimal('70')

        assert wb.get_cell_value(name, "A1") == None
        assert wb.get_cell_value(name, "A1") == None
        assert wb.get_cell_value(name, "A1") == None
        assert wb.get_cell_value(name, "A1") == None

        wb.move_cells(name, 'D2', 'A2', 'E1')
        assert wb.get_cell_value(name, "A2") == None
        assert wb.get_cell_value(name, "B2") == None
        assert wb.get_cell_value(name, "C2") == None
        assert wb.get_cell_value(name, "D2") == None

        assert wb.get_cell_value(name, "E1") == Decimal('1')
        assert wb.get_cell_value(name, "F1") == "hi"
        assert wb.get_cell_value(name, "G1") == Decimal('3')
        assert wb.get_cell_value(name, "H1") == Decimal('70')

        wb.move_cells(name, 'E1', 'H1', 'AA1')
        assert wb.get_cell_value(name, "AA1") == Decimal('1')
        assert wb.get_cell_value(name, "AB1") == "hi"
        assert wb.get_cell_value(name, "AC1") == Decimal('3')
        assert wb.get_cell_value(name, "AD1") == Decimal('70')

        assert wb.get_cell_value(name, "E1") == None
        assert wb.get_cell_value(name, "F1") == None
        assert wb.get_cell_value(name, "G1") == None
        assert wb.get_cell_value(name, "H1") == None


    def test_copy_column_of_cells(self):
        wb = Workbook()
        _, name = wb.new_sheet()
        wb.set_cell_contents(name, "A1", "1")
        wb.set_cell_contents(name, "A2", "hi")
        wb.set_cell_contents(name, "A3", "'3")
        wb.set_cell_contents(name, "A4", "70")
        wb.copy_cells(name, 'A4', 'A1', 'B1')
        assert wb.get_cell_value(name, "B1") == wb.get_cell_value(name, "A1")
        assert wb.get_cell_value(name, "B2") == wb.get_cell_value(name, "A2")
        assert wb.get_cell_value(name, "B3") == wb.get_cell_value(name, "A3")
        assert wb.get_cell_value(name, "B4") == wb.get_cell_value(name, "A4")

        assert wb.get_cell_contents(name, "B1") == wb.get_cell_contents(name, "A1")
        assert wb.get_cell_contents(name, "B2") == wb.get_cell_contents(name, "A2")
        assert wb.get_cell_contents(name, "B3") == wb.get_cell_contents(name, "A3")
        assert wb.get_cell_contents(name, "B4") == wb.get_cell_contents(name, "A4")
        wb.copy_cells(name, 'B1', 'B4', 'AA1')
        assert wb.get_cell_value(name, "AA1") == wb.get_cell_value(name, "A1")
        assert wb.get_cell_value(name, "AA2") == wb.get_cell_value(name, "A2")
        assert wb.get_cell_value(name, "AA3") == wb.get_cell_value(name, "A3")
        assert wb.get_cell_value(name, "AA4") == wb.get_cell_value(name, "A4")

        assert wb.get_cell_contents(name, "AA1") == wb.get_cell_contents(name, "A1")
        assert wb.get_cell_contents(name, "AA2") == wb.get_cell_contents(name, "A2")
        assert wb.get_cell_contents(name, "AA3") == wb.get_cell_contents(name, "A3")
        assert wb.get_cell_contents(name, "AA4") == wb.get_cell_contents(name, "A4")

        wb.copy_cells(name, 'AA1', 'AA4', 'A5')
        assert wb.get_cell_value(name, "A5") == wb.get_cell_value(name, "A1")
        assert wb.get_cell_value(name, "A6") == wb.get_cell_value(name, "A2")
        assert wb.get_cell_value(name, "A7") == wb.get_cell_value(name, "A3")
        assert wb.get_cell_value(name, "A8") == wb.get_cell_value(name, "A4")

        assert wb.get_cell_contents(name, "A5") == wb.get_cell_contents(name, "A1")
        assert wb.get_cell_contents(name, "A6") == wb.get_cell_contents(name, "A2")
        assert wb.get_cell_contents(name, "A7") == wb.get_cell_contents(name, "A3")
        assert wb.get_cell_contents(name, "A8") == wb.get_cell_contents(name, "A4")

    def test_copy_two_by_two(self):
        wb = Workbook()
        _, name = wb.new_sheet()
        wb.set_cell_contents(name, "A1", "1")
        wb.set_cell_contents(name, "A2", "hi")
        wb.set_cell_contents(name, "B1", "'3")
        wb.set_cell_contents(name, "B2", "70")
        wb.copy_cells(name, 'A1', 'B2', 'C1')
        assert wb.get_cell_value(name, "C1") == wb.get_cell_value(name, "A1")
        assert wb.get_cell_value(name, "C2") == wb.get_cell_value(name, "A2")
        assert wb.get_cell_value(name, "D1") == wb.get_cell_value(name, "B1")
        assert wb.get_cell_value(name, "D2") == wb.get_cell_value(name, "B2")

        assert wb.get_cell_contents(name, "C1") == wb.get_cell_contents(name, "A1")
        assert wb.get_cell_contents(name, "C2") == wb.get_cell_contents(name, "A2")
        assert wb.get_cell_contents(name, "D1") == wb.get_cell_contents(name, "B1")
        assert wb.get_cell_contents(name, "D2") == wb.get_cell_contents(name, "B2")

        wb.copy_cells(name, 'D2', 'C1', 'A3')
        assert wb.get_cell_value(name, "A3") == wb.get_cell_value(name, "A1")
        assert wb.get_cell_value(name, "A4") == wb.get_cell_value(name, "A2")
        assert wb.get_cell_value(name, "B3") == wb.get_cell_value(name, "B1")
        assert wb.get_cell_value(name, "B4") == wb.get_cell_value(name, "B2")

        assert wb.get_cell_contents(name, "A3") == wb.get_cell_contents(name, "A1")
        assert wb.get_cell_contents(name, "A4") == wb.get_cell_contents(name, "A2")
        assert wb.get_cell_contents(name, "B3") == wb.get_cell_contents(name, "B1")
        assert wb.get_cell_contents(name, "B4") == wb.get_cell_contents(name, "B2")

    def test_copy_with_overlap(self):
        wb = Workbook()
        _, name = wb.new_sheet()
        wb.set_cell_contents(name, "A1", "1")
        wb.set_cell_contents(name, "A2", "hi")
        wb.set_cell_contents(name, "A3", "17")
        wb.set_cell_contents(name, "B1", "'3")
        wb.set_cell_contents(name, "B2", "70")
        wb.set_cell_contents(name, "B3", "68")
        wb.set_cell_contents(name, "C1", "bye")
        wb.set_cell_contents(name, "C2", "slatt")
        wb.set_cell_contents(name, "C3", "'900")
        wb.move_cells(name, "A1", "C3", "B2")
        # Set to None after cells moved
        assert wb.get_cell_value(name, "A1") == None
        assert wb.get_cell_value(name, "A2") == None
        assert wb.get_cell_value(name, "A3") == None
        assert wb.get_cell_value(name, "B1") == None
        assert wb.get_cell_value(name, "C1") == None
        # Changed in overlap
        assert wb.get_cell_value(name, "B2") == Decimal('1')
        assert wb.get_cell_value(name, "B3") == "hi"
        assert wb.get_cell_value(name, "C2") == Decimal('3')
        assert wb.get_cell_value(name, "C3") == Decimal('70')
        # Newly "added"
        assert wb.get_cell_value(name, "B4") == Decimal('17')
        assert wb.get_cell_value(name, "C4") == Decimal('68')
        assert wb.get_cell_value(name, "D4") == Decimal('900')
        assert wb.get_cell_value(name, "D2") == "bye"
        assert wb.get_cell_value(name, "D3") == "slatt"

    def test_move_cells_second_sheet(self):
        wb = Workbook()
        wb.new_sheet()
        wb.new_sheet()
        wb.set_cell_contents("Sheet1", "A1", "1")
        wb.set_cell_contents("Sheet1", "A2", "hi")
        wb.set_cell_contents("Sheet1", "B1", "'3")
        wb.set_cell_contents("Sheet1", "B2", "70")
        wb.move_cells("Sheet1", "B2", "A1", "C1", "Sheet2")
        assert wb.get_cell_value("Sheet2", "C1") == Decimal('1')
        assert wb.get_cell_value("Sheet2", "C2") == "hi"
        assert wb.get_cell_value("Sheet2", "D1") == Decimal('3')
        assert wb.get_cell_value("Sheet2", "D2") == Decimal('70')

    def test_incorrect_corner_order(self):
        wb = Workbook()
        wb.new_sheet()
        wb.set_cell_contents("Sheet1", "A1", "1")
        wb.set_cell_contents("Sheet1", "A2", "hi")
        wb.set_cell_contents("Sheet1", "B1", "'3")
        wb.set_cell_contents("Sheet1", "B2", "slatt")
        wb.move_cells("Sheet1", "B2", "A1", "D1")
        assert wb.get_cell_value("Sheet1", "D1") == Decimal('1')
        assert wb.get_cell_value("Sheet1", "D2") == "hi"
        assert wb.get_cell_value("Sheet1", "E1") == Decimal('3')
        assert wb.get_cell_value("Sheet1", "E2") == "slatt"
        wb.copy_cells("Sheet1", "E2", "D1", "G1")
        assert wb.get_cell_value("Sheet1", "G1") == Decimal('1')
        assert wb.get_cell_value("Sheet1", "G2") == "hi"
        assert wb.get_cell_value("Sheet1", "H1") == Decimal('3')
        assert wb.get_cell_value("Sheet1", "H2") == "slatt"

        wb.new_sheet()
        wb.set_cell_contents("Sheet2", "A1", "1")
        wb.set_cell_contents("Sheet2", "A2", "hi")
        wb.set_cell_contents("Sheet2", "A3", "'3")
        wb.set_cell_contents("Sheet2", "A4", "slatt")
        wb.move_cells("Sheet2", "A4", "A1", "B1")
        assert wb.get_cell_value("Sheet2", "B1") == Decimal('1')
        assert wb.get_cell_value("Sheet2", "B2") == "hi"
        assert wb.get_cell_value("Sheet2", "B3") == Decimal('3')
        assert wb.get_cell_value("Sheet2", "B4") == "slatt"

        wb.new_sheet()
        wb.set_cell_contents("Sheet3", "A1", "1")
        wb.set_cell_contents("Sheet3", "B1", "hi")
        wb.set_cell_contents("Sheet3", "C1", "'3")
        wb.set_cell_contents("Sheet3", "D1", "slatt")
        wb.move_cells("Sheet3", "D1", "A1", "A2")
        assert wb.get_cell_value("Sheet3", "A2") == Decimal('1')
        assert wb.get_cell_value("Sheet3", "B2") == "hi"
        assert wb.get_cell_value("Sheet3", "C2") == Decimal('3')
        assert wb.get_cell_value("Sheet3", "D2") == "slatt"


if __name__ == "__main__":
    unittest.main()