import context
from sheets.workbook import Workbook
from sheets.sheet import Sheet
from sheets.cellerror import CellError
from sheets.cellerrortype import CellErrorType
from sheets.helper import correct_error

import json
import io

import unittest
from decimal import Decimal

# Test Suite for Workbooks


class SortingTests(unittest.TestCase):
    def test_move_rows_of_cells(self):
        # Test sorting for ascending and descending
        wb = Workbook()
        wb.new_sheet()
        wb.set_cell_contents("Sheet1", "A1", "2")
        wb.set_cell_contents("Sheet1", "A2", "1")
        wb.set_cell_contents("Sheet1", "A3", "77")
        wb.set_cell_contents("Sheet1", "A4", "40")
        wb.set_cell_contents("Sheet1", "B1", "one")
        wb.set_cell_contents("Sheet1", "B2", "two")
        wb.set_cell_contents("Sheet1", "B3", "three")
        wb.set_cell_contents("Sheet1", "B4", "four")
        wb.sort_region("Sheet1", "A1", "B4", [1])
        assert wb.get_cell_value("Sheet1", "A1") == Decimal("1")
        assert wb.get_cell_value("Sheet1", "A2") == Decimal("2")
        assert wb.get_cell_value("Sheet1", "A3") == Decimal("40")
        assert wb.get_cell_value("Sheet1", "A4") == Decimal("77")
        assert wb.get_cell_value("Sheet1", "B1") == "two"
        assert wb.get_cell_value("Sheet1", "B2") == "one"
        assert wb.get_cell_value("Sheet1", "B3") == "four"
        assert wb.get_cell_value("Sheet1", "B4") == "three"
        wb.sort_region("Sheet1", "A1", "B4", [-1])
        assert wb.get_cell_value("Sheet1", "A1") == Decimal("77")
        assert wb.get_cell_value("Sheet1", "A2") == Decimal("40")
        assert wb.get_cell_value("Sheet1", "A3") == Decimal("2")
        assert wb.get_cell_value("Sheet1", "A4") == Decimal("1")
        assert wb.get_cell_value("Sheet1", "B1") == "three"
        assert wb.get_cell_value("Sheet1", "B2") == "four"
        assert wb.get_cell_value("Sheet1", "B3") == "one"
        assert wb.get_cell_value("Sheet1", "B4") == "two"


    def test_precedence_of_objects(self):
        wb = Workbook()
        wb.new_sheet()
        wb.set_cell_contents("Sheet1", "A1", "2")
        wb.set_cell_contents("Sheet1", "A2", "True")
        wb.set_cell_contents("Sheet1", "A3", "slatt")
        wb.set_cell_contents("Sheet1", "A4", "=40/0")
        wb.sort_region("Sheet1", "A1", "A5", [1])
        assert wb.get_cell_value("Sheet1", "A1") == None
        assert correct_error(wb.get_cell_value("Sheet1", "A2"), CellErrorType.DIVIDE_BY_ZERO)
        assert wb.get_cell_value("Sheet1", "A3") == Decimal("2")
        assert wb.get_cell_value("Sheet1", "A4") == "slatt"
        assert wb.get_cell_value("Sheet1", "A5") == True

    def test_precedence_of_objects_neg(self):
        wb = Workbook()
        wb.new_sheet()
        wb.set_cell_contents("Sheet1", "A1", "2")
        wb.set_cell_contents("Sheet1", "A2", "True")
        wb.set_cell_contents("Sheet1", "A3", "slatt")
        wb.set_cell_contents("Sheet1", "A4", "=40/0")
        wb.sort_region("Sheet1", "A1", "A5", [-1])
        assert wb.get_cell_value("Sheet1", "A1") == True
        assert wb.get_cell_value("Sheet1", "A2") == "slatt"
        assert wb.get_cell_value("Sheet1", "A3") == Decimal("2")
        assert correct_error(wb.get_cell_value("Sheet1", "A4"), CellErrorType.DIVIDE_BY_ZERO)
        assert wb.get_cell_value("Sheet1", "A5") == None

    def test_sorting_with_formula(self):
        wb = Workbook()
        wb.new_sheet()
        wb.set_cell_contents("Sheet1", "A1", "2")
        wb.set_cell_contents("Sheet1", "A2", "1")
        wb.set_cell_contents("Sheet1", "A3", "=B3")
        wb.set_cell_contents("Sheet1", "A4", "-2")
        wb.set_cell_contents("Sheet1", "B1", "bye")
        wb.set_cell_contents("Sheet1", "B2", "hi")
        wb.set_cell_contents("Sheet1", "B3", "70")
        wb.set_cell_contents("Sheet1", "B4", "slatt")

        wb.sort_region("Sheet1", "A1", "B4", [1])
        assert wb.get_cell_value("Sheet1", "A1") == Decimal("-2")
        assert wb.get_cell_value("Sheet1", "A2") == Decimal("1")
        assert wb.get_cell_value("Sheet1", "A3") == Decimal("2")
        assert wb.get_cell_value("Sheet1", "A4") == Decimal("70")
        assert wb.get_cell_value("Sheet1", "B1") == "slatt"
        assert wb.get_cell_value("Sheet1", "B2") == "hi"
        assert wb.get_cell_value("Sheet1", "B3") == "bye"
        assert wb.get_cell_value("Sheet1", "B4") == Decimal("70")
        assert wb.get_cell_contents("Sheet1", "A4") == "=B4"

    def test_equality_present(self):
        wb = Workbook()
        wb.new_sheet()
        wb.set_cell_contents("Sheet1", "A1", "2000")
        wb.set_cell_contents("Sheet1", "A2", "200")
        wb.set_cell_contents("Sheet1", "A3", "200")
        wb.set_cell_contents("Sheet1", "A4", "-2")
        wb.set_cell_contents("Sheet1", "B1", "10")
        wb.set_cell_contents("Sheet1", "B2", "7")
        wb.set_cell_contents("Sheet1", "B3", "3")
        wb.set_cell_contents("Sheet1", "B4", "1")
        wb.sort_region("Sheet1", "A1", "B4", [1, 2])
        assert wb.get_cell_value("Sheet1", "A1") == Decimal("-2")
        assert wb.get_cell_value("Sheet1", "A2") == Decimal("200")
        assert wb.get_cell_value("Sheet1", "A3") == Decimal("200")
        assert wb.get_cell_value("Sheet1", "A4") == Decimal("2000")
        assert wb.get_cell_value("Sheet1", "B1") == Decimal("1")
        assert wb.get_cell_value("Sheet1", "B2") == Decimal("3")
        assert wb.get_cell_value("Sheet1", "B3") == Decimal("7")
        assert wb.get_cell_value("Sheet1", "B4") == Decimal("10")
    
    def test_equality_present_neg(self):
        wb = Workbook()
        wb.new_sheet()
        wb.set_cell_contents("Sheet1", "A1", "2000")
        wb.set_cell_contents("Sheet1", "A2", "200")
        wb.set_cell_contents("Sheet1", "A3", "200")
        wb.set_cell_contents("Sheet1", "A4", "-2")
        wb.set_cell_contents("Sheet1", "A5", "200")
        wb.set_cell_contents("Sheet1", "B1", "10")
        wb.set_cell_contents("Sheet1", "B2", "7")
        wb.set_cell_contents("Sheet1", "B3", "3")
        wb.set_cell_contents("Sheet1", "B4", "1")
        wb.set_cell_contents("Sheet1", "B5", "6")

        wb.sort_region("Sheet1", "A1", "B5", [1, -2])
        assert wb.get_cell_value("Sheet1", "A1") == Decimal("-2")
        assert wb.get_cell_value("Sheet1", "A2") == Decimal("200")
        assert wb.get_cell_value("Sheet1", "A3") == Decimal("200")
        assert wb.get_cell_value("Sheet1", "A4") == Decimal("200")
        assert wb.get_cell_value("Sheet1", "A5") == Decimal("2000")
        assert wb.get_cell_value("Sheet1", "B1") == Decimal("1")
        assert wb.get_cell_value("Sheet1", "B2") == Decimal("7")
        assert wb.get_cell_value("Sheet1", "B3") == Decimal("6")
        assert wb.get_cell_value("Sheet1", "B4") == Decimal("3")
        assert wb.get_cell_value("Sheet1", "B5") == Decimal("10")

    def test_sort_strings(self):
        wb = Workbook()
        wb.new_sheet()
        wb.set_cell_contents("Sheet1", "A1", "SLATT")
        wb.set_cell_contents("Sheet1", "A2", "apple")
        wb.set_cell_contents("Sheet1", "A3", "zebra")
        wb.set_cell_contents("Sheet1", "A4", "SlaTt")
        wb.set_cell_contents("Sheet1", "A5", "slatt")
        wb.set_cell_contents("Sheet1", "B1", "1")
        wb.set_cell_contents("Sheet1", "B2", "2")
        wb.set_cell_contents("Sheet1", "B3", "3")
        wb.set_cell_contents("Sheet1", "B4", "4")
        wb.set_cell_contents("Sheet1", "B5", "5")
        wb.sort_region("Sheet1", "A1", "B5", [1, 2])
        assert wb.get_cell_value("Sheet1", "A1") == "apple"
        assert wb.get_cell_value("Sheet1", "A2") == "SLATT"
        assert wb.get_cell_value("Sheet1", "A3") == "SlaTt"
        assert wb.get_cell_value("Sheet1", "A4") == "slatt"
        assert wb.get_cell_value("Sheet1", "A5") == "zebra"
        assert wb.get_cell_value("Sheet1", "B1") == Decimal("2")
        assert wb.get_cell_value("Sheet1", "B2") == Decimal("1")
        assert wb.get_cell_value("Sheet1", "B3") == Decimal("4")
        assert wb.get_cell_value("Sheet1", "B4") == Decimal("5")
        assert wb.get_cell_value("Sheet1", "B5") == Decimal("3")        

    def test_sort_strings_neg(self):
        wb = Workbook()
        wb.new_sheet()
        wb.set_cell_contents("Sheet1", "A1", "SLATT")
        wb.set_cell_contents("Sheet1", "A2", "apple")
        wb.set_cell_contents("Sheet1", "A3", "zebra")
        wb.set_cell_contents("Sheet1", "A4", "SLATT")
        wb.set_cell_contents("Sheet1", "A5", "slatt")
        wb.set_cell_contents("Sheet1", "B1", "1")
        wb.set_cell_contents("Sheet1", "B2", "2")
        wb.set_cell_contents("Sheet1", "B3", "3")
        wb.set_cell_contents("Sheet1", "B4", "4")
        wb.set_cell_contents("Sheet1", "B5", "5")
        wb.sort_region("Sheet1", "A1", "B5", [1, -2])
        assert wb.get_cell_value("Sheet1", "A1") == "apple"
        assert wb.get_cell_value("Sheet1", "A2") == "slatt"
        assert wb.get_cell_value("Sheet1", "A3") == "SLATT"
        assert wb.get_cell_value("Sheet1", "A4") == "SLATT"
        assert wb.get_cell_value("Sheet1", "A5") == "zebra"
        assert wb.get_cell_value("Sheet1", "B1") == Decimal("2")
        assert wb.get_cell_value("Sheet1", "B2") == Decimal("5")
        assert wb.get_cell_value("Sheet1", "B3") == Decimal("4")
        assert wb.get_cell_value("Sheet1", "B4") == Decimal("1")
        assert wb.get_cell_value("Sheet1", "B5") == Decimal("3")    

    def test_stable_sort(self):
        wb = Workbook()
        wb.new_sheet()
        wb.set_cell_contents("Sheet1", "A1", "0")
        wb.set_cell_contents("Sheet1", "A2", "slatt")
        wb.set_cell_contents("Sheet1", "A3", "10")
        wb.set_cell_contents("Sheet1", "A4", "20")
        wb.set_cell_contents("Sheet1", "A5", "zoinks")
        wb.set_cell_contents("Sheet1", "A6", "10")
        wb.set_cell_contents("Sheet1", "A7", "-100")
        wb.set_cell_contents("Sheet1", "A8", "slatt")
        wb.set_cell_contents("Sheet1", "A9", "apple")
        wb.set_cell_contents("Sheet1", "B1", "1")
        wb.set_cell_contents("Sheet1", "B2", "2")
        wb.set_cell_contents("Sheet1", "B3", "3")
        wb.set_cell_contents("Sheet1", "B4", "4")
        wb.set_cell_contents("Sheet1", "B5", "5")
        wb.set_cell_contents("Sheet1", "B6", "6")
        wb.set_cell_contents("Sheet1", "B7", "7")
        wb.set_cell_contents("Sheet1", "B8", "8")
        wb.set_cell_contents("Sheet1", "B9", "9")
        wb.sort_region("Sheet1", "A1", "B9", [1])
        assert wb.get_cell_value("Sheet1", "A1") == Decimal("-100")
        assert wb.get_cell_value("Sheet1", "A2") == Decimal("0")
        assert wb.get_cell_value("Sheet1", "A3") == Decimal("10")
        assert wb.get_cell_value("Sheet1", "A4") == Decimal("10")
        assert wb.get_cell_value("Sheet1", "A5") == Decimal("20")
        assert wb.get_cell_value("Sheet1", "A6") == "apple"
        assert wb.get_cell_value("Sheet1", "A7") == "slatt"
        assert wb.get_cell_value("Sheet1", "A8") == "slatt"
        assert wb.get_cell_value("Sheet1", "A9") == "zoinks"

        assert wb.get_cell_value("Sheet1", "B1") == Decimal("7")
        assert wb.get_cell_value("Sheet1", "B2") == Decimal("1")
        assert wb.get_cell_value("Sheet1", "B3") == Decimal("3")
        assert wb.get_cell_value("Sheet1", "B4") == Decimal("6")
        assert wb.get_cell_value("Sheet1", "B5") == Decimal("4")
        assert wb.get_cell_value("Sheet1", "B6") == Decimal("9")
        assert wb.get_cell_value("Sheet1", "B7") == Decimal("2")
        assert wb.get_cell_value("Sheet1", "B8") == Decimal("8")
        assert wb.get_cell_value("Sheet1", "B9") == Decimal("5")

    def test_sort_errors(self):
        wb = Workbook()
        wb.new_sheet()
        wb.set_cell_contents("Sheet1", "A1", "=A1")
        wb.set_cell_contents("Sheet1", "A2", "=Sheet20!A2")
        wb.set_cell_contents("Sheet1", "A3", "=@")
        wb.set_cell_contents("Sheet1", "A4", "=10/0")
        wb.set_cell_contents("Sheet1", "A5", "=EXACT(A1, A2, A3, A4)")
        wb.set_cell_contents("Sheet1", "A6", "=LOW(True, True, False)")
        wb.sort_region("Sheet1", "A1", "A6", [1])

        assert correct_error(wb.get_cell_value("Sheet1", "A1"), CellErrorType.PARSE_ERROR)
        assert correct_error(wb.get_cell_value("Sheet1", "A2"), CellErrorType.CIRCULAR_REFERENCE)
        assert correct_error(wb.get_cell_value("Sheet1", "A3"), CellErrorType.BAD_REFERENCE)
        assert correct_error(wb.get_cell_value("Sheet1", "A4"), CellErrorType.BAD_NAME)
        assert correct_error(wb.get_cell_value("Sheet1", "A5"), CellErrorType.TYPE_ERROR)
        assert correct_error(wb.get_cell_value("Sheet1", "A6"), CellErrorType.DIVIDE_BY_ZERO)
        
    def test_raise_(self):
        wb = Workbook()
        wb.new_sheet()
        wb.set_cell_contents("Sheet1", "A1", "200")
        wb.set_cell_contents("Sheet1", "A2", "2000")
        wb.set_cell_contents("Sheet1", "A3", "200")
        wb.set_cell_contents("Sheet1", "A4", "-2")
        with self.assertRaises(ValueError):
            wb.sort_region("Sheet1", "A1", "A4", [1, 2])
        wb.set_cell_contents("Sheet1", "B1", "200")
        wb.set_cell_contents("Sheet1", "B2", "2000")
        wb.set_cell_contents("Sheet1", "B33", "200")
        wb.set_cell_contents("Sheet1", "B4", "-2")
        with self.assertRaises(ValueError):
            wb.sort_region("Sheet1", "A1", "A4", [1, 2, 3])

        with self.assertRaises(ValueError):
            wb.sort_region("Sheet1", "A1", "A4", [2, 2])

        with self.assertRaises(ValueError):
            wb.sort_region("Sheet1", "A1", "A4", [-2, 2])

        with self.assertRaises(ValueError):
            wb.sort_region("Sheet1", "A1", "A4", [0, 2])

        with self.assertRaises(ValueError):
            wb.sort_region("Sheet1", "A1", "A4", [1, 1.5])

    def test_refernces_in_sorted_region_update(self):
        wb = Workbook()
        wb.new_sheet()
        wb.set_cell_contents("Sheet1", "A1", "200")
        wb.set_cell_contents("Sheet1", "A2", "2000")
        wb.set_cell_contents("Sheet1", "A3", "200")
        wb.set_cell_contents("Sheet1", "A4", "-2")
        wb.set_cell_contents("Sheet1", "B1", "=A1")
        wb.set_cell_contents("Sheet1", "B2", "=A2")
        wb.set_cell_contents("Sheet1", "B3", "=A3")
        wb.set_cell_contents("Sheet1", "B4", "=A4")
        assert wb.get_cell_value("Sheet1", "A1") == Decimal("-2")
        assert wb.get_cell_value("Sheet1", "A2") == Decimal("200")
        assert wb.get_cell_value("Sheet1", "A3") == Decimal("200")
        assert wb.get_cell_value("Sheet1", "A4") == Decimal("2000")
        assert wb.get_cell_value("Sheet1", "B1") == Decimal("-2")
        assert wb.get_cell_value("Sheet1", "B2") == Decimal("200")
        assert wb.get_cell_value("Sheet1", "B3") == Decimal("200")
        assert wb.get_cell_value("Sheet1", "B4") == Decimal("2000")
        

        
        



        

if __name__ == "__main__":
    unittest.main()