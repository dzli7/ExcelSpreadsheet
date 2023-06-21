import context
from sheets.workbook import Workbook
from sheets.sheet import Sheet
from sheets.cellerror import CellError
from sheets.cellerrortype import CellErrorType
import json
import io

import unittest
from decimal import Decimal

# Test Suite for Comparisons

def correct_error(error, actual_error):
    # Check for specific error type matches
    if isinstance(error, CellError):
        if error.get_type() == actual_error:
            return True
    return False

class ComparisonTests(unittest.TestCase):

    def test_comparison_low_precedence(self):
        wb = Workbook()
        _, n1 = wb.new_sheet()
        wb.set_cell_contents(n1, "A1", "=1+2 < 1+4")
        wb.set_cell_contents(n1, "A2", "=1+2 <= 1+4")
        wb.set_cell_contents(n1, "A3", "=1+2 > 1+4")
        wb.set_cell_contents(n1, "A4", "=1+2 >= 1+4")
        assert wb.get_cell_value(n1, "A1") == True
        assert wb.get_cell_value(n1, "A2") == True
        assert wb.get_cell_value(n1, "A3") == False
        assert wb.get_cell_value(n1, "A4") == False

    def test_less_than(self):
        wb = Workbook()
        wb.new_sheet()
        wb.set_cell_contents("Sheet1", "A1", "5")
        wb.set_cell_contents("Sheet1", "A2", "8")
        wb.set_cell_contents("Sheet1", "A3", "= A1 < A2")
        assert wb.get_cell_value("Sheet1", "A3")
        wb.set_cell_contents("Sheet1", "A6", "= 5 < 8")
        wb.set_cell_contents("Sheet1", "A7", "= 5555 < 8")
        assert wb.get_cell_value("Sheet1", "A6")
        assert not wb.get_cell_value("Sheet1", "A7")

    def test_greater_than(self):
        wb = Workbook()
        wb.new_sheet()
        wb.set_cell_contents("Sheet1", "A1", "5")
        wb.set_cell_contents("Sheet1", "A2", "8")
        wb.set_cell_contents("Sheet1", "A3", "= A1 > A2")
        assert not wb.get_cell_value("Sheet1", "A3")
        wb.set_cell_contents("Sheet1", "A6", "= 5 > 8")
        wb.set_cell_contents("Sheet1", "A7", "= 5555 > 8")
        assert not wb.get_cell_value("Sheet1", "A6")
        assert wb.get_cell_value("Sheet1", "A7")

    def test_less_than_equal_to(self):
        wb = Workbook()
        wb.new_sheet()
        wb.set_cell_contents("Sheet1", "A1", "5")
        wb.set_cell_contents("Sheet1", "A2", "8")
        wb.set_cell_contents("Sheet1", "A3", "= A1 <= A2")
        assert wb.get_cell_value("Sheet1", "A3")
        wb.set_cell_contents("Sheet1", "A6", "= 5 <= 8")
        wb.set_cell_contents("Sheet1", "A7", "= 5555 <= 8")
        assert wb.get_cell_value("Sheet1", "A6")
        assert not wb.get_cell_value("Sheet1", "A7")

    def test_greater_than_equal_to(self):
        wb = Workbook()
        wb.new_sheet()
        wb.set_cell_contents("Sheet1", "A1", "5")
        wb.set_cell_contents("Sheet1", "A2", "8")
        wb.set_cell_contents("Sheet1", "A3", "= A1 >= A2")
        assert not wb.get_cell_value("Sheet1", "A3")
        wb.set_cell_contents("Sheet1", "A6", "= 5 >= 8")
        wb.set_cell_contents("Sheet1", "A7", "= 5555 >= 8")
        assert not wb.get_cell_value("Sheet1", "A6")
        assert wb.get_cell_value("Sheet1", "A7")

    def test_equal_to_same_type(self):
        wb = Workbook()
        wb.new_sheet()
        wb.set_cell_contents("Sheet1", "A1", "5")
        wb.set_cell_contents("Sheet1", "A2", "8")
        wb.set_cell_contents("Sheet1", "A3", "5")
        wb.set_cell_contents("Sheet1", "A4", "= A1 = A2")
        wb.set_cell_contents("Sheet1", "A5", "= A1 == A2")
        wb.set_cell_contents("Sheet1", "A6", "= A1 = A3")
        wb.set_cell_contents("Sheet1", "A7", "= A1 == A3")
        assert not wb.get_cell_value("Sheet1", "A4")
        assert not wb.get_cell_value("Sheet1", "A5")
        assert wb.get_cell_value("Sheet1", "A6")
        assert wb.get_cell_value("Sheet1", "A7")

        wb.set_cell_contents("Sheet1", "B1", "= 5 = 8")
        wb.set_cell_contents("Sheet1", "B2", "= 5 == 8")
        wb.set_cell_contents("Sheet1", "B3", "= 5 = 5")
        wb.set_cell_contents("Sheet1", "B4", "= 5 == 5")
        assert not wb.get_cell_value("Sheet1", "B1")
        assert not wb.get_cell_value("Sheet1", "B2")
        assert wb.get_cell_value("Sheet1", "B3")
        assert wb.get_cell_value("Sheet1", "B4")

        wb.set_cell_contents("Sheet1", "C1", "= A1 = 8")
        wb.set_cell_contents("Sheet1", "C2", "= A1 == 8")
        wb.set_cell_contents("Sheet1", "C3", "= A1 = 5")
        wb.set_cell_contents("Sheet1", "C4", "= A1 == 5")
        assert not wb.get_cell_value("Sheet1", "C1")
        assert not wb.get_cell_value("Sheet1", "C2")
        assert wb.get_cell_value("Sheet1", "C3")
        assert wb.get_cell_value("Sheet1", "C4")

    def test_comparison_bool_str_num(self):
        wb = Workbook()
        _, name = wb.new_sheet()
        wb.set_cell_contents(name, "a2", '="12" > 12')
        assert wb.get_cell_value(name, "a2") == True
        wb.set_cell_contents(name, "a2", '="TRUE" > FALSE')
        assert wb.get_cell_value(name, "a2") == False
        wb.set_cell_contents(name, "a2", '=FALSE < TRUE')
        assert wb.get_cell_value(name, "a2") == True
        wb.set_cell_contents(name, "a2", '="a" < "["')
        assert wb.get_cell_value(name, "a2") == False
        wb.set_cell_contents(name, "a2", '=A3 = 0')
        assert wb.get_cell_value(name, "a2") == True

    def test_comparison_empty_operand_num(self):
        wb = Workbook()
        wb.new_sheet()
        wb.set_cell_contents("Sheet1", "A1", "5")
        wb.set_cell_contents("Sheet1", "A4", "= A1 = B4")
        wb.set_cell_contents("Sheet1", "A5", "= A1 == B4")
        wb.set_cell_contents("Sheet1", "A6", "= A1 <= B4")
        wb.set_cell_contents("Sheet1", "A7", "= A1 >= B4")
        wb.set_cell_contents("Sheet1", "A8", "= A1 < B4")
        wb.set_cell_contents("Sheet1", "A9", "= A1 > B4")
        wb.set_cell_contents("Sheet1", "A10", "= A1 <> B4")
        wb.set_cell_contents("Sheet1", "A11", "= A1 != B4")
        assert not wb.get_cell_value("Sheet1", "A4")
        assert not wb.get_cell_value("Sheet1", "A5")
        assert not wb.get_cell_value("Sheet1", "A6")
        assert wb.get_cell_value("Sheet1", "A7")
        assert not wb.get_cell_value("Sheet1", "A8")
        assert wb.get_cell_value("Sheet1", "A9")
        assert wb.get_cell_value("Sheet1", "A10")
        assert wb.get_cell_value("Sheet1", "A11")

        wb.set_cell_contents("Sheet1", "A1", "0")
        assert wb.get_cell_value("Sheet1", "A4")
        assert wb.get_cell_value("Sheet1", "A5")
        assert wb.get_cell_value("Sheet1", "A6")
        assert wb.get_cell_value("Sheet1", "A7")
        assert not wb.get_cell_value("Sheet1", "A8")
        assert not wb.get_cell_value("Sheet1", "A9")
        assert not wb.get_cell_value("Sheet1", "A10")
        assert not wb.get_cell_value("Sheet1", "A11")

        

    def test_comparison_empty_operand_string(self):
        wb = Workbook()
        wb.new_sheet()
        wb.set_cell_contents("Sheet1", "A1", "'HellO")
        wb.set_cell_contents("Sheet1", "A4", "= A1 = B4")
        wb.set_cell_contents("Sheet1", "A5", "= A1 == B4")
        wb.set_cell_contents("Sheet1", "A6", "= A1 <= B4")
        wb.set_cell_contents("Sheet1", "A7", "= A1 >= B4")
        wb.set_cell_contents("Sheet1", "A8", "= A1 < B4")
        wb.set_cell_contents("Sheet1", "A9", "= A1 > B4")
        wb.set_cell_contents("Sheet1", "A10", "= A1 <> B4")
        wb.set_cell_contents("Sheet1", "A11", "= A1 != B4")
        assert not wb.get_cell_value("Sheet1", "A4")
        assert not wb.get_cell_value("Sheet1", "A5")
        assert not wb.get_cell_value("Sheet1", "A6")
        assert wb.get_cell_value("Sheet1", "A7")
        assert not wb.get_cell_value("Sheet1", "A8")
        assert wb.get_cell_value("Sheet1", "A9")
        assert wb.get_cell_value("Sheet1", "A10")
        assert wb.get_cell_value("Sheet1", "A11")

        wb.set_cell_contents("Sheet1", "A1", "")
        assert wb.get_cell_value("Sheet1", "A4")
        assert wb.get_cell_value("Sheet1", "A5")
        assert wb.get_cell_value("Sheet1", "A6")
        assert wb.get_cell_value("Sheet1", "A7")
        assert not wb.get_cell_value("Sheet1", "A8")
        assert not wb.get_cell_value("Sheet1", "A9")
        assert not wb.get_cell_value("Sheet1", "A10")
        assert not wb.get_cell_value("Sheet1", "A11")

    def test_comparison_empty_operand_bool(self):
        wb = Workbook()
        wb.new_sheet()
        wb.set_cell_contents("Sheet1", "A1", "5")
        wb.set_cell_contents("Sheet1", "A4", "= A1 = B4")
        wb.set_cell_contents("Sheet1", "A5", "= A1 == B4")
        wb.set_cell_contents("Sheet1", "A6", "= A1 <= B4")
        wb.set_cell_contents("Sheet1", "A7", "= A1 >= B4")
        wb.set_cell_contents("Sheet1", "A8", "= A1 < B4")
        wb.set_cell_contents("Sheet1", "A9", "= A1 > B4")
        wb.set_cell_contents("Sheet1", "A10", "= A1 <> B4")
        wb.set_cell_contents("Sheet1", "A11", "= A1 != B4")
        assert not wb.get_cell_value("Sheet1", "A4")
        assert not wb.get_cell_value("Sheet1", "A5")
        assert not wb.get_cell_value("Sheet1", "A6")
        assert wb.get_cell_value("Sheet1", "A7")
        assert not wb.get_cell_value("Sheet1", "A8")
        assert wb.get_cell_value("Sheet1", "A9")
        assert wb.get_cell_value("Sheet1", "A10")
        assert wb.get_cell_value("Sheet1", "A11")

        wb.set_cell_contents("Sheet1", "A1", "False")
        assert wb.get_cell_value("Sheet1", "A4")
        assert wb.get_cell_value("Sheet1", "A5")
        assert wb.get_cell_value("Sheet1", "A6")
        assert wb.get_cell_value("Sheet1", "A7")
        assert not wb.get_cell_value("Sheet1", "A8")
        assert not wb.get_cell_value("Sheet1", "A9")
        assert not wb.get_cell_value("Sheet1", "A10")
        assert not wb.get_cell_value("Sheet1", "A11")

    def test_comparison_str_case_insensitive(self):
        wb = Workbook()
        wb.new_sheet()
        wb.set_cell_contents("Sheet1", "A1", "'HeLlO")
        wb.set_cell_contents("Sheet1", "A2", "hELLo")
        wb.set_cell_contents("Sheet1", "A4", "= A1 = A2")
        wb.set_cell_contents("Sheet1", "A5", "= A1 == A2")
        wb.set_cell_contents("Sheet1", "A6", "= A1 <= A2")
        wb.set_cell_contents("Sheet1", "A7", "= A1 >= A2")
        wb.set_cell_contents("Sheet1", "A8", "= A1 < A2")
        wb.set_cell_contents("Sheet1", "A9", "= A1 > A2")
        wb.set_cell_contents("Sheet1", "A10", "= A1 <> A2")
        wb.set_cell_contents("Sheet1", "A11", "= A1 != A2")
        assert wb.get_cell_value("Sheet1", "A4")
        assert wb.get_cell_value("Sheet1", "A5")
        assert wb.get_cell_value("Sheet1", "A6")
        assert wb.get_cell_value("Sheet1", "A7")
        assert not wb.get_cell_value("Sheet1", "A8")
        assert not wb.get_cell_value("Sheet1", "A9")
        assert not wb.get_cell_value("Sheet1", "A10")
        assert not wb.get_cell_value("Sheet1", "A11")

    

    

if __name__ == "__main__":
    unittest.main()

