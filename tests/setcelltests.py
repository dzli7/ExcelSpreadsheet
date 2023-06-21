import context
from sheets.workbook import Workbook
from sheets.sheet import Sheet
from sheets.cellerror import CellError
from sheets.cellerrortype import CellErrorType

import unittest
from decimal import Decimal

# Test Suites for Sheets: Currently all test cases pass
# Can comment if desired, will be moved to seperate file eventually


class SetCellTests(unittest.TestCase):
    def test_set_cell_decimals(self):
        wb = Workbook()
        wb.new_sheet()
        wb.set_cell_contents("Sheet1", "A1", "'-.000123")
        wb.set_cell_contents("Sheet1", "A2", "'0.000123123004")
        wb.set_cell_contents("Sheet1", "B1", "=A2 * 679")
        assert (wb.get_cell_value("Sheet1", "A1") == Decimal('-.000123'))
        assert (wb.get_cell_value("Sheet1", "A2") == Decimal('0.000123123004'))
        assert (wb.get_cell_value("Sheet1", "B1") == Decimal('0.083600519716'))

        wb.set_cell_contents("Sheet1", "A3", "'105")
        wb.set_cell_contents("Sheet1", "A4", "=A3 / 14.2835")
        wb.set_cell_contents("Sheet1", "A5", "=A3 - 14.2835")
        wb.set_cell_contents("Sheet1", "A6", "=A3 + 14.2835")
        assert (wb.get_cell_value("Sheet1", "A4") ==
                Decimal('7.351139426611124724332271502'))
        assert (wb.get_cell_value("Sheet1", "A5") == Decimal('90.7165'))
        assert (wb.get_cell_value("Sheet1", "A6") == Decimal('119.2835'))

    def test_set_cell_contents_unary(self):
        wb = Workbook()
        wb.new_sheet()
        wb.set_cell_contents("Sheet1", "A1", "'1.2")
        wb.set_cell_contents("Sheet1", "A2", "=-A1")
        wb.set_cell_contents("Sheet1", "B1", "=-A2")
        assert (wb.get_cell_value("Sheet1", "A1") == Decimal('1.2'))
        assert (wb.get_cell_value("Sheet1", "A2") == Decimal('-1.2'))
        assert (wb.get_cell_value("Sheet1", "B1") == Decimal('1.2'))

        wb.set_cell_contents("Sheet1", "A1", "'3")
        assert (wb.get_cell_value("Sheet1", "A1") == Decimal('3'))
        assert (wb.get_cell_value("Sheet1", "A2") == Decimal('-3'))
        assert (wb.get_cell_value("Sheet1", "B1") == Decimal('3'))

    def test_set_cell_contents_addition(self):
        wb = Workbook()
        wb.new_sheet()
        wb.set_cell_contents("Sheet1", "A1", "'2")
        wb.set_cell_contents("Sheet1", "A2", "=A1+2")
        wb.set_cell_contents("Sheet1", "B1", "=A2+A1")
        assert (wb.get_cell_value("Sheet1", "A1") == Decimal('2'))
        assert (wb.get_cell_value("Sheet1", "A2") == Decimal('4'))
        assert (wb.get_cell_value("Sheet1", "B1") == Decimal('6'))

        wb.set_cell_contents("Sheet1", "A1", "'3")
        assert (wb.get_cell_value("Sheet1", "A1") == Decimal('3'))
        assert (wb.get_cell_value("Sheet1", "A2") == Decimal('5'))
        assert (wb.get_cell_value("Sheet1", "B1") == Decimal('8'))

    def test_set_cell_contents_subtraction(self):
        wb = Workbook()
        wb.new_sheet()
        wb.set_cell_contents("Sheet1", "A1", "'7")
        wb.set_cell_contents("Sheet1", "A3", "'3")
        wb.set_cell_contents("Sheet1", "A2", "=A1-A3")
        wb.set_cell_contents("Sheet1", "B1", "=A1-A2")
        assert (wb.get_cell_value("Sheet1", "A1") == Decimal('7'))
        assert (wb.get_cell_value("Sheet1", "A3") == Decimal('3'))
        assert (wb.get_cell_value("Sheet1", "A2") == Decimal('4'))
        assert (wb.get_cell_value("Sheet1", "B1") == Decimal('3'))

        wb.set_cell_contents("Sheet1", "A3", "'5")
        assert (wb.get_cell_value("Sheet1", "A1") == Decimal('7'))
        assert (wb.get_cell_value("Sheet1", "A3") == Decimal('5'))
        assert (wb.get_cell_value("Sheet1", "A2") == Decimal('2'))
        assert (wb.get_cell_value("Sheet1", "B1") == Decimal('5'))

    def test_set_cell_contents_multiplication(self):
        wb = Workbook()
        wb.new_sheet()
        wb.set_cell_contents("Sheet1", "A1", "'2")
        wb.set_cell_contents("Sheet1", "A3", "'3")
        wb.set_cell_contents("Sheet1", "A2", "=A1*A3")
        wb.set_cell_contents("Sheet1", "B1", "=A2*A3")
        assert (wb.get_cell_value("Sheet1", "A1") == Decimal('2'))
        assert (wb.get_cell_value("Sheet1", "A3") == Decimal('3'))
        assert (wb.get_cell_value("Sheet1", "A2") == Decimal('6'))
        assert (wb.get_cell_value("Sheet1", "B1") == Decimal('18'))

        wb.set_cell_contents("Sheet1", "A1", "'3")
        assert (wb.get_cell_value("Sheet1", "A1") == Decimal('3'))
        assert (wb.get_cell_value("Sheet1", "A3") == Decimal('3'))
        assert (wb.get_cell_value("Sheet1", "A2") == Decimal('9'))
        assert (wb.get_cell_value("Sheet1", "B1") == Decimal('27'))

    def test_set_cell_contents_division(self):
        wb = Workbook()
        wb.new_sheet()
        wb.set_cell_contents("Sheet1", "A1", "'36")
        wb.set_cell_contents("Sheet1", "A3", "'2")
        wb.set_cell_contents("Sheet1", "A2", "=A1/A3")
        wb.set_cell_contents("Sheet1", "B1", "=A2/A3")
        assert (wb.get_cell_value("Sheet1", "A1") == Decimal('36'))
        assert (wb.get_cell_value("Sheet1", "A3") == Decimal('2'))
        assert (wb.get_cell_value("Sheet1", "A2") == Decimal('18'))
        assert (wb.get_cell_value("Sheet1", "B1") == Decimal('9'))
        wb.set_cell_contents("Sheet1", "A3", "'3")
        assert (wb.get_cell_value("Sheet1", "A1") == Decimal('36'))
        assert (wb.get_cell_value("Sheet1", "A3") == Decimal('3'))
        assert (wb.get_cell_value("Sheet1", "A2") == Decimal('12'))
        assert (wb.get_cell_value("Sheet1", "B1") == Decimal('4'))

    def test_set_cell_contents_all_op(self):
        wb = Workbook()
        wb.new_sheet()
        wb.set_cell_contents("Sheet1", "A1", "'2")
        wb.set_cell_contents("Sheet1", "A2", "=A1*6/3+16/2")
        wb.set_cell_contents("Sheet1", "A3", "=-10+-2*A2/A1")
        wb.set_cell_contents("Sheet1", "A4", "=-A3*2-A1")
        assert (wb.get_cell_value("Sheet1", "A1") == Decimal('2'))
        assert (wb.get_cell_value("Sheet1", "A2") == Decimal('12'))
        assert (wb.get_cell_value("Sheet1", "A3") == Decimal('-22'))
        assert (wb.get_cell_value("Sheet1", "A4") == Decimal('42'))

        wb.set_cell_contents("Sheet1", "A1", "'4")
        assert (wb.get_cell_value("Sheet1", "A1") == Decimal('4'))
        assert (wb.get_cell_value("Sheet1", "A2") == Decimal('16'))
        assert (wb.get_cell_value("Sheet1", "A3") == Decimal('-18'))
        assert (wb.get_cell_value("Sheet1", "A4") == Decimal('32'))

    def test_set_cell_invalid_loc_add(self):
        wb = Workbook()
        wb.new_sheet()
        wb.set_cell_contents("Sheet1", "A1", "'2")
        wb.set_cell_contents("Sheet1", "A3", "=A1+A2")
        assert (wb.get_cell_value("Sheet1", "A2") == None)
        assert (wb.get_cell_value("Sheet1", "A3") == Decimal('2'))

    def test_set_cell_invalid_loc_concat(self):
        wb = Workbook()
        wb.new_sheet()
        wb.set_cell_contents("Sheet1", "A1", "'2")
        wb.set_cell_contents("Sheet1", "A3", "=A1&A2")
        assert (wb.get_cell_value("Sheet1", "A2") == None)
        assert (wb.get_cell_value("Sheet1", "A3") == "2")

    def test_concat_str_w_paren(self):
        wb = Workbook()
        wb.new_sheet()
        wb.set_cell_contents("Sheet1", "A1", "'1")
        wb.set_cell_contents("Sheet1", "A2", "'2")
        wb.set_cell_contents("Sheet1", "A3", "=A2&(A1&A2&A1)")
        assert (wb.get_cell_value("Sheet1", "A3") == "2121")

    def test_concat_str(self):
        wb = Workbook()
        wb.new_sheet()
        wb.set_cell_contents("Sheet1", "A1", "'hello")
        wb.set_cell_contents("Sheet1", "A2", "'sup bro")

        wb.set_cell_contents("Sheet1", "A3", '=a2&"   3"')
        assert (wb.get_cell_value("Sheet1", "a3") == "sup bro   3")

        wb.set_cell_contents("Sheet1", "A4", '=a2  +  3')
        value = wb.get_cell_value("Sheet1", "A4")
        assert (isinstance(value, CellError))
        assert (value.get_type() == CellErrorType.TYPE_ERROR)

        wb.set_cell_contents("Sheet1", "A5", '=a2  *  3')
        value = wb.get_cell_value("Sheet1", "A5")
        assert (isinstance(value, CellError))
        assert (value.get_type() == CellErrorType.TYPE_ERROR)

    def test_set_cell_equals_empty_string(self):
        wb = Workbook()
        wb.new_sheet()
        wb.set_cell_contents("Sheet1", "A1", '=""')
        assert (wb.get_cell_contents("Sheet1", "A1") == '=""')
        assert (wb.get_cell_value("Sheet1", "A1") == '')

    def test_update_cells_from_delete_sheet_action(self):
        wb = Workbook()
        wb.new_sheet()
        wb.new_sheet()
        wb.set_cell_contents("Sheet2", "A2", "2")
        wb.set_cell_contents("Sheet1", "A1", "=Sheet2!A2")
        assert (wb.get_sheet("Sheet1").get_cell("A1").get_parent_cells() != [])
        wb.del_sheet("Sheet2")
        assert (wb.get_cell_contents("Sheet1", "A1")
                == "=Sheet2!A2")
        value = wb.get_cell_value("Sheet1", "A1")
        assert (isinstance(value, CellError))
        assert (value.get_type() == CellErrorType.BAD_REFERENCE)

    def test_set_cell_circ_clear_cell(self):
        wb = Workbook()
        _, n1 = wb.new_sheet()
        wb.set_cell_contents(n1, "A1", "=A2")
        wb.set_cell_contents(n1, "A2", "=A1")
        value = wb.get_cell_value(n1, "A1")
        assert(isinstance(value, CellError))
        assert(value.get_type() == CellErrorType.CIRCULAR_REFERENCE)
        value = wb.get_cell_value(n1, "A2")
        assert(isinstance(value, CellError))
        assert(value.get_type() == CellErrorType.CIRCULAR_REFERENCE)
        
        wb.set_cell_contents(n1, "A1", "5")
        assert wb.get_cell_value(n1, "A2") == Decimal("5")
        wb.clear_cell(n1, "A1")
        assert(wb.get_cell_contents(n1, "A1") == None)
        assert(wb.get_cell_value(n1, "A1") == None)
        wb.clear_cell(n1, "A2")
        assert(wb.get_cell_contents(n1, "A2") == None)
        assert(wb.get_cell_value(n1, "A2") == None)

    def test_set_cell_to_sheet_name(self):
        wb = Workbook()
        _, n1 = wb.new_sheet()
        wb.set_cell_contents(n1, "A1", "=Sheet1 + Sheet2")
        value = wb.get_cell_value(n1, "A1")
        assert(isinstance(value, CellError))
        assert(value.get_type() == CellErrorType.BAD_REFERENCE)


if __name__ == "__main__":
    unittest.main()
