import context
from sheets.workbook import Workbook
from sheets.sheet import Sheet
from sheets.cellerror import CellError
from sheets.cellerrortype import CellErrorType

import unittest
from decimal import Decimal


class CellErrorTests(unittest.TestCase):

    def test_division_by_zero(self):
        wb = Workbook()
        wb.new_sheet()
        wb.set_cell_contents("Sheet1", "A1", "'36")
        wb.set_cell_contents("Sheet1", "A3", "'0")
        wb.set_cell_contents("Sheet1", "A2", "=A1/A3")
        wb.set_cell_contents("Sheet1", "B1", "=A2/A3")
        assert (wb.get_cell_value("Sheet1", "A1") == Decimal('36'))
        value = wb.get_cell_value("Sheet1", 'a2')
        assert (wb.get_cell_contents("Sheet1", 'a2') == "=A1/A3")
        assert (isinstance(value, CellError))
        assert (value.get_type() == CellErrorType.DIVIDE_BY_ZERO)

        value = wb.get_cell_value("Sheet1", 'b1')
        assert (wb.get_cell_contents("Sheet1", 'b1') == "=A2/A3")
        assert (isinstance(value, CellError))
        assert (value.get_type() == CellErrorType.DIVIDE_BY_ZERO)

        wb.set_cell_contents("Sheet1", "a4", "=(A1*4+(4/0)+9*2)/2")

        value = wb.get_cell_value("Sheet1", "a4")
        assert (wb.get_cell_contents("Sheet1", 'a4') == "=(A1*4+(4/0)+9*2)/2")
        assert (isinstance(value, CellError))
        assert (value.get_type() == CellErrorType.DIVIDE_BY_ZERO)

        wb.set_cell_contents("Sheet1", "d6", "= #DIV/0!")
        val2 = wb.get_cell_value("Sheet1", "d6")
        assert (wb.get_cell_contents("Sheet1", 'd6') == "= #DIV/0!")
        assert isinstance(val2, CellError)
        assert val2.get_type() == CellErrorType.DIVIDE_BY_ZERO

        wb.set_cell_contents("Sheet1", "d7", "= -d6")
        val3 = wb.get_cell_value("Sheet1", "d7")
        assert (wb.get_cell_contents("Sheet1", 'd7') == "= -d6")
        assert isinstance(val3, CellError)
        assert val3.get_type() == CellErrorType.DIVIDE_BY_ZERO

    def test_bad_reference(self):
        wb = Workbook()
        wb.new_sheet()
        wb.set_cell_contents("Sheet1", "d2", "1")
        wb.set_cell_contents("Sheet1", 'd3', '=nonexistent!b4')
        value = wb.get_cell_value("Sheet1", 'd3')
        assert (wb.get_cell_contents("Sheet1", 'd3') == "=nonexistent!b4")
        assert isinstance(value, CellError)
        assert value.get_type() == CellErrorType.BAD_REFERENCE

        wb.set_cell_contents("Sheet1", "d4", "=d3 + 5")
        val2 = wb.get_cell_value("Sheet1", "d4")
        assert (wb.get_cell_contents("Sheet1", 'd4') == "=d3 + 5")
        assert isinstance(val2, CellError)
        assert val2.get_type() == CellErrorType.BAD_REFERENCE

        wb.set_cell_contents("Sheet1", "d5", "=d2 + ZZZZZ99999 ")
        val2 = wb.get_cell_value("Sheet1", "d5")
        assert (wb.get_cell_contents("Sheet1", 'd5') == "=d2 + ZZZZZ99999 ")
        assert isinstance(val2, CellError)
        assert val2.get_type() == CellErrorType.BAD_REFERENCE

        wb.set_cell_contents("Sheet1", "d6", "= #REF!")
        val2 = wb.get_cell_value("Sheet1", "d6")
        assert (wb.get_cell_contents("Sheet1", 'd6') == "= #REF!")
        assert isinstance(val2, CellError)
        assert val2.get_type() == CellErrorType.BAD_REFERENCE

        wb.set_cell_contents("Sheet1", "d7", "= -d6")
        val3 = wb.get_cell_value("Sheet1", "d7")
        assert (wb.get_cell_contents("Sheet1", 'd7') == "= -d6")
        assert isinstance(val3, CellError)
        assert val3.get_type() == CellErrorType.BAD_REFERENCE

    def test_circular_reference_2(self):
        wb = Workbook()
        wb.new_sheet()
        wb.set_cell_contents("Sheet1", "a1", "'2")
        wb.set_cell_contents("Sheet1", "a2", "=a1*2")
        wb.set_cell_contents("Sheet1", "a1", "=a2*2")

        v1 = wb.get_cell_value("Sheet1", "a1")
        v2 = wb.get_cell_value("Sheet1", "a2")
        assert (wb.get_cell_contents("Sheet1", "a1") == "=a2*2")
        assert (wb.get_cell_contents("Sheet1", "a2") == "=a1*2")

        assert isinstance(v1, CellError)
        assert v1.get_type() == CellErrorType.CIRCULAR_REFERENCE
        assert isinstance(v2, CellError)
        assert v2.get_type() == CellErrorType.CIRCULAR_REFERENCE

        wb.set_cell_contents("Sheet1", "d7", "= -a1")
        val3 = wb.get_cell_value("Sheet1", "d7")
        assert (wb.get_cell_contents("Sheet1", 'd7') == "= -a1")
        assert isinstance(val3, CellError)
        assert val3.get_type() == CellErrorType.CIRCULAR_REFERENCE

    def test_circular_reference_4(self):
        wb = Workbook()
        wb.new_sheet()
        wb.set_cell_contents("Sheet1", "a1", "'2")
        wb.set_cell_contents("Sheet1", "a2", "=a1*2")
        wb.set_cell_contents("Sheet1", "a3", "=a2*2")
        wb.set_cell_contents("Sheet1", "a4", "=a3*2")
        wb.set_cell_contents("Sheet1", "a1", "=a4*2")

        v1 = wb.get_cell_value("Sheet1", "a1")
        v2 = wb.get_cell_value("Sheet1", "a2")
        v3 = wb.get_cell_value("Sheet1", "a3")
        v4 = wb.get_cell_value("Sheet1", "a4")

        assert (wb.get_cell_contents("Sheet1", "a1") == "=a4*2")
        assert (wb.get_cell_contents("Sheet1", "a2") == "=a1*2")
        assert (wb.get_cell_contents("Sheet1", "a3") == "=a2*2")
        assert (wb.get_cell_contents("Sheet1", "a4") == "=a3*2")

        assert isinstance(v1, CellError)
        assert v1.get_type() == CellErrorType.CIRCULAR_REFERENCE
        assert isinstance(v2, CellError)
        assert v2.get_type() == CellErrorType.CIRCULAR_REFERENCE
        assert isinstance(v3, CellError)
        assert v3.get_type() == CellErrorType.CIRCULAR_REFERENCE
        assert isinstance(v4, CellError)
        assert v4.get_type() == CellErrorType.CIRCULAR_REFERENCE

    def test_parse_error(self):
        wb = Workbook()
        wb.new_sheet()

        wb.set_cell_contents("Sheet1", "a1", "=aeiou")
        value = wb.get_cell_value("Sheet1", "a1")
        assert (wb.get_cell_contents("Sheet1", "a1") == "=aeiou")
        assert isinstance(value, CellError)
        assert value.get_type() == CellErrorType.PARSE_ERROR

        wb.set_cell_contents("Sheet1", "a2", "=a1a1a1a1a1a1a1")
        value = wb.get_cell_value("Sheet1", "a2")
        assert (wb.get_cell_contents("Sheet1", "a2") == "=a1a1a1a1a1a1a1")
        assert isinstance(value, CellError)
        assert value.get_type() == CellErrorType.PARSE_ERROR

        wb.set_cell_contents("Sheet1", "a3", "=-+")
        value = wb.get_cell_value("Sheet1", "a3")
        assert (wb.get_cell_contents("Sheet1", "a3") == "=-+")
        assert isinstance(value, CellError)
        assert value.get_type() == CellErrorType.PARSE_ERROR

        wb.set_cell_contents("Sheet1", "a4", "=()")
        value = wb.get_cell_value("Sheet1", "a4")
        assert (wb.get_cell_contents("Sheet1", "a4") == "=()")
        assert isinstance(value, CellError)
        assert value.get_type() == CellErrorType.PARSE_ERROR

        wb.set_cell_contents("Sheet1", "a5", "=29+3*7+@$")
        value = wb.get_cell_value("Sheet1", "a5")
        assert (wb.get_cell_contents("Sheet1", "a5") == "=29+3*7+@$")
        assert isinstance(value, CellError)
        assert value.get_type() == CellErrorType.PARSE_ERROR

        wb.set_cell_contents(
            "Sheet1", "a6", "=@(#%&@(#&%$)@#(!@#(*&&%/.,/.,/][")
        value = wb.get_cell_value("Sheet1", "a6")
        assert (wb.get_cell_contents("Sheet1", "a6")
                == "=@(#%&@(#&%$)@#(!@#(*&&%/.,/.,/][")
        assert isinstance(value, CellError)
        assert value.get_type() == CellErrorType.PARSE_ERROR

    def test_type_error(self):
        wb = Workbook()
        wb.new_sheet()

        wb.set_cell_contents("Sheet1", "a1", "'I am a string")
        wb.set_cell_contents("Sheet1", "a2", "=a1+2")

        value = wb.get_cell_value("Sheet1", "a2")
        assert (wb.get_cell_contents("Sheet1", "a2") == "=a1+2")
        assert isinstance(value, CellError)
        assert value.get_type() == CellErrorType.TYPE_ERROR

        wb.set_cell_contents("Sheet1", "a3", "'I am a string")
        wb.set_cell_contents("Sheet1", "a4", "'Another string")
        wb.set_cell_contents("Sheet1", "a5", "=a3+a4")

        value = wb.get_cell_value("Sheet1", "a5")
        assert (wb.get_cell_contents("Sheet1", "a5") == "=a3+a4")
        assert isinstance(value, CellError)
        assert value.get_type() == CellErrorType.TYPE_ERROR

        wb.set_cell_contents("Sheet1", "a6", "=a3-a4*a1/5+6*7")
        value = wb.get_cell_value("Sheet1", "a6")
        assert (wb.get_cell_contents("Sheet1", "a6") == "=a3-a4*a1/5+6*7")
        assert isinstance(value, CellError)
        assert value.get_type() == CellErrorType.TYPE_ERROR

    def test_multiple_circular_dependencies(self):
        wb = Workbook()
        _, n1 = wb.new_sheet()
        wb.set_cell_contents(n1, "A1", "=A2+B1")
        wb.set_cell_contents(n1, "A2", "=A3")
        wb.set_cell_contents(n1, "A3", "=A1")
        wb.set_cell_contents(n1, "B1", "2")
        assert (wb.get_cell_value(n1, "b1") == Decimal("2"))
        value = wb.get_cell_value(n1, "A1")
        assert isinstance(value, CellError)
        assert value.get_type() == CellErrorType.CIRCULAR_REFERENCE

    def test_setting_error(self):
        wb = Workbook()
        _, n1 = wb.new_sheet()
        wb.set_cell_contents(n1, "A1", "#DIV/0!")
        v = wb.get_cell_value(n1, "A1")
        assert isinstance(v, CellError)
        assert v.get_type() == CellErrorType.DIVIDE_BY_ZERO

    def test_circular_reference_outside_cell(self):
        wb = Workbook()
        _, n1 = wb.new_sheet()
        wb.set_cell_contents(n1, "A1", "=A2 + B1")
        wb.set_cell_contents(n1, "A2", "=A3")
        wb.set_cell_contents(n1, "A3", "=A1 + Other!A1")
        wb.set_cell_contents(n1, "B1", "=B2")
        wb.set_cell_contents(n1, "B2", "=B3")
        wb.set_cell_contents(n1, "B3", "=B1")
        
        v1 = wb.get_cell_value(n1, "A1")
        assert isinstance(v1, CellError)
        assert v1.get_type() == CellErrorType.CIRCULAR_REFERENCE

        v2 = wb.get_cell_value(n1, "A2")
        assert isinstance(v2, CellError)
        assert v2.get_type() == CellErrorType.CIRCULAR_REFERENCE

        v3 = wb.get_cell_value(n1, "A3")
        assert isinstance(v3, CellError)
        assert v3.get_type() == CellErrorType.CIRCULAR_REFERENCE

        v4 = wb.get_cell_value(n1, "B1")
        assert isinstance(v4, CellError)
        assert v4.get_type() == CellErrorType.CIRCULAR_REFERENCE

        v5 = wb.get_cell_value(n1, "B2")
        assert isinstance(v5, CellError)
        assert v5.get_type() == CellErrorType.CIRCULAR_REFERENCE

        v6 = wb.get_cell_value(n1, "B3")
        assert isinstance(v6, CellError)
        assert v6.get_type() == CellErrorType.CIRCULAR_REFERENCE

if __name__ == "__main__":
    unittest.main()
