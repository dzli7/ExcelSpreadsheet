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

class FunctionsTests(unittest.TestCase):

    def test_AND(self):
        wb = Workbook()
        wb.new_sheet()
        # Incorrect number of arguments should yield type error
        wb.set_cell_contents("Sheet1", "B1", "=AND()")
        assert correct_error(wb.get_cell_value("Sheet1", "B1"), CellErrorType.TYPE_ERROR)

        wb.set_cell_contents("Sheet1", "A1", "=true")
        wb.set_cell_contents("Sheet1", "A2", "=false")
        wb.set_cell_contents("Sheet1", "A3", "=AND(A1, A2)")
        assert not wb.get_cell_value("Sheet1", "A3")
        
        wb.set_cell_contents("Sheet1", "A4", "=AND(True, True)")
        assert wb.get_cell_value("Sheet1", "A4")
        wb.set_cell_contents("Sheet1", "A41", "=AND(True, False)")
        assert not wb.get_cell_value("Sheet1", "A41")
        wb.set_cell_contents("Sheet1", "A42", "=and(False, False)")
        assert not wb.get_cell_value("Sheet1", "A42")

        wb.set_cell_contents("Sheet1", "A5", "=false")
        wb.set_cell_contents("Sheet1", "A6", "=ANd(A5, A4)")
        assert not wb.get_cell_value("Sheet1", "A6")
        wb.set_cell_contents("Sheet1", "A10", "=aNd(true, TRUE, tRue, TruE, True, True)")
        assert wb.get_cell_value("Sheet1", "A10")
        wb.set_cell_contents("Sheet1", "A11", "=AnD(True, True, True, True, True, False)")
        assert not wb.get_cell_value("Sheet1", "A11")

        # all operands must be evaluated for AND and OR; function will not short-circuit
        wb.set_cell_contents("Sheet1", "A12", '=AND(5, "yes")')
        assert correct_error(wb.get_cell_value("Sheet1", "A12"), CellErrorType.TYPE_ERROR)

        wb.set_cell_contents("Sheet1", "A12", '=AND(5, "true")')
        assert wb.get_cell_value("Sheet1", "A12")

        wb.set_cell_contents("Sheet1", "A12", '=AND(0, "true")')
        assert not wb.get_cell_value("Sheet1", "A12")

    def test_OR(self):
        wb = Workbook()
        wb.new_sheet()
        # Incorrect number of arguments should yield type error
        wb.set_cell_contents("Sheet1", "B1", "=OR()")
        assert correct_error(wb.get_cell_value("Sheet1", "B1"), CellErrorType.TYPE_ERROR)

        wb.set_cell_contents("Sheet1", "A1", "=true")
        wb.set_cell_contents("Sheet1", "A2", "=false")
        wb.set_cell_contents("Sheet1", "A3", "=OR(A1, A2)")
        assert wb.get_cell_value("Sheet1", "A3")

        wb.set_cell_contents("Sheet1", "A4", "=OR(True, True)")
        assert wb.get_cell_value("Sheet1", "A4")
        wb.set_cell_contents("Sheet1", "A41", "=OR(True, False)")
        assert wb.get_cell_value("Sheet1", "A41")
        wb.set_cell_contents("Sheet1", "A42", "=OR(False, False)")
        assert not wb.get_cell_value("Sheet1", "A42")
        
        wb.set_cell_contents("Sheet1", "A5", "=false")
        wb.set_cell_contents("Sheet1", "A6", "=OR(A5, A4)")
        assert wb.get_cell_value("Sheet1", "A6")

        wb.set_cell_contents("Sheet1", "A10", "=or(true, TRUE, tRue, TruE, True, True)")
        assert wb.get_cell_value("Sheet1", "A10")
        wb.set_cell_contents("Sheet1", "A11", "=Or(True, True, True, True, True, False)")
        assert  wb.get_cell_value("Sheet1", "A11")
        wb.set_cell_contents("Sheet1", "A11", "=oR(FALSE, False, FalsE, false, False, FalSE)")
        assert not wb.get_cell_value("Sheet1", "A11")

        # all operands must be evaluated for AND and OR; function will not short-circuit
        wb.set_cell_contents("Sheet1", "A12", '=OR(5, "yes")')
        assert correct_error(wb.get_cell_value("Sheet1", "A12"), CellErrorType.TYPE_ERROR)

        wb.set_cell_contents("Sheet1", "A12", '=OR(5, "false")')
        assert wb.get_cell_value("Sheet1", "A12")

        wb.set_cell_contents("Sheet1", "A12", '=OR(0, "false")')
        assert not wb.get_cell_value("Sheet1", "A12")

    def test_XOR(self):
        wb = Workbook()
        wb.new_sheet()
        # Incorrect number of arguments should yield type error
        wb.set_cell_contents("Sheet1", "B1", "=XOR()")
        assert correct_error(wb.get_cell_value("Sheet1", "B1"), CellErrorType.TYPE_ERROR)

        wb.set_cell_contents("Sheet1", "A1", "=true")
        wb.set_cell_contents("Sheet1", "A2", "=false")
        wb.set_cell_contents("Sheet1", "A3", "=XOR(A1, A2)")
        assert wb.get_cell_value("Sheet1", "A3")
        wb.set_cell_contents("Sheet1", "A2", "=TRUE")
        assert not wb.get_cell_value("Sheet1", "A3")

        wb.set_cell_contents("Sheet1", "A4", "=xOr(True, True)")
        assert not wb.get_cell_value("Sheet1", "A4")
        wb.set_cell_contents("Sheet1", "A41", "=XOR(True, False)")
        assert wb.get_cell_value("Sheet1", "A41")
        wb.set_cell_contents("Sheet1", "A42", "=XOR(False, False)")
        assert not wb.get_cell_value("Sheet1", "A42")

        wb.set_cell_contents("Sheet1", "A10", "=XOr(true, TRUE, tRue, TruE, True, True)")
        assert not wb.get_cell_value("Sheet1", "A10")
        wb.set_cell_contents("Sheet1", "A11", "=xOR(True, False, True, False, True, False)")
        assert  wb.get_cell_value("Sheet1", "A11")
        wb.set_cell_contents("Sheet1", "A11", "=xor(FALSE, False, FalsE, false, False, FalSE)")
        assert not wb.get_cell_value("Sheet1", "A11")

        wb.set_cell_contents("Sheet1", "A12", '=XOR(5, "yes")')
        assert correct_error(wb.get_cell_value("Sheet1", "A12"), CellErrorType.TYPE_ERROR)

        wb.set_cell_contents("Sheet1", "A12", '=XOR(5, "false")')
        assert wb.get_cell_value("Sheet1", "A12")

        wb.set_cell_contents("Sheet1", "A12", '=XOR(0, "true")')
        assert wb.get_cell_value("Sheet1", "A12")


    def test_NOT(self):
        wb = Workbook()
        wb.new_sheet()
        # Incorrect number of arguments should yield type error
        wb.set_cell_contents("Sheet1", "B1", "=NOT()")
        assert correct_error(wb.get_cell_value("Sheet1", "B1"), CellErrorType.TYPE_ERROR)

        wb.set_cell_contents("Sheet1", "A1", "=true")
        wb.set_cell_contents("Sheet1", "A2", "=false")

        wb.set_cell_contents("Sheet1", "B2", "=NOT(A1, A2)")
        assert correct_error(wb.get_cell_value("Sheet1", "B2"), CellErrorType.TYPE_ERROR)

        wb.set_cell_contents("Sheet1", "C1", "=nOT(A1)")
        wb.set_cell_contents("Sheet1", "C2", "=not(A2)")
        assert not wb.get_cell_value("Sheet1", "C1")
        assert wb.get_cell_value("Sheet1", "C2")

        wb.set_cell_contents("Sheet1", "A12", '=NOT("yes")')
        assert correct_error(wb.get_cell_value("Sheet1", "A12"), CellErrorType.TYPE_ERROR)

        wb.set_cell_contents("Sheet1", "A12", '=NOT(5)')
        assert not wb.get_cell_value("Sheet1", "A12")

        wb.set_cell_contents("Sheet1", "A12", '=NOT(0)')
        assert wb.get_cell_value("Sheet1", "A12")

        wb.set_cell_contents("Sheet1", "A12", '=NOT("true")')
        assert not wb.get_cell_value("Sheet1", "A12")
        wb.set_cell_contents("Sheet1", "A12", '=NOT("false")')
        assert wb.get_cell_value("Sheet1", "A12")

    def test_EXACT(self):
        wb = Workbook()
        wb.new_sheet()
        wb.set_cell_contents("Sheet1", "A1", "hello")
        wb.set_cell_contents("Sheet1", "A2", "hello")
        
        wb.set_cell_contents("Sheet1", "A3", "=EXACT()")
        assert correct_error(wb.get_cell_value("Sheet1", "A3"), CellErrorType.TYPE_ERROR)
        wb.set_cell_contents("Sheet1", "A3", "=EXACT(A1)")
        assert correct_error(wb.get_cell_value("Sheet1", "A3"), CellErrorType.TYPE_ERROR)
        
        wb.set_cell_contents("Sheet1", "A3", "=ExaCT(A1, A2)")
        assert wb.get_cell_contents("Sheet1", "A3")

        wb.set_cell_contents("Sheet1", "A2", "HELLO")
        assert not wb.get_cell_value("Sheet1", "A3")

        wb.set_cell_contents("Sheet1", "A4", "slatt")
        wb.set_cell_contents("Sheet1", "A3", "=EXACT(A1, A2, A4)")
        assert correct_error(wb.get_cell_value("Sheet1", "A3"), CellErrorType.TYPE_ERROR)

        wb.set_cell_contents("Sheet1", "A40", "=exact(True, True)")
        assert wb.get_cell_value("Sheet1", "A40")

        wb.set_cell_contents("Sheet1", "A40", "=eXact(True, False)")
        assert not wb.get_cell_value("Sheet1", "A40")

        wb.set_cell_contents("Sheet1", "A40", "=EXAct(12, 12)")
        assert wb.get_cell_value("Sheet1", "A40")

        wb.set_cell_contents("Sheet1", "A40", "=EXACT(12, 122332)")
        assert not wb.get_cell_value("Sheet1", "A40")

        wb.set_cell_contents("Sheet1", "A40", "=EXACT('slatt', 'slatt')")
        assert wb.get_cell_value("Sheet1", "A40")

        wb.set_cell_contents("Sheet1", "A40", '=exact(True, "TRUE")')
        assert wb.get_cell_value("Sheet1", "A40")

        wb.set_cell_contents("Sheet1", "A40", '=exact(False, "FALSE")')
        assert wb.get_cell_value("Sheet1", "A40")

        wb.set_cell_contents("Sheet1", "A40", '=exact(False, "false")')
        assert not wb.get_cell_value("Sheet1", "A40")

        wb.set_cell_contents("Sheet1", "A40", '=exact(1, "1")')
        assert wb.get_cell_value("Sheet1", "A40")

        wb.set_cell_contents("Sheet1", "A40", '=exact(A99, "0")')
        assert not wb.get_cell_value("Sheet1", "A40")

    def test_IF(self):
        wb = Workbook()
        wb.new_sheet()
        wb.set_cell_contents("Sheet1", "A1", "=True")
        wb.set_cell_contents("Sheet1", "A2", "we lit")
        wb.set_cell_contents("Sheet1", "A3", "we not lit")
        wb.set_cell_contents("Sheet1", "B1", "toolie")
        wb.set_cell_contents("Sheet1", "A4", "=IF()")
        assert correct_error(wb.get_cell_value("Sheet1", "A4"), CellErrorType.TYPE_ERROR)

        wb.set_cell_contents("Sheet1", "A4", "=IF(A1)")
        assert correct_error(wb.get_cell_value("Sheet1", "A4"), CellErrorType.TYPE_ERROR)

        wb.set_cell_contents("Sheet1", "A4", "=IF(A1, A2, A3, B1)")
        assert correct_error(wb.get_cell_value("Sheet1", "A4"), CellErrorType.TYPE_ERROR)

        wb.set_cell_contents("Sheet1", "A4", "=IF(A1, A2, A3)")
        assert wb.get_cell_value("Sheet1", "A4") == "we lit"

        wb.set_cell_contents("Sheet1", "A1", "=False")
        assert wb.get_cell_value("Sheet1", "A4") == "we not lit"

        wb.set_cell_contents("Sheet1", "A4", "=IF(A1, A2)")
        assert wb.get_cell_value("Sheet1", "A4") == False

        wb.set_cell_contents('Sheet1', 'B1', '=1/0')
        wb.set_cell_contents('Sheet1', 'C1', '=IF(B1, A2, A3)')
        assert correct_error(wb.get_cell_value('Sheet1', 'C1'), CellErrorType.TYPE_ERROR)

    def test_IFERROR(self):
        wb = Workbook()
        wb.new_sheet()
        wb.set_cell_contents("Sheet1", "A1", "=90/10")
        wb.set_cell_contents("Sheet1", "A2", "=IFERROR()")
        assert correct_error(wb.get_cell_value("Sheet1", "A2"), CellErrorType.TYPE_ERROR)

        wb.set_cell_contents("Sheet1", "A2", "=IFERROR(A1)")
        assert wb.get_cell_value("Sheet1", "A2") == wb.get_cell_value("Sheet1", "A1")
        wb.set_cell_contents("Sheet1", "A1", "=90/0")
        assert wb.get_cell_value("Sheet1", "A2") == ""
        wb.set_cell_contents("Sheet1", "A2", "=IFERROR(A1, A3)")
        wb.set_cell_contents("Sheet1", "A3", "Slatt")
        assert wb.get_cell_value("Sheet1", "A2") == "Slatt"
                

    def test_ISERROR_cycle(self):
        wb = Workbook()
        wb.new_sheet()
        wb.set_cell_contents("Sheet1", "A1", "=B1")
        wb.set_cell_contents("Sheet1", "B1", "=A1")
        wb.set_cell_contents("Sheet1", "C1", "=iserror(B1)")
        assert correct_error(wb.get_cell_value("Sheet1", "A1"), CellErrorType.CIRCULAR_REFERENCE)
        assert correct_error(wb.get_cell_value("Sheet1", "B1"), CellErrorType.CIRCULAR_REFERENCE)
        assert wb.get_cell_value("Sheet1", "C1") == True

    #     wb.set_cell_contents("Sheet1", "A1", "=ISERROR(B1)")
    #     wb.set_cell_contents("Sheet1", "B1", "=ISERROR(A1)")
    #     assert correct_error(wb.get_cell_value("Sheet1", "A1"), CellErrorType.CIRCULAR_REFERENCE)
    #     assert correct_error(wb.get_cell_value("Sheet1", "B1"), CellErrorType.CIRCULAR_REFERENCE)
    #     assert wb.get_cell_value("Sheet1", "C1") == True
    
    def test_ISERROR_cycle2(self):
        wb = Workbook()
        wb.new_sheet()
        wb.set_cell_contents("Sheet1", "A1", "=ISERROR(B1)")
        wb.set_cell_contents("Sheet1", "B1", "=ISErrOR(A1)")
        wb.set_cell_contents("Sheet1", "C1", "=isERROR(B1)")
        assert correct_error(wb.get_cell_value("Sheet1", "A1"), CellErrorType.CIRCULAR_REFERENCE)
        assert correct_error(wb.get_cell_value("Sheet1", "B1"), CellErrorType.CIRCULAR_REFERENCE)
        assert wb.get_cell_value("Sheet1", "C1") == True


    def test_IF_cycle(self):
        wb = Workbook()
        wb.new_sheet()
        wb.set_cell_contents("Sheet1", "A3", "False")
        wb.set_cell_contents("Sheet1", "A1", "=A2")
        wb.set_cell_contents("Sheet1", "A2", "=IF(A3, A1, A3)")
        assert wb.get_cell_value("Sheet1", "A1") == False
        assert wb.get_cell_value("Sheet1", "A2") == False
        wb.set_cell_contents("Sheet1", "A3", "True")
        assert correct_error(wb.get_cell_value("Sheet1", "A1"), CellErrorType.CIRCULAR_REFERENCE)
        assert correct_error(wb.get_cell_value("Sheet1", "A2"), CellErrorType.CIRCULAR_REFERENCE)

    def test_IFERROR_2(self):
        wb = Workbook()
        wb.new_sheet()
        wb.set_cell_contents("Sheet1", "A3", "False")
        wb.set_cell_contents("Sheet1", "A1", "=A2")
        # Should A2 value should be set to False
        wb.set_cell_contents("Sheet1", "A2", "=IF(A3, A1, A3)")

        assert wb.get_cell_value("Sheet1", "A1") == False
        assert wb.get_cell_value("Sheet1", "A2") == False
        wb.set_cell_contents("Sheet1", "A3", "True")
        assert correct_error(wb.get_cell_value("Sheet1", "A1"), CellErrorType.CIRCULAR_REFERENCE)
        assert correct_error(wb.get_cell_value("Sheet1", "A2"), CellErrorType.CIRCULAR_REFERENCE)
    
    def test_IFERROR_3(self):
        wb = Workbook()
        _, n1 = wb.new_sheet()
        wb.set_cell_contents(n1, "A1", "=A1")
        wb.set_cell_contents(n1, "B1", "10")
        wb.set_cell_contents(n1, "A2", "=IFERROR        (A1, 1)")
        wb.set_cell_contents(n1, "A3", "=IfErroR(A1, #REF!)")
        wb.set_cell_contents(n1, "A4", "=iferror   (A1)")
        wb.set_cell_contents(n1, "A5", "=ifERROR(B1, #REF!)")
        wb.set_cell_contents(n1, "A6", "=IFERROR(B1)")
        wb.set_cell_contents(n1, "A7", "=IFERROR(#REF!, 1)")
        assert wb.get_cell_value(n1, "A2") == Decimal("1")
        assert correct_error(wb.get_cell_value(n1, "A3"), CellErrorType.BAD_REFERENCE)
        assert wb.get_cell_value(n1, "A4") == ""
        assert wb.get_cell_value(n1, "A5") == Decimal("10")
        assert wb.get_cell_value(n1, "A6") == Decimal("10")
        assert wb.get_cell_value(n1, "a7") == Decimal("1")

    def test_ISBLANK(self):
        wb = Workbook()
        wb.new_sheet()
        wb.set_cell_contents("Sheet1", "A1", "")
        wb.set_cell_contents("Sheet1", "A3", "=ISBLANK()")
        assert correct_error(wb.get_cell_value("Sheet1", "A3"), CellErrorType.TYPE_ERROR)
        wb.set_cell_contents("Sheet1", "A3", "=ISBLANK(True, A1)")
        assert correct_error(wb.get_cell_value("Sheet1", "A3"), CellErrorType.TYPE_ERROR)
        wb.set_cell_contents("Sheet1", "A3", "=ISBLANK(A22, A1)")
        assert correct_error(wb.get_cell_value("Sheet1", "A3"), CellErrorType.TYPE_ERROR)
        wb.set_cell_contents("Sheet1", "A22", "slatt")
        assert correct_error(wb.get_cell_value("Sheet1", "A3"), CellErrorType.TYPE_ERROR)

        wb.set_cell_contents("Sheet1", "A3", '=ISBLANK("sky!")')
        assert wb.get_cell_value("Sheet1", "A3") == False
        wb.set_cell_contents("Sheet1", "A3", '=ISBLANK("sky!", A1)')
        assert correct_error(wb.get_cell_value("Sheet1", "A3"), CellErrorType.TYPE_ERROR)

        wb.set_cell_contents("Sheet1", "A3", '=ISBLANK("")')
        assert not wb.get_cell_value("Sheet1", "A3")
        wb.set_cell_contents("Sheet1", "A1", "0")
        assert not wb.get_cell_value("Sheet1", "A3")
        wb.set_cell_contents("Sheet1", "A1", "False")
        assert not wb.get_cell_value("Sheet1", "A3")
        wb.set_cell_contents("Sheet1", "A3", '=ISBLANK(A333)')
        assert wb.get_cell_value("Sheet1", "A3")

    def test_ISBLANK_cycle(self):
        wb = Workbook()
        wb.new_sheet()
        wb.set_cell_contents("Sheet1", "A1", "=ISBLANK(A1)")
        assert correct_error(wb.get_cell_value("Sheet1", "A1"), CellErrorType.CIRCULAR_REFERENCE)

        wb.set_cell_contents("Sheet1", "A1", "=ISBLANK(A2)")
        wb.set_cell_contents("Sheet1", "A2", "=ISBLANK(A1)")
        assert correct_error(wb.get_cell_value("Sheet1", "A1"), CellErrorType.CIRCULAR_REFERENCE)
        assert correct_error(wb.get_cell_value("Sheet1", "A2"), CellErrorType.CIRCULAR_REFERENCE)

    

    def test_INDIRECT_error(self):
        wb = Workbook()
        _, name = wb.new_sheet()
        wb.set_cell_contents(name, "a2", "=INDIRECT(Sheet!!A1)")
        assert correct_error(wb.get_cell_value(name, "a2"), CellErrorType.PARSE_ERROR)
        wb.set_cell_contents(name, "a2", "=INDIRECT(Sheet2!A1)")
        assert correct_error(wb.get_cell_value(name, "a2"), CellErrorType.PARSE_ERROR)
        wb.set_cell_contents(name, "a2", '=INDIRECT("Sheet2!A1")')
        assert correct_error(wb.get_cell_value(name, "a2"), CellErrorType.BAD_REFERENCE)
        wb.set_cell_contents(name, "a2", '=INDIRECT("Sheet1!A9999999")')
        assert correct_error(wb.get_cell_value(name, "a2"), CellErrorType.BAD_REFERENCE)
        
        wb.set_cell_contents(name, "a1", '=a1')
        wb.set_cell_contents(name, "a2", f'=INDIRECT("{name}!A1")')
        assert correct_error(wb.get_cell_value(name, "a2"), CellErrorType.CIRCULAR_REFERENCE)

        wb.set_cell_contents(name, "a1", '=a2')
        wb.set_cell_contents(name, "a2", '=INDIRECT("A1")')
        assert correct_error(wb.get_cell_value(name, "a2"), CellErrorType.CIRCULAR_REFERENCE)

        wb.set_cell_contents(name, "a4", '=a5')
        wb.set_cell_contents(name, "a5", '4')
        wb.set_cell_contents(name, "a6", '=INDIRECT("A4")')
        assert wb.get_cell_value(name, "a6") == Decimal("4")
        wb.set_cell_contents(name, "a5", '=a6')
        assert correct_error(wb.get_cell_value(name, "a6"), CellErrorType.CIRCULAR_REFERENCE)

        #can't clear a1,a2,a3
        wb.set_cell_contents(name, "a1", '=a2')
        wb.set_cell_contents(name, "a2", '4')
        wb.set_cell_contents(name, "a3", '=INDIRECT("A1")')
        assert wb.get_cell_value(name, "a3") == Decimal("4")
        wb.set_cell_contents(name, "a2", '=a3')
        assert correct_error(wb.get_cell_value(name, "a3"), CellErrorType.CIRCULAR_REFERENCE)
                

    def test_invalid_func_name(self):
        wb = Workbook()
        _, n1 = wb.new_sheet()
        wb.set_cell_contents(n1, "A1", "=FART(True, True, False)")
        wb.set_cell_contents(n1, "A2", "=I(A1)")
        wb.set_cell_contents(n1, "A3", "=NOR(A1 < A2)")
        wb.set_cell_contents(n1, "A4", "=ILOVEcs(True, True, False)")
        wb.set_cell_contents(n1, "A5", "=HIGH(False)")
        wb.set_cell_contents(n1, "A6", "=LOW(True, True, False)")

        assert correct_error(wb.get_cell_value(n1, "A1"), CellErrorType.BAD_NAME)
        assert correct_error(wb.get_cell_value(n1, "A2"), CellErrorType.BAD_NAME)
        assert correct_error(wb.get_cell_value(n1, "A3"), CellErrorType.BAD_NAME)
        assert correct_error(wb.get_cell_value(n1, "A4"), CellErrorType.BAD_NAME)
        assert correct_error(wb.get_cell_value(n1, "A5"), CellErrorType.BAD_NAME)
        assert correct_error(wb.get_cell_value(n1, "A6"), CellErrorType.BAD_NAME)
        

    def test_simple_func(self):
        wb = Workbook()
        _, n1 = wb.new_sheet()
        wb.set_cell_contents(n1, "A1", "=A2")
        wb.set_cell_contents(n1, "A2", '=INDIRECT("A1")')
        assert correct_error(wb.get_cell_value(n1, "A1"), CellErrorType.CIRCULAR_REFERENCE)
        assert correct_error(wb.get_cell_value(n1, "A2"), CellErrorType.CIRCULAR_REFERENCE)

        wb.set_cell_contents(n1, "A2", "4")
        assert wb.get_cell_value(n1, "A1") == Decimal("4")
        assert wb.get_cell_value(n1, "A2") == Decimal("4")

        wb.set_cell_contents(n1, "A3", '=INDIRECT("A1")')
        assert wb.get_cell_value(n1, "A3") == Decimal("4")

        wb.set_cell_contents(n1, "A2", "=A3")

        assert correct_error(wb.get_cell_value(n1, "A1"), CellErrorType.CIRCULAR_REFERENCE)
        assert correct_error(wb.get_cell_value(n1, "A2"), CellErrorType.CIRCULAR_REFERENCE)
        assert correct_error(wb.get_cell_value(n1, "A3"), CellErrorType.CIRCULAR_REFERENCE)

    def test_CHOOSE(self):
        wb = Workbook()
        wb.new_sheet()
        wb.set_cell_contents("Sheet1", "A1", "0")
        wb.set_cell_contents("Sheet1", "A2", "slatt")
        wb.set_cell_contents("Sheet1", "A3", "haas")
        wb.set_cell_contents("Sheet1", "A4", "Camp")
        wb.set_cell_contents("Sheet1", "B1", "=CHOOSE()")
        assert correct_error(wb.get_cell_value("Sheet1", "B1"), CellErrorType.TYPE_ERROR)

        wb.set_cell_contents("Sheet1", "B1", "=CHOOSE(A1)")
        assert correct_error(wb.get_cell_value("Sheet1", "B1"), CellErrorType.TYPE_ERROR)

        wb.set_cell_contents("Sheet1", "B1", "=CHOOSE(A1, A2, A3, A4)")
        assert correct_error(wb.get_cell_value("Sheet1", "B1"), CellErrorType.TYPE_ERROR)

        wb.set_cell_contents("Sheet1", "A1", "1")
        assert wb.get_cell_value("Sheet1", "B1") == "slatt"

        wb.set_cell_contents("Sheet1", "A1", "False")
        assert correct_error(wb.get_cell_value("Sheet1", "B1"), CellErrorType.TYPE_ERROR)

        wb.set_cell_contents("Sheet1", "A1", "True")
        assert wb.get_cell_value("Sheet1", "B1") == "slatt"

        wb.set_cell_contents("Sheet1", "A1", "'2")
        assert wb.get_cell_value("Sheet1", "B1") == "haas"

        wb.set_cell_contents("Sheet1", "A1", "100")
        assert correct_error(wb.get_cell_value("Sheet1", "B1"), CellErrorType.TYPE_ERROR)

        wb.set_cell_contents("Sheet1", "B1", "=CHOOSE(A1, A2, A3, A4, A99)")
        wb.set_cell_contents("Sheet1", "A1", "4")
        assert wb.get_cell_value("Sheet1", "B1") == None

    def test_cell_range(self):
        wb = Workbook()
        _, n1 = wb.new_sheet()
        _, n2 = wb.new_sheet()
        wb.set_cell_contents(n1, "A1", "1")
        wb.set_cell_contents(n1, "A2", "2")
        wb.set_cell_contents(n1, "A3", "3")
        wb.set_cell_contents(n2, "A4", "=MAX(Sheet1!A1:A3)")
        assert(wb.get_cell_value(n2, "A4") == Decimal("3"))

    # def test_false_and_true_converts_to_string(self):
    #     wb = Workbook()
    #     wb.new_sheet()
    #     wb.set_cell_contents("Sheet1", "A1", "=False")
    #     wb.set_cell_contents("Sheet1", "B1", "=True")
    #     wb.set_cell_contents("Sheet1", "A2", "'Hello")
    #     wb.set_cell_contents("Sheet1", "A3", "=A1 & A2")
    #     assert isinstance(wb.get_cell_value('Sheet1', 'A1'), bool)
    #     assert(wb.get_cell_value("Sheet1", "A3") == "FALSEHello")
    #     wb.set_cell_contents("Sheet1", "A3", "=A1 & A2 & B1")
    #     assert(wb.get_cell_value("Sheet1", "A3") == "FALSEHelloTRUE")


    def test_IF_when_condition_zero(self):
        wb = Workbook()
        wb.new_sheet()
        wb.set_cell_contents("Sheet1", "A1", "Something ")
        wb.set_cell_contents("Sheet1", "B1", "0")
        wb.set_cell_contents("Sheet1", "C1", "something else")
        wb.set_cell_contents("Sheet1", "A1", "=IF(B1, C1, A1)")
        assert correct_error(wb.get_cell_value("Sheet1", "A1"), CellErrorType.CIRCULAR_REFERENCE)

    # def test_first_arg_missing_parse_error(self):
    #     wb = Workbook()
    #     wb.new_sheet()
    #     wb.set_cell_contents("Sheet1", "A1", "hello")
    #     wb.set_cell_contents("Sheet1", "A2", "bye")
    #     wb.set_cell_contents("Sheet1", "A3", "=EXACT(, A1)")
    #     assert correct_error(wb.get_cell_value("Sheet1", "A3"), CellErrorType.PARSE_ERROR)

    def test_hlookup(self):
        wb = Workbook()
        _, n1 = wb.new_sheet()
        wb.set_cell_contents(n1, "A1", "a")
        wb.set_cell_contents(n1, "A2", "b")
        wb.set_cell_contents(n1, "A3", "c")
        wb.set_cell_contents(n1, "B1", "4")
        wb.set_cell_contents(n1, "B2", "5")
        wb.set_cell_contents(n1, "B3", "6")
        wb.set_cell_contents(n1, "C1", "7")
        wb.set_cell_contents(n1, "C2", "8")
        wb.set_cell_contents(n1, "C3", "9")

        wb.set_cell_contents(n1, "A10", "=HLOOKUP(4, A1:C3, 2)")
        wb.set_cell_contents(n1, "A11", '=HLOOKUP("a", A1:C3, 2)')
        wb.set_cell_contents(n1, "A12", "=HLOOKUP(7, A1:C3, 4)")
        wb.set_cell_contents(n1, "C3", "=HLOOKUP(7, A1:C3, 3)")

        assert (wb.get_cell_value(n1, "A10") == Decimal("5"))
        assert (wb.get_cell_value(n1, "A11") == "b")
        assert correct_error(wb.get_cell_value(n1, "A12"), CellErrorType.TYPE_ERROR)
        assert correct_error(wb.get_cell_value(n1, "C3"), CellErrorType.CIRCULAR_REFERENCE)

    def test_vlookup(self):
        wb = Workbook()
        _, n1 = wb.new_sheet()
        wb.set_cell_contents(n1, "A1", "a")
        wb.set_cell_contents(n1, "A2", "b")
        wb.set_cell_contents(n1, "A3", "c")
        wb.set_cell_contents(n1, "B1", "4")
        wb.set_cell_contents(n1, "B2", "5")
        wb.set_cell_contents(n1, "B3", "6")
        wb.set_cell_contents(n1, "C1", "7")
        wb.set_cell_contents(n1, "C2", "8")
        wb.set_cell_contents(n1, "C3", "9")

        wb.set_cell_contents(n1, "A10", '=VLOOKUP("a", A1:C3, 1)')
        wb.set_cell_contents(n1, "A11", '=VLOOKUP("b", A1:C3, 2)')
        wb.set_cell_contents(n1, "A12", '=VLOOKUP("b", A1:C3, 4)')
        wb.set_cell_contents(n1, "C2", '=VLOOKUP("c", A1:C3, 3)')
        wb.set_cell_contents(n1, "C3", '=VLOOKUP("c", A1:C3, 3)')

        assert (wb.get_cell_value(n1, "A10") == "a")
        assert (wb.get_cell_value(n1, "A11") == Decimal("5"))
        assert correct_error(wb.get_cell_value(n1, "A12"), CellErrorType.TYPE_ERROR)
        assert correct_error(wb.get_cell_value(n1, "C2"), CellErrorType.CIRCULAR_REFERENCE)
        assert correct_error(wb.get_cell_value(n1, "C3"), CellErrorType.CIRCULAR_REFERENCE)


    def test_IF_with_cell_range(self):
        wb = Workbook()
        _, n1 = wb.new_sheet()
        wb.set_cell_contents(n1, "A1", "1")
        wb.set_cell_contents(n1, "A2", "2")
        wb.set_cell_contents(n1, "A3", "3")
        wb.set_cell_contents(n1, "B1", "4")
        wb.set_cell_contents(n1, "B2", "5")
        wb.set_cell_contents(n1, "B3", "6")
        wb.set_cell_contents(n1, "Z9", "=SUM(IF(C1, A1:A3, B1:B3))")

        wb.set_cell_contents(n1, "C1", "True")
        assert wb.get_cell_value(n1, "Z9") == Decimal("6")
        wb.set_cell_contents(n1, "C1", "False")
        assert wb.get_cell_value(n1, "Z9") == Decimal("15")

    def test_SUM_cycle(self):
        wb = Workbook()
        _, n1 = wb.new_sheet()
        wb.set_cell_contents(n1, "A1", "=SUM(A2:A4)")
        wb.set_cell_contents(n1, "A2", "1")
        wb.set_cell_contents(n1, "A3", "2")
        wb.set_cell_contents(n1, "A4", "=A1")

        assert correct_error(wb.get_cell_value(n1, "A1"), CellErrorType.CIRCULAR_REFERENCE)
        assert wb.get_cell_value(n1, "A2") == Decimal("1")
        assert wb.get_cell_value(n1, "A3") == Decimal("2")
        assert correct_error(wb.get_cell_value(n1, "A4"), CellErrorType.CIRCULAR_REFERENCE)


    def test_EXACT_wrong_input(self):
        wb = Workbook()
        _, n1 = wb.new_sheet()
        wb.set_cell_contents(n1, "A50", "=EXACT(A2:A5, B2:B5)")
        assert correct_error(wb.get_cell_value(n1, "A50"), CellErrorType.TYPE_ERROR)


if __name__ == "__main__":
    unittest.main()

