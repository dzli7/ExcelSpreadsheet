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


class WorkbookTests(unittest.TestCase):
    def test_add_three_sheets(self):
        wb = Workbook()
        idx, name = wb.new_sheet()
        assert (wb.sheets[0].get_name() == "Sheet1")
        assert (wb.sheets[0].get_index() == 0)
        assert (idx == 0)
        assert (name == "Sheet1")

        idx, name = wb.new_sheet("My New_sheet")
        assert (wb.sheets[1].get_name() == "My New_sheet")
        assert (wb.sheets[1].get_index() == 1)
        assert (idx == 1)
        assert (name == "My New_sheet")

        idx, name = wb.new_sheet()
        assert (wb.sheets[2].get_name() == "Sheet2")
        assert (wb.sheets[2].get_index() == 2)
        assert (idx == 2)
        assert (name == "Sheet2")

    def test_sheet_name_exists_exact_match(self):
        wb = Workbook()
        _idx, _name = wb.new_sheet()
        wb.new_sheet("my_sheet")
        with self.assertRaises(ValueError):
            wb.new_sheet("my_sheet")

    def test_sheet_name_exists_case_mismatch(self):
        wb = Workbook()
        _idx, _name = wb.new_sheet()
        wb.new_sheet("my_sheet")
        with self.assertRaises(ValueError):
            wb.new_sheet("My_Sheet")

    def test_sheet_name_case_preserved(self):
        wb = Workbook()
        _idx, name = wb.new_sheet("My Sheet")
        assert name == "My Sheet"
        assert wb.sheets[0].get_name() == "My Sheet"

    def test_sheet_names(self):
        wb = Workbook()
        _idx, name = wb.new_sheet("new sheet")
        assert (name == "new sheet")
        _idx, name = wb.new_sheet("new new sheet")
        assert (name == "new new sheet")
        _idx, name = wb.new_sheet(".?!,:;!@#$%^&*()-_")
        assert (name == ".?!,:;!@#$%^&*()-_")
        _idx, name = wb.new_sheet("abc.?!,:;!@#$%^&*()-_123")
        assert (name == "abc.?!,:;!@#$%^&*()-_123")

    def test_sheet_bad_name(self):
        wb = Workbook()
        with self.assertRaises(ValueError):
            wb.new_sheet("")
        with self.assertRaises(ValueError):
            wb.new_sheet(" ")
        with self.assertRaises(ValueError):
            wb.new_sheet("aa ")
        with self.assertRaises(ValueError):
            wb.new_sheet(" aa")
        with self.assertRaises(ValueError):
            wb.new_sheet(" aba ")
        with self.assertRaises(ValueError):
            wb.new_sheet("[abc]")
        with self.assertRaises(ValueError):
            wb.new_sheet("\{abc\}")
        with self.assertRaises(ValueError):
            wb.new_sheet("\"\"")
        with self.assertRaises(ValueError):
            wb.new_sheet("\'\'")

    def test_del_default_sheet(self):
        wb = Workbook()
        wb.new_sheet()
        assert (len(wb.sheets) == 1)
        assert (wb.default_sheet_nums == {1})

        wb.new_sheet()
        assert (len(wb.sheets) == 2)
        assert (wb.default_sheet_nums == {1, 2})

        wb.new_sheet()
        assert (len(wb.sheets) == 3)
        assert (wb.default_sheet_nums == {1, 2, 3})

        wb.del_sheet("Sheet1")
        assert (len(wb.sheets) == 2)
        assert (wb.sheets[0].get_name() == "Sheet2")
        assert (wb.sheets[0].get_index() == 0)
        assert (wb.sheets[1].get_name() == "Sheet3")
        assert (wb.sheets[1].get_index() == 1)
        assert (wb.default_sheet_nums == {2, 3})

        wb.new_sheet()
        assert (len(wb.sheets) == 3)
        assert (wb.sheets[0].get_name() == "Sheet2")
        assert (wb.sheets[0].get_index() == 0)
        assert (wb.sheets[1].get_name() == "Sheet3")
        assert (wb.sheets[1].get_index() == 1)
        assert (wb.sheets[2].get_name() == "Sheet1")
        assert (wb.sheets[2].get_index() == 2)
        assert (wb.default_sheet_nums == {1, 2, 3})

        wb.new_sheet("test sheet")
        assert (len(wb.sheets) == 4)
        assert (wb.sheets[3]._name == "test sheet")

        wb.del_sheet("test sheet")
        assert (len(wb.sheets) == 3)
        assert (wb.sheets[0].get_name() == "Sheet2")
        assert (wb.sheets[0].get_index() == 0)
        assert (wb.sheets[1].get_name() == "Sheet3")
        assert (wb.sheets[1].get_index() == 1)
        assert (wb.sheets[2].get_name() == "Sheet1")
        assert (wb.sheets[2].get_index() == 2)

        wb.del_sheet("Sheet3")
        assert (len(wb.sheets) == 2)
        assert (wb.sheets[0].get_name() == "Sheet2")
        assert (wb.sheets[0].get_index() == 0)
        assert (wb.sheets[1].get_name() == "Sheet1")
        assert (wb.sheets[1].get_index() == 1)
        assert (wb.default_sheet_nums == {1, 2})

        wb.del_sheet("Sheet2")
        assert (len(wb.sheets) == 1)
        assert (wb.sheets[0].get_name() == "Sheet1")
        assert (wb.sheets[0].get_index() == 0)
        assert (wb.default_sheet_nums == {1})

        wb.del_sheet("Sheet1")
        assert (len(wb.sheets) == 0)
        assert (len(wb.default_sheet_nums) == 0)

        with self.assertRaises(KeyError):
            wb.del_sheet("Sheet1")

    def test_multiple_workbooks(self):
        wb = Workbook()
        _idx1, name1 = wb.new_sheet()
        assert name1 == "Sheet1"
        _idx2, name2 = wb.new_sheet()
        assert name2 == "Sheet2"

        wb2 = Workbook()
        _idx3, name3 = wb2.new_sheet()
        assert name3 == "Sheet1"

        wb.set_cell_contents(name1, "a1", "'hello world")
        assert (wb.get_cell_value(name1, "a1") == "hello world")
        wb.set_cell_contents(name2, "a2", "'himmothy")
        assert (wb.get_cell_value(name2, "a2") == "himmothy")
        wb2.set_cell_contents(name3, "c2", "'2")
        assert (wb2.get_cell_value(name3, "c2") == Decimal(2))
        with self.assertRaises(KeyError):
            wb2.set_cell_contents(name2, "c2", "'2")

    def test_sheet_extent(self):
        wb = Workbook()
        _idx, name = wb.new_sheet()
        assert (name == "Sheet1")
        assert (wb.get_sheet_extent(name) == (0, 0))
        wb.set_cell_contents(name, "a1", "'2")
        assert (wb.get_sheet_extent(name) == (1, 1))
        wb.set_cell_contents(name, "b1", "'3")
        assert (wb.get_sheet_extent(name) == (2, 1))
        wb.set_cell_contents(name, "a2", "'4")
        assert (wb.get_sheet_extent(name) == (2, 2))
        wb.set_cell_contents(name, "f1", "'4")
        assert (wb.get_sheet_extent(name) == (6, 2))
        wb.set_cell_contents(name, "i8", "'4")
        assert (wb.get_sheet_extent(name) == (9, 8))

        with self.assertRaises(ValueError):
            wb.set_cell_contents(name, "i10000", "'4")

        wb.set_cell_contents(name, "aaa100", "'4")
        assert (wb.get_sheet_extent(name) == (703, 100))
        wb.set_cell_contents(name, "bbb222", "'4")
        assert (wb.get_sheet_extent(name) == (1406, 222))
        wb.set_cell_contents(name, "i1000", "'4")
        assert (wb.get_sheet_extent(name) == (1406, 1000))
        wb.set_cell_contents(name, "zzzz9999", "'4")
        assert (wb.get_sheet_extent(name) == (475254, 9999))

        wb.set_cell_contents(name, "c3", "'4")
        assert (wb.get_sheet_extent(name) == (475254, 9999))
        wb.set_cell_contents(name, "aaa100", "'5")
        assert (wb.get_sheet_extent(name) == (475254, 9999))
        wb.set_cell_contents(name, "ccc100", "'5")
        assert (wb.get_sheet_extent(name) == (475254, 9999))

    def test_decrease_extent(self):
        wb = Workbook()
        _idx, name = wb.new_sheet()
        assert (name == "Sheet1")
        assert (wb.get_sheet_extent(name) == (0, 0))

        wb.set_cell_contents(name, "b2", "'2")
        assert (wb.get_sheet_extent(name) == (2, 2))
        wb.set_cell_contents(name, "b2", "")
        assert (wb.get_sheet_extent(name) == (0, 0))

        wb.set_cell_contents(name, "c2", "'2")
        assert (wb.get_sheet_extent(name) == (3, 2))
        wb.set_cell_contents(name, "e4", "'3")
        assert (wb.get_sheet_extent(name) == (5, 4))
        wb.set_cell_contents(name, "e4", "")
        assert (wb.get_sheet_extent(name) == (3, 2))
        wb.set_cell_contents(name, "e4", "'3")
        assert (wb.get_sheet_extent(name) == (5, 4))
        wb.set_cell_contents(name, "e4", None)
        assert (wb.get_sheet_extent(name) == (3, 2))
        wb.set_cell_contents(name, "c2", None)
        assert (wb.get_sheet_extent(name) == (0, 0))

    def test_reference_other_sheet(self):
        wb = Workbook()
        _idx1, name1 = wb.new_sheet()
        _idx2, name2 = wb.new_sheet("cool sheet B)")
        _idx3, name3 = wb.new_sheet()
        _idx4, name4 = wb.new_sheet("!@#$()!")

        wb.set_cell_contents(name1, "a1", "2")
        wb.set_cell_contents(name2, "a1", "3")
        wb.set_cell_contents(name3, "a1", "4")
        wb.set_cell_contents(name4, "a1", "5")

        wb.set_cell_contents(name1, "a2", "=cool sheet B)!a1 * 2")
        value = wb.get_cell_value(name1, "a2")
        assert (wb.get_cell_contents(name1, "a2") == "=cool sheet B)!a1 * 2")
        assert (isinstance(value, CellError))
        assert (value.get_type() == CellErrorType.PARSE_ERROR)

        wb.set_cell_contents(name1, "a2", "='cool sheet B)'!a1 * 2")
        assert (wb.get_cell_value(name1, "a2") == Decimal('6'))

        wb.set_cell_contents(name2, "a2", "=Sheet1!a1 * 2")
        assert (wb.get_cell_value(name2, "a2") == Decimal('4'))

        wb.set_cell_contents(name3, "a2", "=!@#$()!!a1 * 2")
        value = wb.get_cell_value(name3, "a2")
        assert (wb.get_cell_contents(name3, "a2") == "=!@#$()!!a1 * 2")
        assert (isinstance(value, CellError))
        assert (value.get_type() == CellErrorType.PARSE_ERROR)

        wb.set_cell_contents(name3, "a2", "='!@#$()!'!a1 * 2")
        assert (wb.get_cell_value(name3, "a2") == Decimal('10'))

        wb.set_cell_contents(name4, "a2", "=Sheet10!a1 * 2")
        value = wb.get_cell_value(name4, "a2")
        assert (isinstance(value, CellError))
        assert (value.get_type() == CellErrorType.BAD_REFERENCE)

        _idx5, name5 = wb.new_sheet("sheet10")
        wb.set_cell_contents(name5, "a1", "6")
        assert (wb.get_cell_value(name4, "a2") == Decimal('12'))

    def test_sheet_name_has_space(self):
        wb = Workbook()
        wb.new_sheet("July Totals")
        wb.set_cell_contents("July Totals", "A1", "6")
        wb.new_sheet()
        wb.set_cell_contents("Sheet1", "A1", "='July Totals'!A1")
        value = wb.get_cell_value("Sheet1", "A1")
        assert value == Decimal('6')

    def test_adding_deleting_sheets_references(self):
        wb = Workbook()
        _, n1 = wb.new_sheet()
        _, n2 = wb.new_sheet()
        _, n3 = wb.new_sheet()
        wb.set_cell_contents(n1, "A1", ".02000000")
        wb.set_cell_contents(n2, "A2", "=Sheet1!A1")
        wb.set_cell_contents(n3, "A3", "=Sheet2!A2")
        wb.del_sheet(n2)
        assert (wb.get_cell_value(n1, "A1") == Decimal(".02"))
        value = wb.get_cell_value(n3, "A3")
        assert (isinstance(value, CellError))
        assert (value.get_type() == CellErrorType.BAD_REFERENCE)

    def test_save_workbook_function(self):
        wb = Workbook()
        wb.new_sheet()
        wb.set_cell_contents("Sheet1", "A1", "6")
        wb.set_cell_contents("Sheet1", "A2", "12")
        wb.set_cell_contents("Sheet1", "A3", "=A1 * A2")
        wb.new_sheet()
        wb.set_cell_contents("Sheet2", "B1", "9")
        wb.set_cell_contents("Sheet2", "B2", "18")
        wb.set_cell_contents("Sheet2", "B3", "=B2 / B1")
        with open("tests/test_json_objects/test_save_copy.json", 'w') as f:
            wb.save_workbook(f)

        with open("tests/test_json_objects/test_save_copy.json", 'r') as f:
            test_saved_wb = f.read()

        with open("tests/test_json_objects/test_save.json", 'r') as f:
            actual_saved_wb = f.read()

        assert test_saved_wb == actual_saved_wb

    def test_load_workbook_function(self):
        wb = Workbook.load_workbook("tests/test_json_objects/workbook1.json")
        assert (wb.get_cell_contents("Sheet1", "A1") == "'123")
        assert (wb.get_cell_contents("Sheet1", "B1") == "5.3")
        assert (wb.get_cell_value("Sheet1", "B1") == Decimal('5.3'))
        assert (wb.get_cell_contents("Sheet1", "C1") == "=A1+B1")
        assert (wb.get_cell_value("Sheet1", "C1") == Decimal('128.3'))

        assert (wb.get_cell_contents("Sheet2", "C1") == "=A1-B1")
        assert (wb.get_cell_value("Sheet2", "C1") == Decimal('118'))

    def test_move_sheets(self):
        wb = Workbook()
        wb.new_sheet()
        wb.new_sheet()
        wb.new_sheet()
        wb.new_sheet()
        wb.new_sheet()
        wb.move_sheet("Sheet5", 2)
        assert wb.sheet_to_idx["sheet1"] == 0
        assert wb.sheet_to_idx["sheet2"] == 1
        assert wb.sheet_to_idx["sheet5"] == 2
        assert wb.sheet_to_idx["sheet3"] == 3
        assert wb.sheet_to_idx["sheet4"] == 4

    def test_copy_sheets(self):
        wb = Workbook()
        wb.new_sheet()
        wb.set_cell_contents("Sheet1", "A1", "6")
        wb.set_cell_contents("Sheet1", "A2", "12")
        wb.set_cell_contents("Sheet1", "A3", "24")
        wb.copy_sheet("Sheet1")
        assert wb.sheet_num == 2
        assert wb.sheets[1]._name == "Sheet1_1"
        assert (wb.get_cell_value("Sheet1_1", "A1") == Decimal('6'))
        assert (wb.get_cell_value("Sheet1_1", "A2") == Decimal('12'))
        assert (wb.get_cell_value("Sheet1_1", "A3") == Decimal('24'))

        # Make sure values do not cross over
        wb.set_cell_contents("Sheet1", "A1", "69")
        wb.set_cell_contents("Sheet1_1", "A1", "419")
        assert (wb.get_cell_value("Sheet1", "A1") == Decimal('69')) and (
            (wb.get_cell_value("Sheet1_1", "A1") != Decimal('69')))
        assert (wb.get_cell_value("Sheet1_1", "A1") == Decimal('419')) and (
            (wb.get_cell_value("Sheet1", "A1") != Decimal('419')))

        wb.copy_sheet("Sheet1")
        assert wb.sheet_num == 3
        assert wb.sheets[2].get_name() == "Sheet1_2"
        assert (wb.get_cell_value("Sheet1_2", "A1") == Decimal('69'))

        wb.copy_sheet("Sheet1_1")
        assert wb.sheet_num == 4
        assert wb.sheets[3].get_name() == "Sheet1_1_1"
        assert (wb.get_cell_value("Sheet1_1_1", "A1") == Decimal('419'))

    def test_rename_sheets_bad_reference_to_value(self):
        wb = Workbook()
        _, n1 = wb.new_sheet()
        _, n2 = wb.new_sheet()
        wb.set_cell_contents(n1, "A1", "15")
        wb.set_cell_contents(n2, "A2", "=Sheet1!A1")
        wb.set_cell_contents(n2, "A3", "='Sheet 151'!A1")
        value = wb.get_cell_value(n2, "A3")
        assert (isinstance(value, CellError))
        assert (value.get_type() == CellErrorType.BAD_REFERENCE)
        wb.rename_sheet(n1, "Sheet 151")
        assert (wb.get_cell_contents(n2, "A2") == "='Sheet 151'!A1")
        assert (wb.get_cell_value(n2, "A2") == Decimal("15"))
        assert (wb.get_cell_value(n2, "A3") == Decimal("15"))

    def test_rename_sheets_quote_to_unquote(self):
        wb = Workbook()
        _, n1 = wb.new_sheet("Sheet 15")
        _, n2 = wb.new_sheet()
        wb.set_cell_contents(n1, "A1", "15")
        wb.set_cell_contents(n2, "A2", "='Sheet 15'!A1")
        assert (wb.get_cell_value(n2, "A2") == Decimal("15"))
        wb.rename_sheet(n1, "Sheet2")
        assert (wb.get_cell_value(n2, "A2") == Decimal("15"))
        assert (wb.get_cell_contents(n2, "A2") == "=Sheet2!A1")

    def test_rename_sheet_names(self):
        wb = Workbook()
        wb.new_sheet()
        wb.new_sheet()
        wb.new_sheet("Sheet 15")
        wb.new_sheet()
        wb.new_sheet("as39 kd")
        assert (wb.list_sheets() == [
                "Sheet1", "Sheet2", "Sheet 15", "Sheet3", "as39 kd"])
        wb.rename_sheet("Sheet3", "Sheet5")
        wb.rename_sheet("Sheet5", "nice try")
        assert (wb.list_sheets() == [
                "Sheet1", "Sheet2", "Sheet 15", "nice try", "as39 kd"])

    def test_rename_sheets_invalid_rename(self):
        wb = Workbook()
        _, n1 = wb.new_sheet()
        _, n2 = wb.new_sheet()
        _, n3 = wb.new_sheet()
        with self.assertRaises(ValueError):
            wb.rename_sheet("Sheet1", "Sheet2")
        with self.assertRaises(ValueError):
            wb.rename_sheet("Sheet2", "Sheet3")
        with self.assertRaises(ValueError):
            wb.rename_sheet("Sheet3", "Sheet1")

    def test_rename_sheets_default_name(self):
        wb = Workbook()
        _, n1 = wb.new_sheet()
        wb.rename_sheet(n1, "Sheet10")
        wb.del_sheet("Sheet10")
        assert(wb.list_sheets() == [])

    def test_rename_sheets_reference_other_sheet(self):
        wb = Workbook()
        _, n1 = wb.new_sheet()
        _, n2 = wb.new_sheet()
        _, n3 = wb.new_sheet()
        wb.set_cell_contents(n1, "A1", ".02000000")
        wb.set_cell_contents(n2, "A2", "=Sheet1!A1")
        wb.set_cell_contents(n3, "A3", "=Sheet2!A2")
        wb.del_sheet(n2)
        assert (wb.get_cell_value(n1, "A1") == Decimal(".02"))
        value = wb.get_cell_value(n3, "A3")
        assert (isinstance(value, CellError))
        assert (value.get_type() == CellErrorType.BAD_REFERENCE)

    def test_rename_sheet_preserves_contents(self):
        wb = Workbook()
        _, name1 = wb.new_sheet()
        _, name2 = wb.new_sheet()
        wb.set_cell_contents(name1, "a1", "2")
        wb.set_cell_contents(name2, "a1", f"=(({name1}!a1))   *   ((2))")
        assert (wb.get_cell_contents(name2, "a1") == f"=(({name1}!a1))   *   ((2))")

        name3 = "My Sheet"
        wb.rename_sheet(name1, name3)
        assert (wb.get_cell_contents(name2, "a1") == f"=(('{name3}'!a1)) * ((2))")

    def test_rename_sheet_fixes_refs(self):
        wb = Workbook()
        _, name1 = wb.new_sheet()
        _, name2 = wb.new_sheet()
        name3 = "My Sheet"
        wb.set_cell_contents(name1, "a1", "2")
        wb.set_cell_contents(name2, "a1", f"='{name3}'!a1 * 2")
        assert (wb.get_cell_contents(name2, "a1") == f"='{name3}'!a1 * 2")

        value = wb.get_cell_value(name2, "a1")
        assert isinstance(value, CellError)
        assert value.get_type() == CellErrorType.BAD_REFERENCE

        wb.rename_sheet(name1, name3)
        assert (wb.get_cell_contents(name2, "a1") == f"='{name3}'!a1 * 2")
        assert (wb.get_cell_value(name2, "a1") == Decimal("4"))

        wb.del_sheet(name3)
        value = wb.get_cell_value(name2, "a1")
        assert isinstance(value, CellError)
        assert value.get_type() == CellErrorType.BAD_REFERENCE

    def test_rename_sheet_invalid_to_valid(self):
        wb = Workbook()
        _, n1 = wb.new_sheet()
        wb.set_cell_contents(n1, "A1", "=A2")
        wb.set_cell_contents(n1, "A2", "=A3")
        wb.set_cell_contents(n1, "A3", "=A4")
        wb.set_cell_contents(n1, "A4", "1")
        n2 = "Sheeet4"
        wb.rename_sheet(n1, n2)
        assert wb.get_cell_value(n2, "A1") == Decimal("1")
        assert wb.get_cell_value(n2, "A2") == Decimal("1")
        assert wb.get_cell_value(n2, "A3") == Decimal("1")
        assert wb.get_cell_value(n2, "A4") == Decimal("1")




    def test_save_workbook_function(self):
        wb = Workbook()
        wb.new_sheet()
        wb.set_cell_contents("Sheet1", "A1", "6")
        wb.set_cell_contents("Sheet1", "A2", "12")
        wb.set_cell_contents("Sheet1", "A3", "=A1 * A2")
        wb.new_sheet()
        wb.set_cell_contents("Sheet2", "B1", "9")
        wb.set_cell_contents("Sheet2", "B2", "18")
        wb.set_cell_contents("Sheet2", "B3", "=B2 / B1")
        with open("tests/test_json_objects/test_save_copy.json", 'w') as f:
            wb.save_workbook(f)

        with open("tests/test_json_objects/test_save_copy.json", 'r') as f:
            test_saved_wb = f.read()

        with open("tests/test_json_objects/test_save.json", 'r') as f:
            actual_saved_wb = f.read()

        assert test_saved_wb == actual_saved_wb

    def test_load_workbook_function(self):
        wb = Workbook.load_workbook("tests/test_json_objects/workbook1.json")
        assert (wb.get_cell_contents("Sheet1", "A1") == "'123")
        assert (wb.get_cell_contents("Sheet1", "B1") == "5.3")
        assert (wb.get_cell_value("Sheet1", "B1") == Decimal('5.3'))
        assert (wb.get_cell_contents("Sheet1", "C1") == "=A1+B1")
        assert (wb.get_cell_value("Sheet1", "C1") == Decimal('128.3'))

        assert (wb.get_cell_contents("Sheet2", "C1") == "=A1-B1")
        assert (wb.get_cell_value("Sheet2", "C1") == Decimal('118'))

    def test_save_small_workbook(self):
        with io.StringIO() as fp:
            wb = Workbook()
            wb.new_sheet("HeLlO WoRlD")
            wb.set_cell_contents("HeLlO WoRlD", "A1", "6")
            wb.set_cell_contents("HeLlO WoRlD", "A2", "12")
            wb.set_cell_contents("HeLlO WoRlD", "A3", "=A1 * A2")
            wb.new_sheet()
            wb.set_cell_contents("Sheet1", "B1", "9")
            wb.set_cell_contents("Sheet1", "B2", "18")
            wb.set_cell_contents("Sheet1", "B3", "=B2 / B1")
            wb.save_workbook(fp)
            fp.seek(0)
            res = json.load(fp)

            ans = {
                "sheets":[
                    {
                        "name":"HeLlO WoRlD",
                        "cell-contents":{
                            "A1":"6",
                            "A2":"12",
                            "A3":"=A1 * A2"
                        }
                    },
                    {
                        "name":"Sheet1",
                        "cell-contents":{
                            "B1":"9",
                            "B2":"18",
                            "B3":"=B2 / B1"
                        }
                    }
                ]
            }

            assert (res == ans)
        
    def test_move_sheets(self):
        wb = Workbook()
        wb.new_sheet()
        wb.new_sheet()
        wb.new_sheet()
        wb.new_sheet()
        wb.new_sheet()
        wb.move_sheet("Sheet5", 2)
        assert wb.sheet_to_idx["sheet1"] == 0
        assert wb.sheet_to_idx["sheet2"] == 1
        assert wb.sheet_to_idx["sheet5"] == 2
        assert wb.sheet_to_idx["sheet3"] == 3
        assert wb.sheet_to_idx["sheet4"] == 4

    def test_copy_sheets(self):
        wb = Workbook()
        wb.new_sheet()
        wb.set_cell_contents("Sheet1", "A1", "6")
        wb.set_cell_contents("Sheet1", "A2", "12")
        wb.set_cell_contents("Sheet1", "A3", "24")
        wb.copy_sheet("Sheet1")
        assert wb.sheet_num == 2
        assert wb.sheets[1].get_name() == "Sheet1_1"
        assert (wb.get_cell_value("Sheet1_1", "A1") == Decimal('6'))
        assert (wb.get_cell_value("Sheet1_1", "A2") == Decimal('12'))
        assert (wb.get_cell_value("Sheet1_1", "A3") == Decimal('24'))

        # Make sure values do not cross over
        wb.set_cell_contents("Sheet1", "A1", "69")
        wb.set_cell_contents("Sheet1_1", "A1", "419")
        assert (wb.get_cell_value("Sheet1", "A1") == Decimal('69')) and (
            (wb.get_cell_value("Sheet1_1", "A1") != Decimal('69')))
        assert (wb.get_cell_value("Sheet1_1", "A1") == Decimal('419')) and (
            (wb.get_cell_value("Sheet1", "A1") != Decimal('419')))

        wb.copy_sheet("Sheet1")
        assert wb.sheet_num == 3
        assert wb.sheets[2].get_name() == "Sheet1_2"
        assert (wb.get_cell_value("Sheet1_2", "A1") == Decimal('69'))

        wb.copy_sheet("Sheet1_1")
        assert wb.sheet_num == 4
        assert wb.sheets[3].get_name() == "Sheet1_1_1"
        assert (wb.get_cell_value("Sheet1_1_1", "A1") == Decimal('419'))

        assert(wb.list_sheets() == ["Sheet1", "Sheet1_1", "Sheet1_2", "Sheet1_1_1"])

    def test_rename_sheets_bad_reference_to_value(self):
        wb = Workbook()
        _, n1 = wb.new_sheet()
        _, n2 = wb.new_sheet()
        wb.set_cell_contents(n1, "A1", "15")
        wb.set_cell_contents(n2, "A2", "=Sheet1!A1")
        wb.set_cell_contents(n2, "A3", "='Sheet 15'!A1")
        value = wb.get_cell_value(n2, "A3")
        assert (isinstance(value, CellError))
        assert (value.get_type() == CellErrorType.BAD_REFERENCE)
        wb.rename_sheet(n1, "Sheet 15")
        assert (wb.get_cell_contents(n2, "A2") == "='Sheet 15'!A1")
        assert (wb.get_cell_value(n2, "A2") == Decimal("15"))
        assert (wb.get_cell_value(n2, "A3") == Decimal("15"))

    def test_rename_sheets_quote_to_unquote(self):
        wb = Workbook()
        _, n1 = wb.new_sheet("Sheet 15")
        _, n2 = wb.new_sheet()
        wb.set_cell_contents(n1, "A1", "15")
        wb.set_cell_contents(n2, "A2", "='Sheet 15'!A1")
        assert (wb.get_cell_value(n2, "A2") == Decimal("15"))
        wb.rename_sheet(n1, "Sheet2")
        assert (wb.get_cell_value(n2, "A2") == Decimal("15"))
        assert (wb.get_cell_contents(n2, "A2") == "=Sheet2!A1")

    def test_rename_sheet_names(self):
        wb = Workbook()
        wb.new_sheet()
        wb.new_sheet()
        wb.new_sheet("Sheet 15")
        wb.new_sheet()
        wb.new_sheet("as39 kd")
        assert (wb.list_sheets() == [
                "Sheet1", "Sheet2", "Sheet 15", "Sheet3", "as39 kd"])
        wb.rename_sheet("Sheet3", "Sheet5")
        wb.rename_sheet("Sheet5", "nice try")
        assert (wb.list_sheets() == [
                "Sheet1", "Sheet2", "Sheet 15", "nice try", "as39 kd"])

    def test_rename_sheets_invalid_rename(self):
        wb = Workbook()
        _, n1 = wb.new_sheet()
        _, n2 = wb.new_sheet()
        _, n3 = wb.new_sheet()
        with self.assertRaises(ValueError):
            wb.rename_sheet("Sheet1", "Sheet2")
        with self.assertRaises(ValueError):
            wb.rename_sheet("Sheet2", "Sheet3")
        with self.assertRaises(ValueError):
            wb.rename_sheet("Sheet3", "Sheet1")

    def test_rename_sheets_default_name(self):
        wb = Workbook()
        _, n1 = wb.new_sheet()
        _, n2 = wb.new_sheet()
        wb.rename_sheet(n1, "Sheet10")
        wb.set_cell_contents(n2, "A1", "=Sheet1!A1")
        value = wb.get_cell_value(n2, "A1")
        assert(isinstance(value, CellError))
        assert(value.get_type() == CellErrorType.BAD_REFERENCE)
        wb.del_sheet(n2)
        assert(wb.list_sheets() == ["Sheet10"])
        

        # This test prints update notifications

    # def test_cell_update_notifications_prints(self):
    #     wb = Workbook()
    #     def on_cells_changed(workbook, changed_cells):
    #         print(f'Cell(s) changed:  {changed_cells}')
    #     wb.notify_cells_changed(on_cells_changed)
    #     _, n1 = wb.new_sheet()
    #     _, n2 = wb.new_sheet()
    #     wb.set_cell_contents(n1, "A1", "'123")
    #     wb.set_cell_contents(n1, "C1", "=A1+B1")
    #     wb.set_cell_contents(n1, "B1", "6.9")
    #     wb.set_cell_contents(n1, "D1", "=Sheet5!A2")
    #     value = wb.get_cell_value(n1, "D1")
    #     assert(isinstance(value, CellError))
    #     assert(value.get_type() == CellErrorType.BAD_REFERENCE)
    #     wb.rename_sheet(n2, "Sheet5")
    #     wb.set_cell_contents("Sheet5", "A2", "4")
    #     assert(wb.get_cell_value(n1, "D1") == Decimal("4"))

    def test_cell_update_notifications_values(self):
        wb = Workbook()
        cells = []

        def on_cells_changed(workbook, changed_cells):
            cells.append(changed_cells)
        wb.notify_cells_changed(on_cells_changed)
        _, n1 = wb.new_sheet()
        _, n2 = wb.new_sheet()
        wb.set_cell_contents(n1, "A1", "'123")
        assert (len(cells[0]) == 1)
        assert (cells[0] == [("Sheet1", "A1")])
        cells.pop()
        wb.set_cell_contents(n1, "C1", "=A1+B1")
        assert (len(cells[0]) == 1)
        assert (cells[0] == [("Sheet1", "C1")])
        cells.pop()
        wb.set_cell_contents(n1, "B1", "6.9")
        assert (len(cells[0]) == 2)
        assert (set(cells[0]) == set([("Sheet1", "B1"), ("Sheet1", "C1")]))
        cells.pop()
        wb.set_cell_contents(n1, "D1", "=Sheet5!A2")
        assert (len(cells[0]) == 1)
        assert (cells[0] == [("Sheet1", "D1")])
        value = wb.get_cell_value(n1, "D1")
        assert (isinstance(value, CellError))
        assert (value.get_type() == CellErrorType.BAD_REFERENCE)
        cells.pop()
        wb.rename_sheet(n2, "Sheet5")
        assert (len(cells[0]) == 1)
        assert (cells[0] == [("Sheet1", "D1")])
        assert (wb.get_cell_value("Sheet1", "D1") is None)
        cells.pop()
        wb.set_cell_contents("Sheet5", "A2", "4")
        assert (len(cells[0]) == 2)
        assert (set(cells[0]) == set([("Sheet5", "A2"), ("Sheet1", "D1")]))
        assert (wb.get_cell_value("Sheet1", "D1") == Decimal("4"))


    def test_sheet_extent2(self):
        wb = Workbook()
        _, n1 = wb.new_sheet()
        wb.set_cell_contents(n1, "A1", "=D4")
        assert(wb.get_sheet_extent(n1) == (1,1))

    def test_rename_with_errors(self):
        wb = Workbook()
        _, n1 = wb.new_sheet("Some")
        wb.set_cell_contents(n1, "A1", "123")
        wb.set_cell_contents(n1, "A2", "=12+12")
        wb.set_cell_contents(n1, "A3", "=Some!A2")
        wb.set_cell_contents(n1, "A4", "=#CIRCREF!")
        wb.set_cell_contents(n1, "A5", "=Some!A4")
        n2 = "Other"
        wb.rename_sheet(n1, n2)
        assert(wb.get_cell_value(n2, "A1") == Decimal("123"))
        assert(wb.get_cell_value(n2, "A2") == Decimal("24"))
        assert(wb.get_cell_value(n2, "A3") == Decimal("24"))
        assert(wb.get_cell_contents(n2, "A3") == "=Other!A2")
        val1 = wb.get_cell_value(n2, "A4")
        assert isinstance(val1, CellError)
        assert val1.get_type() == CellErrorType.CIRCULAR_REFERENCE
        assert(wb.get_cell_contents(n2, "A4") == "=#CIRCREF!")
        val2 = wb.get_cell_value(n2, "A5")
        assert isinstance(val2, CellError)
        assert val2.get_type() == CellErrorType.CIRCULAR_REFERENCE
        assert(wb.get_cell_contents(n2, "A5") == "=Other!A4")

    def test_copy_row_of_cells(self):
        wb = Workbook()
        _, name = wb.new_sheet()
        wb.set_cell_contents(name, "A1", "1")
        wb.set_cell_contents(name, "B1", "hi")
        wb.set_cell_contents(name, "C1", "'3")
        wb.set_cell_contents(name, "D1", "70")
        wb.copy_cells(name, 'A1', 'D1', 'A2')
        assert wb.get_cell_value(name, "A2") == wb.get_cell_value(name, "A1")
        assert wb.get_cell_value(name, "B2") == wb.get_cell_value(name, "B1")
        assert wb.get_cell_value(name, "C2") == wb.get_cell_value(name, "C1")
        assert wb.get_cell_value(name, "D2") == wb.get_cell_value(name, "D1")

        assert wb.get_cell_contents(name, "A2") == wb.get_cell_contents(name, "A1")
        assert wb.get_cell_contents(name, "B2") == wb.get_cell_contents(name, "B1")
        assert wb.get_cell_contents(name, "C2") == wb.get_cell_contents(name, "C1")
        assert wb.get_cell_contents(name, "D2") == wb.get_cell_contents(name, "D1")

        wb.copy_cells(name, 'D1', 'A1', 'E1')
        assert wb.get_cell_value(name, "E1") == wb.get_cell_value(name, "A1")
        assert wb.get_cell_value(name, "F1") == wb.get_cell_value(name, "B1")
        assert wb.get_cell_value(name, "G1") == wb.get_cell_value(name, "C1")
        assert wb.get_cell_value(name, "H1") == wb.get_cell_value(name, "D1")

        assert wb.get_cell_contents(name, "E1") == wb.get_cell_contents(name, "A1")
        assert wb.get_cell_contents(name, "F1") == wb.get_cell_contents(name, "B1")
        assert wb.get_cell_contents(name, "G1") == wb.get_cell_contents(name, "C1")
        assert wb.get_cell_contents(name, "H1") == wb.get_cell_contents(name, "D1")

        wb.copy_cells(name, 'D1', 'A1', 'AA1')
        assert wb.get_cell_value(name, "AA1") == wb.get_cell_value(name, "A1")
        assert wb.get_cell_value(name, "AB1") == wb.get_cell_value(name, "B1")
        assert wb.get_cell_value(name, "AC1") == wb.get_cell_value(name, "C1")
        assert wb.get_cell_value(name, "AD1") == wb.get_cell_value(name, "D1")

        assert wb.get_cell_contents(name, "AA1") == wb.get_cell_contents(name, "A1")
        assert wb.get_cell_contents(name, "AB1") == wb.get_cell_contents(name, "B1")
        assert wb.get_cell_contents(name, "AC1") == wb.get_cell_contents(name, "C1")
        assert wb.get_cell_contents(name, "AD1") == wb.get_cell_contents(name, "D1")

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
        wb.copy_cells(name, "A1", "C3", "B2")
        # Remain unchanged
        assert wb.get_cell_value(name, "A1") == Decimal('1')
        assert wb.get_cell_value(name, "A2") == "hi"
        assert wb.get_cell_value(name, "A3") == Decimal('17')
        assert wb.get_cell_value(name, "B1") == Decimal('3')
        assert wb.get_cell_value(name, "C1") == "bye"
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

    def test_copy_cells_with_ref(self):
        wb = Workbook()
        _, name = wb.new_sheet()
        wb.set_cell_contents(name, "A1", "1")
        wb.set_cell_contents(name, "B1", "2")
        wb.set_cell_contents(name, "C1", "=A1 + B1")
        wb.set_cell_contents(name, "A2", "3")
        wb.set_cell_contents(name, "B2", "4")
        wb.copy_cells(name, 'C1', 'C1', 'C2')

        assert wb.get_cell_value(name, "C2") == Decimal('7')
        assert wb.get_cell_contents(name, "C2") == ("=A2 + B2")

        wb.copy_cells(name, 'C2', 'A2', 'A3')
        
        assert wb.get_cell_value(name, "C3") == Decimal('7')
        assert wb.get_cell_contents(name, "C3") == ("=A3 + B3")

        wb.copy_cells(name, 'A3', 'C3', 'A99')
        
        assert wb.get_cell_value(name, "C99") == Decimal('7')
        assert wb.get_cell_contents(name, "C99") == ("=A99 + B99")


    def test_copy_cell_to_bad_ref(self):
        wb = Workbook()
        _, n1 = wb.new_sheet()
        wb.set_cell_contents(n1, "A1", "2")
        wb.set_cell_contents(n1, "A2", "3")
        wb.set_cell_contents(n1, "B1", "4")
        wb.set_cell_contents(n1, "B2", "5")
        wb.set_cell_contents(n1, "C1", "=A1*B1")
        wb.set_cell_contents(n1, "C2", "=A2*B2")
        wb.copy_cells(n1, "C1", "C2", "B1")
        v1 = wb.get_cell_value(n1, "B1")
        assert isinstance(v1, CellError)
        assert v1.get_type() == CellErrorType.BAD_REFERENCE
        v2 = wb.get_cell_value(n1, "B2")
        assert isinstance(v2, CellError)
        assert v2.get_type() == CellErrorType.BAD_REFERENCE

    def test_load_empty_json(self):
        wb = Workbook.load_workbook("tests/test_json_objects/empty.json")
        assert wb.num_sheets() == 0

    def test_move_cells_overlap_dependent(self):
        wb = Workbook()
        _, n1 = wb.new_sheet()
        wb.set_cell_contents(n1, "A1", "=A2+B1")
        wb.set_cell_contents(n1, "A2", "2")
        wb.set_cell_contents(n1, "B1", "=$B2")
        wb.set_cell_contents(n1, "B2", "1")
        wb.move_cells(n1, "A1", "B2", "B2")
        assert(wb.get_cell_contents(n1, "B2") == "=B3 + C2")
        assert(wb.get_cell_contents(n1, "B3") == "2")
        assert(wb.get_cell_contents(n1, "C2") == "=$B3")
        assert(wb.get_cell_contents(n1, "C3") == "1")

    def test_load_single_sheet(self):
        wb = Workbook.load_workbook("tests/test_json_objects/one_sheet_wb.json")
        assert len(wb.sheets) == 1
        assert wb.sheets[0].get_name() == "Sheet17"
        assert wb.get_cell_contents("Sheet17", "A1") == "'123"
        assert wb.get_cell_contents("Sheet17", "B1") == "5.3"
        assert wb.get_cell_contents("Sheet17", "C1") == "=A1+B1"
        
        assert wb.get_cell_value("Sheet17", "A1") == Decimal("123")
        assert wb.get_cell_value("Sheet17", "B1") == Decimal("5.3")
        assert wb.get_cell_value("Sheet17", "C1") == Decimal("128.3")

        

if __name__ == "__main__":
    unittest.main()
