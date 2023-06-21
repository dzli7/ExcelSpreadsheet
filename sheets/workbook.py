# ==============================================================================
# Caltech CS130 - Winter 2023
#
# This file specifies the API that we expect your implementation to conform to.
# You will likely want to move these classes into various files, but the tests
# will expect these to be available when the "sheets" module is imported.

# If you are unfamiliar with Python 3 type annotations, see the Python standard
# library documentation for the typing module here:
#
#     https://docs.python.org/3/library/typing.html
#
# NOTE:  THIS FILE WILL NOT WORK AS-IS.  You are expected to incorporate it
#        into your project in whatever way you see fit.
from __future__ import annotations

from collections import defaultdict

import itertools
import re
import copy
import json
from lark import Lark, exceptions

from typing import Optional, List, TextIO, Tuple, Any
from itertools import count, filterfalse

from .sheet import Sheet
from .cell import Cell
from .cellerror import CellError
from .cellerrortype import CellErrorType
from .formula_parser import FormulaParser
from .formula_editor import FormulaEditor
from .graph import Graph
from .functions import functions
from .sorter import rowAdapterObject

# Illegal literal values map to CellErrorType
_ERROR_TYPES = {"#ERROR!": CellErrorType.PARSE_ERROR,
                "#CIRCREF": CellErrorType.CIRCULAR_REFERENCE,
                "#REF!": CellErrorType.BAD_REFERENCE,
                "#NAME?": CellErrorType.BAD_NAME,
                "#VALUE!": CellErrorType.TYPE_ERROR,
                "#DIV/0!": CellErrorType.DIVIDE_BY_ZERO
                }

class Workbook:
    # A workbook containing zero or more named spreadsheets.
    # Any and all operations on a workbook that may affect calculated cell
    # values should cause the workbook's contents to be updated properly.

    PARSER = Lark.open('sheets/formulas.lark', start='formula')

    def __init__(self):
        # Initialize a new empty workbook.

        # List containing all of the sheets in the Workbook
        self.sheets = []
        # Dict mapping lower case sheet names to index in the self.sheets list
        self.sheet_to_idx = {}
        # Number of sheets
        self.sheet_num = 0
        # Next default number if no sheet_name is input
        self.default_sheet_nums = set()
        # Formula parser
        # self.parser = Lark.open('sheets/formulas.lark', start='formula')
        # Sheet name -> [ list of cells]
        self.floating_cells = defaultdict(list)
        # (cell sheet name, cell loc) -> cell object
        self.bad_ref_cells = defaultdict(Cell)
        # Notifications cells list
        self.notification_functions = []
        # Cells whose values have been updated
        self.updated_cells = []
        # get_functions returns dictionary {func_name_str : func()}
        self.functions = functions()
        self.fp = FormulaParser(self)


    def get_functions(self):
        # Return the dictionary of functions defined by functions.py
        return self.functions

    def num_sheets(self) -> int:
        # Return the number of spreadsheets in the workbook.
        return self.sheet_num

    def get_sheet(self, sheet_name: str) -> Sheet:
        # Return the sheet specificied by sheet_name

        # Check if the sheet name is valid
        if sheet_name.lower() not in list(self.sheet_to_idx.keys()):
            raise KeyError(f"{sheet_name} is not a valid sheet!")

        # Find the sheet from dict
        return self.sheets[self.sheet_to_idx[sheet_name.lower()]]

    def list_sheets(self) -> List[str]:
        # Return a list of the spreadsheet names in the workbook, with the
        # capitalization specified at creation, and in the order that the sheets
        # appear within the workbook.
        #
        # In this project, the sheet names appear in the order that the user
        # created them; later, when the user is able to move and copy sheets,
        # the ordering of the sheets in this function's result will also reflect
        # such operations.
        #
        # A user should be able to mutate the return-value without affecting the
        # workbook's internal state.
        return [sheet.get_name() for sheet in self.sheets]

    def new_sheet(self, sheet_name: Optional[str] = None) -> Tuple[int, str]:
        # Add a new sheet to the workbook.  If the sheet name is specified, it
        # must be unique.  If the sheet name is None, a unique sheet name is
        # generated.  "Uniqueness" is determined in a case-insensitive manner,
        # but the case specified for the sheet name is preserved.
        #
        # The function returns a tuple with two elements:
        # (0-based index of sheet in workbook, sheet name).  This allows the
        # function to report the sheet's name when it is auto-generated.
        #
        # If the spreadsheet name is an empty string (not None), or it is
        # otherwise invalid, a ValueError is raised.
        pattern = re.compile(
            "^[a-zA-Z0-9.?!,:;!@#$%^&*()\\-_]+(?: +[a-zA-Z0-9.?!,:;!@#$%^&*()\\-_]+)*$")
        pattern_num = re.compile("^[0-9]+$")

        # If no sheet_name is input, sheet name is defaulted to "Sheet#"
        # Here, "#" is the next default sheet_num
        if sheet_name is None:
            num = self._new_default_sheet_num()
            sheet_name = f"Sheet{(num)}"
            self.default_sheet_nums.add(num)

        # Raise value error in the input sheet_name already exists or
        # there are invalid characters in the sheet_name
        elif sheet_name.lower() in self.sheet_to_idx.keys() or not re.match(pattern, sheet_name):
            raise ValueError("Invalid name")

        # Check if a sheet is made with format "Sheet{n}", where n in any
        # number
        elif sheet_name[:5].lower() == "sheet" and re.match(pattern_num, sheet_name[5:]):
            self.default_sheet_nums.add(int(sheet_name[5:]))

        # Create sheet, update sheet_to_idx, update sheets, update number of
        # sheets
        sheet = Sheet(self.sheet_num, sheet_name)
        self.sheet_to_idx[sheet_name.lower()] = self.sheet_num
        self.sheets.append(sheet)
        self.sheet_num += 1

        # If this sheet is in our floating cells list, we want to fill out the new sheet with
        # previous cells that were assigned to this sheet, keeping the contents
        # and children
        if sheet_name.lower() in self.floating_cells.keys():
            cells = self.floating_cells[sheet_name.lower()]
            for cell in cells:
                cell_sheet_name, cell_loc, child_cell_contents = cell.get_info()
                for (child_sheet_name, child_loc) in cell.get_children_cells():
                    self.set_cell_contents(
                        child_sheet_name, child_loc, child_cell_contents, False)
                self.set_cell_contents(cell_sheet_name, cell_loc, None, False)

            # Remove this sheet from the floating cells "workbook"
            del self.floating_cells[sheet_name.lower()]

        # Update any cells that might correspond to the new sheet
        self._update_bad_ref(sheet_name)
        self._notify_cells_helper(self.updated_cells)
        return (self.sheet_num - 1, sheet_name)

    def del_sheet(self, sheet_name: str) -> None:
        # Delete the spreadsheet with the specified name.
        #
        # The sheet name match is case-insensitive; the text must match but the
        # case does not have to.
        #
        # If the specified sheet name is not found, a KeyError is raised.

        # Check if sheet_name exists
        if sheet_name.lower() not in self.sheet_to_idx.keys():
            raise KeyError("Sheet not found")

        sheet = self.sheets[self.sheet_to_idx[sheet_name.lower()]]
        cells = sheet.get_cells()

        # If the name is default name, remove the number from the set of
        # default nums
        pattern = re.compile("^Sheet[0-9]+$")
        if re.match(pattern, sheet_name):
            num = int(sheet_name[5:])
            self.default_sheet_nums.remove(num)

        # Remove the sheet_name from the indexing dictionary
        idx = self.sheet_to_idx.pop(sheet_name.lower())

        # Update the indices of each sheet
        for i in range(idx + 1, self.sheet_num):
            temp = self.sheets[i]
            temp.set_index(temp.get_index() - 1)
            self.sheet_to_idx[temp.get_name().lower()] -= 1
        # Remove the sheet from self.sheets
        self.sheets.pop(idx)
        self.sheet_num -= 1

        # Update all cells affected
        self._update_cells(cells, sheet)

    def get_sheet_extent(self, sheet_name: str) -> Tuple[int, int]:
        # Return a tuple (num-cols, num-rows) indicating the current extent of
        # the specified spreadsheet.
        #
        # The sheet name match is case-insensitive; the text must match but the
        # case does not have to.
        #
        # If the specified sheet name is not found, a KeyError is raised.

        # Get extent of sheet
        try:
            val = self.get_sheet(sheet_name).get_extent()
        except KeyError:
            raise KeyError(f"{sheet_name} is not a valid sheet!")
        return val

    def set_cell_contents(self, sheet_name: str, location: str,
                          contents: Optional[str], notify: bool = True) -> None:
        # Set the contents of the specified cell on the specified sheet.
        #
        # The sheet name match is case-insensitive; the text must match but the
        # case does not have to.  Additionally, the cell location can be
        # specified in any case.
        #
        # If the specified sheet name is not found, a KeyError is raised.
        # If the cell location is invalid, a ValueError is raised.
        #
        # A cell may be set to "empty" by specifying a contents of None.
        #
        # Leading and trailing whitespace are removed from the contents before
        # storing them in the cell.  Storing a zero-length string "" (or a
        # string composed entirely of whitespace) is equivalent to setting the
        # cell contents to None.
        #
        # If the cell contents appear to be a formula, and the formula is
        # invalid for some reason, this method does not raise an exception;
        # rather, the cell's value will be a CellError object indicating the
        # naure of the issue.

        if sheet_name[0] == "'" and sheet_name[-1] == "'":
            sheet_name = sheet_name[1:-1]

        # Check that the (sheet_name, location) tuple is valid
        sheet, loc = self._check_loc(sheet_name, location.upper())

        # Store the previous cell if it exists
        prev_cell = None
        if loc in sheet.get_cells().keys():
            prev_cell = sheet.get_cell(loc)
        elif (sheet_name,loc) in self.bad_ref_cells.keys():
            prev_cell = self.bad_ref_cells.pop((sheet_name,loc))

        # Create a new cell with the new given contents and strip it
        new_cell = Cell(sheet_name, loc, contents)
        content = new_cell.parse_cell()
        value = None
        prev_parent_cells = []
        new_parent_cells = []
        # Carry the children cells over from previous cell if it exists
        if prev_cell is not None:
            new_cell.set_children_cells(prev_cell.get_children_cells())
            prev_parent_cells = set(prev_cell.get_parent_cells())

        # If the new_cell is a formula:
        if new_cell.get_formula_cell_flag():
            try:
                # Parse the content of the formula and get/set the value
                tree = self.PARSER.parse(content)
                self.fp.new_parsing(sheet_name, self)
                value = self.fp.visit(tree)  
                new_cell.set_value(value)
                # Get the parent cells from the formula parser
                new_parent_cells = set(self.fp.get_parent_cells())

            except (exceptions.LarkError, exceptions.UnexpectedCharacters) as e:
                # If the formula can't be parsed, then set the value
                # accordingly
                new_cell.set_value(
                    CellError(
                        CellErrorType.PARSE_ERROR,
                        detail=f'{contents} could not be parsed'))
        # If the new_cell is not a formula
        else:
            # If the contents is an error, set value as error
            if content in _ERROR_TYPES.keys():
                new_cell.set_value(CellError(_ERROR_TYPES[content], detail=f"{str(content)}"))
            # Otherwise, set the value as is
            else:
                new_cell.set_value(content)

        # If the parent cells are not the same, then remove the children from the previous parents
        if prev_parent_cells != new_parent_cells:
            for (p_sheet, p_loc) in prev_parent_cells:
                try:
                    p_cell = self.sheets[self.sheet_to_idx[p_sheet.lower()]].get_cell(p_loc.upper())
                    p_cell_children = p_cell.get_children_cells()
                    p_cell_children.remove((sheet_name.lower(), loc.upper()))
                    p_cell.set_children_cells(p_cell_children)
                except (KeyError, ValueError):
                    continue
        
        # Update the new parent cells
        new_cell.set_parent_cells(list(new_parent_cells))

        # Check if the value is updated for the new cell
        if (not prev_cell and new_cell.get_value() is not None) or (
                prev_cell is not None and new_cell.get_value() != prev_cell.get_value()):
            u_sheet_name, u_loc, _ = new_cell.get_info()
            self.updated_cells.append((u_sheet_name, u_loc))

        # Check if the new cell is a bad reference
        if isinstance(new_cell.get_value(), CellError) and new_cell.get_value(
        ).get_type() == CellErrorType.BAD_REFERENCE:
            self.bad_ref_cells[(sheet_name.lower(), loc.upper())] = new_cell

        # Update sheet.cells with the new_cell
        sheet.set_cell(loc, new_cell)

        # Update sheet extent
        sheet.update_sheet_extent()
        # Update the children_cells of the parent cells
        self._update_children_cells(new_cell, sheet_name, loc)
        # Check for circular dependencies in the newly set cell by iterating through children
        # until the first cell visited is encountered again
        circular_dependency_detected, visited_children = \
                self._check_circular_dependency(sheet_name, loc)
        

        # If there is circular dependency: new_cell value is the error
        # Otherwise update the values of all the children of new_cell
        if circular_dependency_detected:
            for cell in visited_children:
                # GETTING OVERWRITTEN HERE
                # if not isinstance(cell.get_value(), CellError):
                cell.set_value(
                    CellError(
                        CellErrorType.CIRCULAR_REFERENCE,
                        detail="Circular reference detected"))
        # If there is no circular dependency, create a dependency graph
        else:
            if new_cell.get_children_cells() != []:
                dependency_graph = Graph(defaultdict(list), 0)
                dependency_graph = self._build_dependency_graph(
                    new_cell, dependency_graph)
                cells_in_dependency_graph = []
                # Go through all keys (parent node) and values (children nodes)
                # to find all nodes in graph
                for key in dependency_graph.graph:
                    cells_in_dependency_graph.append(key)
                    for value in dependency_graph.graph[key]:
                        cells_in_dependency_graph.append(value)
                # Ensure no duplicates
                cells_in_dependency_graph = list(
                    set(cells_in_dependency_graph))
                num_of_cells = len(cells_in_dependency_graph)
                # Convert Cell objects to numeric equivalents to perform
                # topological sort
                dependency_graph_numeric_form = defaultdict(list)
                # 2-way mapping of Cell objects to an index
                # Index used as proxy to Cell object after topological sort
                index_to_cell_dict = {}  # { Cell object : index (integer) }
                cell_to_index_dict = {}  # { index: Cell object }
                for i, cell in enumerate(cells_in_dependency_graph):
                    index_to_cell_dict[i] = cell
                    cell_to_index_dict[cell] = i

                # Build dependency graph in numeric form without recursion
                for i, cell in enumerate(cell_to_index_dict.keys()):
                    if cell.get_children_cells == []:
                        continue
                    for child_cell_tuple in cell.get_children_cells():
                        child_cell_sheet_name = child_cell_tuple[0]
                        child_cell_location = child_cell_tuple[1]
                        child_cell_sheet_idx = self.sheet_to_idx[child_cell_sheet_name.lower(
                        )]
                        child_cell = self.sheets[child_cell_sheet_idx].get_cell(
                            child_cell_location)
                        parent_cell_numeric = cell_to_index_dict[cell]
                        child_cell_numeric = cell_to_index_dict[child_cell]

                        dependency_graph_numeric_form[parent_cell_numeric].append(
                            child_cell_numeric)
                        
                numeric_g = Graph(dependency_graph_numeric_form, num_of_cells)
                # numeric_g.graph = dependency_graph_numeric_form
                # numeric_g.v = num_of_cells
                # Perform topological sort
                order = numeric_g.topological_sort()
                cell_order_evaluation = []
                for pos in order[1:]:
                    if index_to_cell_dict[pos] not in cell_order_evaluation:
                        cell_order_evaluation.append(index_to_cell_dict[pos])

                # Update the values of the cells according to the order
                # given by the topological sort
                for cell in cell_order_evaluation:
                    sheet_name, loc, contents = cell.get_info()
                    old_val = cell.get_value()
                    if (sheet_name, loc, contents) != new_cell.get_info():
                        # If the new cell is a formula, parse it and set the value
                        if cell.get_formula_cell_flag():
                            try:
                                tree = self.PARSER.parse(contents)
                                self.fp.new_parsing(sheet_name, self)
                                value = self.fp.visit(tree)  
                                cell.set_value(value)
                                cell.set_parent_cells(list(set(self.fp.get_parent_cells())))
                            except (exceptions.LarkError, exceptions.UnexpectedCharacters) as e:
                                # If the formula can't be parsed, then set the
                                # value accordingly
                                cell.set_value(
                                    CellError(
                                        CellErrorType.PARSE_ERROR,
                                        detail=f'{contents} could not be parsed'))
                        # If the cell is not a formula, we dont't need to
                        # update anything
                        else:
                            pass
                    
                    # Update the children cells of the new cell
                    self._update_children_cells(cell, sheet_name, loc)

                    # Add the cell to nofitications if the value changes
                    if cell.get_value() != old_val:
                        self.updated_cells.append((sheet_name, loc))

                # With the new cell contents, check if there is a circular dependency
                circular_dependency_detected, visited_children = \
                        self._check_circular_dependency(sheet_name, loc)
                
                # If there is a circular dependency, set all relevant cells to
                # the CIRCULAR_REFERENCE error
                if circular_dependency_detected:
                    for cell in visited_children:
                        cell.set_value(
                            CellError(
                                CellErrorType.CIRCULAR_REFERENCE,
                                detail="Circular reference detected"))
        # Notify cells that have been updated
        # Check that some cell values have been updated
        if len(self.updated_cells) != 0 and notify:
            self._notify_cells_helper(self.updated_cells)

    def get_cell_contents(
            self,
            sheet_name: str,
            location: str) -> Optional[str]:
        # Return the contents of the specified cell on the specified sheet.
        #
        # The sheet name match is case-insensitive; the text must match but the
        # case does not have to.  Additionally, the cell location can be
        # specified in any case.
        #
        # If the specified sheet name is not found, a KeyError is raised.
        # If the cell location is invalid, a ValueError is raised.
        #
        # Any string returned by this function will not have leading or trailing
        # whitespace, as this whitespace will have been stripped off by the
        # set_cell_contents() function.
        #
        # This method will never return a zero-length string; instead, empty
        # cells are indicated by a value of None.

        # Check if (sheet name, location) exists
        sheet, loc = self._check_loc(sheet_name, location)

        # if location doesn't exist: return None
        if loc not in sheet.get_cells().keys():
            return None

        # Return the contents of the sheet
        return sheet.get_cell(loc).get_contents()

    def get_cell_value(self, sheet_name: str, location) -> Any:
        # Return the evaluated value of the specified cell on the specified
        # sheet.
        #
        # The sheet name match is case-insensitive; the text must match but the
        # case does not have to.  Additionally, the cell location can be
        # specified in any case.
        #
        # If the specified sheet name is not found, a KeyError is raised.
        # If the cell location is invalid, a ValueError is raised.
        #
        # The value of empty cells is None.  Non-empty cells may contain a
        # value of str, decimal.Decimal, or CellError.
        #
        # Decimal values will not have trailing zeros to the right of any
        # decimal place, and will not include a decimal place if the value is a
        # whole number.  For example, this function would not return
        # Decimal('1.000'); rather it would return Decimal('1').

        # Check if (sheet name, location) exists
        sheet, loc = self._check_loc(sheet_name, location)

        # if location doesn't exist: return None
        if loc not in sheet.get_cells().keys():
            return None

        # Return the contents of the sheet
        return sheet.get_cell(loc).get_value()

    def save_workbook(self, fp: TextIO) -> None:
        # Instance method (not a static/class method) to save a workbook to a
        # text file or file-like object in JSON format.  Note that the _caller_
        # of this function is expected to have opened the file; this function
        # merely writes the file.
        #
        # If an IO write error occurs (unlikely but possible), let any raised
        # exception propagate through.

        json_dict = {}
        # The json will be formatted so that
        # "sheets": sheets_value_list
        sheets_value_list = []
        for sheet in self.sheets:
            # The json will be formatted so that
            # "cell-contents": json_formatted_cell_dict
            json_formatted_cell_dict = {}
            for key in sheet.get_cells():
                cell_object = sheet.get_cells()[key]
                cell_object_contents = cell_object.get_contents()
                # Only non-empty cells in workbook get saved
                if self.get_cell_value(sheet.get_name(), key) is not None:
                    json_formatted_cell_dict[key.upper(
                    )] = cell_object_contents

            sheets_dict = {
                "name": sheet.get_name(),
                "cell-contents": json_formatted_cell_dict
            }
            sheets_value_list.append(sheets_dict)

        json_dict["sheets"] = sheets_value_list

        json.dump(json_dict, fp)

    @staticmethod
    def load_workbook(fp: TextIO) -> Workbook:
        # This is a static method (not an instance method) to load a workbook
        # from a text file or file-like object in JSON format, and return the
        # new Workbook instance.  Note that the _caller_ of this function is
        # expected to have opened the file; this function merely reads the file.
        #
        # If the contents of the input cannot be parsed by the Python json
        # module then a json.JSONDecodeError should be raised by the method.
        # (Just let the json module's exceptions propagate through.)  Similarly,
        # if an IO read error occurs (unlikely but possible), let any raised
        # exception propagate through.
        #
        # If any expected value in the input JSON is missing (e.g. a sheet
        # object doesn't have the "cell-contents" key), raise a KeyError with
        # a suitably descriptive message.
        #
        # If any expected value in the input JSON is not of the proper type
        # (e.g. an object instead of a list, or a number instead of a string),
        # raise a TypeError with a suitably descriptive message.
        wb = Workbook()
        with open(fp, 'r') as f:
            json_str = f.read()
        loaded_workbook_dict = json.loads(json_str)
        # Loading empty json should result in empty workbook
        if not loaded_workbook_dict:
            return wb

        if "sheets" not in loaded_workbook_dict:
            raise KeyError('Expected "sheets" key in json')

        if not isinstance(loaded_workbook_dict["sheets"], list):
            raise TypeError("Sheets expected to be a list")

        # Accessing "sheets" list from json
        sheets_in_wb = loaded_workbook_dict["sheets"]
        for sheet in sheets_in_wb:
            # Each element in sheets_in_wb list (iterating through value
            # of key "sheets") should be of type dictionary with keys
            # "name" and "cell-contents"
            if not isinstance(sheet, dict):
                raise TypeError(
                    "A json sheet representation should be a dictionary")

            if "name" not in sheet:
                raise KeyError("A sheet does not have a name input")

            sheet_name = sheet['name']
            if not isinstance(sheet_name, str):
                raise TypeError("A sheet has a name that is not a string")

            if "cell-contents" not in sheet:
                raise KeyError("A sheet does not have a cell-contents input")
            # Accessing "cell-contents" dictionary from json
            sheet_cells = sheet['cell-contents']
            if not isinstance(sheet_cells, dict):
                raise TypeError(
                    "A sheet has cell-contents that is not in dictionary form")

            # Add new sheet and associated cells if no errors so far
            wb.new_sheet(sheet_name)
            for cell in sheet_cells:
                if not isinstance(cell, str):
                    raise TypeError("A cell location is not a string")

                cell_content = sheet_cells[cell]
                if not isinstance(cell_content, str):
                    raise TypeError("A cell's content is not a string")

                wb.set_cell_contents(sheet_name, cell, cell_content)
        return wb

    def move_sheet(self, sheet_name: str, index: int):
        # Move the specified sheet to the specified index in the workbook's
        # ordered sequence of sheets.  The index can range from 0 to
        # workbook.num_sheets() - 1.  The index is interpreted as if the
        # specified sheet were removed from the list of sheets, and then
        # re-inserted at the specified index.
        #
        # The sheet name match is case-insensitive; the text must match but the
        # case does not have to.
        #
        # If the specified sheet name is not found, a KeyError is raised.
        #
        # If the index is outside the valid range, an IndexError is raised.

        # Check if sheet name already exists in workbook
        if sheet_name.lower() not in list(self.sheet_to_idx.keys()):
            raise KeyError(f"{sheet_name} is not a valid sheet!")

        # Check if index to be moved to is a valid index
        if index > self.sheet_num - 1 or index < 0:
            raise IndexError("Sheet index out of bounds for workbook")

        # Get index of sheet that is being moved
        old_sheet_index = self.sheet_to_idx[sheet_name.lower()]
        sheet_to_be_moved = self.sheets.pop(old_sheet_index)
        # Move sheet to new index
        self.sheets.insert(index, sheet_to_be_moved)
        self.sheet_to_idx[sheet_name.lower()] = index
        # Update indexes of sheets in sheet_name to index dictionary after
        # insertion
        for i in range(index + 1, self.sheet_num):
            sheet = self.sheets[i]
            name_of_sheet = sheet.get_name()
            self.sheet_to_idx[name_of_sheet.lower()] = i

    def copy_sheet(self, sheet_name: str) -> Tuple[int, str]:
        # Make a copy of the specified sheet, storing the copy at the end of the
        # workbook's sequence of sheets.  The copy's name is generated by
        # appending "_1", "_2", ... to the original sheet's name (preserving the
        # original sheet name's case), incrementing the number until a unique
        # name is found.  As usual, "uniqueness" is determined in a
        # case-insensitive manner.
        #
        # The sheet name match is case-insensitive; the text must match but the
        # case does not have to.
        #
        # The copy should be added to the end of the sequence of sheets in the
        # workbook.  Like new_sheet(), this function returns a tuple with two
        # elements:  (0-based index of copy in workbook, copy sheet name).  This
        # allows the function to report the new sheet's name and index in the
        # sequence of sheets.
        #
        # If the specified sheet name is not found, a KeyError is raised.

        # Check if sheet name already exists in workbook
        if sheet_name.lower() not in list(self.sheet_to_idx.keys()):
            raise KeyError(f"{sheet_name} is not a valid sheet!")

        # Access the sheet that is going to be copied
        sheet_to_be_copied = self.sheets[self.sheet_to_idx[sheet_name.lower()]]
        sheet_name_preserved_case = sheet_to_be_copied.get_name()
        cells_to_be_copied = sheet_to_be_copied.get_cells()

        # Find first available name that satisfies sheet copy naming convention
        appended_num = 1
        copy_name = sheet_name_preserved_case + "_" + str(appended_num)
        while copy_name.lower() in list(self.sheet_to_idx.keys()):
            appended_num += 1
            copy_name = sheet_name_preserved_case + "_" + str(appended_num)

        # Add new sheet with name of copy
        self.new_sheet(copy_name)

        # Set cells in new sheet to be cells of copied sheet
        for cell_loc in cells_to_be_copied:
            old_cell_contents = cells_to_be_copied[cell_loc].get_contents()
            self.set_cell_contents(copy_name, cell_loc, old_cell_contents, notify=False)

        # Update any cells that have parents with sheet name copy_cell
        self._update_bad_ref(copy_name)
        self._notify_cells_helper(self.updated_cells)
        return (self.sheet_num - 1, copy_name)

    def rename_sheet(self, sheet_name: str, new_sheet_name: str) -> None:
        # Rename the specified sheet to the new sheet name.  Additionally, all
        # cell formulas that referenced the original sheet name are updated to
        # reference the new sheet name (using the same case as the new sheet
        # name, and single-quotes iff [if and only if] necessary).
        #
        # The sheet_name match is case-insensitive; the text must match but the
        # case does not have to.
        #
        # As with new_sheet(), the case of the new_sheet_name is preserved by
        # the workbook.
        #
        # If the sheet_name is not found, a KeyError is raised.
        #
        # If the new_sheet_name is an empty string or is otherwise invalid, a
        # ValueError is raised.

        # Check if the sheet_name or new_sheet_name is invalid
        pattern = re.compile(
            "^[a-zA-Z0-9.?!,:;!@#$%^&*()\\-_]+(?: +[a-zA-Z0-9.?!,:;!@#$%^&*()\\-_]+)*$")
        if not re.match(
                pattern,
                sheet_name) or not re.match(
                pattern,
                new_sheet_name):
            raise ValueError("Invalid sheet name")

        # Check if the new sheet name exists already
        if new_sheet_name.lower() in self.sheet_to_idx.keys():
            raise ValueError("Invalid name")

        # Check if the old sheet name exists
        try:
            sheet = self.get_sheet(sheet_name.lower())
        except KeyError:
            raise KeyError(f"{sheet_name} does not exist")
        pattern = re.compile("^Sheet[0-9]+$")
        if re.match(pattern, new_sheet_name):
            num = int(new_sheet_name[5::])
            self.default_sheet_nums.add(num)

        pattern = re.compile("^Sheet[0-9]+$")
        if re.match(pattern, sheet_name):
            num = int(sheet_name[5::])
            self.default_sheet_nums.remove(num)

        # Set the new name of the sheet
        temp_sheet = copy.deepcopy(sheet)
        sheet.set_name(new_sheet_name)

        # Updating sheet name to idx dictionary
        idx = self.sheet_to_idx.pop(sheet_name.lower())
        self.sheet_to_idx[new_sheet_name.lower()] = idx
        self.sheets[idx] = sheet

        # Temporarily keep the old sheet for computation
        self.sheets.append(temp_sheet)
        self.sheet_to_idx[sheet_name.lower()] = len(self.sheets) - 1
        # Update the sheets list with the newly named sheet
        self.sheets[idx] = sheet
        # Update all related cells in the sheet
        cells = sheet.get_cells()
        for cell_loc in cells:
            # For each cell in the sheet:
            cell = sheet.get_cell(cell_loc)
            # For each child of the cell
            for (child_sheet, child_loc) in cell.get_children_cells():
                child = self.get_sheet(child_sheet.lower()).get_cell(child_loc)
                # if the child is a formula, we will compute on it
                if child.get_formula_cell_flag():
                    # Compute the new formula with the changed name
                    contents = self._update_formula(
                        child, sheet_name, new_sheet_name, None)
                    # Update the contents/value of the cell
                    self.set_cell_contents(child_sheet, child_loc, contents, notify=False)
                    # Remove the tuple of the old name and add the new one to
                    # the parents
            # For each parent of the cell
            for (parent_sheet, parent_loc) in cell.get_parent_cells():
                parent = self.get_sheet(
                    parent_sheet.lower()).get_cell(parent_loc)
                # Remove the tuple of the old name and add the new one to the
                # children
                try:
                    children_of_parents = parent.get_children_cells()
                    children_of_parents.remove((sheet_name.lower(), cell_loc.upper()))
                    children_of_parents.append((new_sheet_name.lower(), cell_loc.upper()))
                    parent.set_children_cells(children_of_parents)
                except ValueError:
                    continue

            # If the cell is a formula, update it with the new sheet name if we need to
            # Recalculate the value of the cell with the new formula
            if cell.get_formula_cell_flag():
                contents = self._update_formula(
                    cell, sheet_name, new_sheet_name, None)
            else:
                contents = cell.get_contents()
            self.set_cell_contents(new_sheet_name, cell_loc, contents, notify=False)

        idx = self.sheet_to_idx.pop(sheet_name.lower())
        del self.sheets[idx]
        # Update any cells that may have parents with new_sheet_name
        self._notify_cells_helper(self.updated_cells)
        self._update_bad_ref(new_sheet_name)

    def notify_cells_changed(self,
                             notify_function) -> None:
        # Request that all changes to cell values in the workbook are reported
        # to the specified notify_function.  The values passed to the notify
        # function are the workbook, and an iterable of 2-tuples of strings,
        # of the form ([sheet name], [cell location]).  The notify_function is
        # expected not to return any value; any return-value will be ignored.
        #
        # Multiple notification functions may be registered on the workbook;
        # functions will be called in the order that they are registered.
        #
        # A given notification function may be registered more than once; it
        # will receive each notification as many times as it was registered.
        #
        # If the notify_function raises an exception while handling a
        # notification, this will not affect workbook calculation updates or
        # calls to other notification functions.
        #
        # A notification function is expected to not mutate the workbook or
        # iterable that it is passed to it.  If a notification function violates
        # this requirement, the behavior is undefined.
        self.notification_functions.append(notify_function)

    def move_cells(
            self,
            sheet_name: str,
            start_location: str,
            end_location: str,
            to_location: str,
            to_sheet: Optional[str] = None) -> None:
        # Move cells from one location to another, possibly moving them to
        # another sheet.  All formulas in the area being moved will also have
        # all relative and mixed cell-references updated by the relative
        # distance each formula is being copied.
        #
        # Cells in the source area (that are not also in the target area) will
        # become empty due to the move operation.
        #
        # The start_location and end_location specify the corners of an area of
        # cells in the sheet to be moved.  The to_location specifies the
        # top-left corner of the target area to move the cells to.
        #
        # Both corners are included in the area being moved; for example,
        # copying cells A1-A3 to B1 would be done by passing
        # start_location="A1", end_location="A3", and to_location="B1".
        #
        # The start_location value does not necessarily have to be the top left
        # corner of the area to move, nor does the end_location value have to be
        # the bottom right corner of the area; they are simply two corners of
        # the area to move.
        #
        # This function works correctly even when the destination area overlaps
        # the source area.
        #
        # The sheet name matches are case-insensitive; the text must match but
        # the case does not have to.
        #
        # If to_sheet is None then the cells are being moved to another
        # location within the source sheet.
        #
        # If any specified sheet name is not found, a KeyError is raised.
        # If any cell location is invalid, a ValueError is raised.
        #
        # If the target area would extend outside the valid area of the
        # spreadsheet (i.e. beyond cell ZZZZ9999), a ValueError is raised, and
        # no changes are made to the spreadsheet.
        #
        # If a formula being moved contains a relative or mixed cell-reference
        # that will become invalid after updating the cell-reference, then the
        # cell-reference is replaced with a #REF! error-literal in the formula.
        cells_to_copy, destination_cell_loc_strings, destination_cell_contents, shift =  \
            self._copy_and_move_helper(sheet_name, start_location, end_location, to_location, to_sheet)

        if to_sheet is None:
            to_sheet = sheet_name

        # "Delete" contents of cells that were selected to be moved
        for cell in cells_to_copy:
            self.set_cell_contents(sheet_name, cell.get_info()[1], None, False)

        # "Paste" contents of copied cells into destination cells
        for i, cell_loc_str in enumerate(destination_cell_loc_strings):
            # Update formula of cell, then set contents
            self.set_cell_contents(
                to_sheet,
                cell_loc_str,
                destination_cell_contents[i], False)
            
        self._notify_cells_helper(self.updated_cells)

    def copy_cells(
            self,
            sheet_name: str,
            start_location: str,
            end_location: str,
            to_location: str,
            to_sheet: Optional[str] = None) -> None:
        # Copy cells from one location to another, possibly copying them to
        # another sheet.  All formulas in the area being copied will also have
        # all relative and mixed cell-references updated by the relative
        # distance each formula is being copied.
        #
        # Cells in the source area (that are not also in the target area) are
        # left unchanged by the copy operation.
        #
        # The start_location and end_location specify the corners of an area of
        # cells in the sheet to be copied.  The to_location specifies the
        # top-left corner of the target area to copy the cells to.
        #
        # Both corners are included in the area being copied; for example,
        # copying cells A1-A3 to B1 would be done by passing
        # start_location="A1", end_location="A3", and to_location="B1".
        #
        # The start_location value does not necessarily have to be the top left
        # corner of the area to copy, nor does the end_location value have to be
        # the bottom right corner of the area; they are simply two corners of
        # the area to copy.
        #
        # This function works correctly even when the destination area overlaps
        # the source area.
        #
        # The sheet name matches are case-insensitive; the text must match but
        # the case does not have to.
        #
        # If to_sheet is None then the cells are being copied to another
        # location within the source sheet.
        #
        # If any specified sheet name is not found, a KeyError is raised.
        # If any cell location is invalid, a ValueError is raised.
        #
        # If the target area would extend outside the valid area of the
        # spreadsheet (i.e. beyond cell ZZZZ9999), a ValueError is raised, and
        # no changes are made to the spreadsheet.
        #
        # If a formula being copied contains a relative or mixed cell-reference
        # that will become invalid after updating the cell-reference, then the
        # cell-reference is replaced with a #REF! error-literal in the formula.
        _, destination_cell_loc_strings, destination_cell_contents, shift =  \
            self._copy_and_move_helper(sheet_name, start_location, end_location, to_location, to_sheet)

        if to_sheet is None:
            to_sheet = sheet_name

        # "Paste" contents of copied cells into destination cells
        for i, cell_loc_str in enumerate(destination_cell_loc_strings):
            # Update formula of cell, then set contents
            self.set_cell_contents(
                to_sheet,
                cell_loc_str,
                destination_cell_contents[i], False)
        self._notify_cells_helper(self.updated_cells)

            
    def sort_region(self, sheet_name: str, start_location: str, end_location: str, sort_cols: List[int]):
            # Sort the specified region of a spreadsheet with a stable sort, using
            # the specified columns for the comparison.
            #
            # The sheet name match is case-insensitive; the text must match but the
            # case does not have to.
            #
            # The start_location and end_location specify the corners of an area of
            # cells in the sheet to be sorted.  Both corners are included in the
            # area being sorted; for example, sorting the region including cells B3
            # to J12 would be done by specifying start_location="B3" and
            # end_location="J12".
            #
            # The start_location value does not necessarily have to be the top left
            # corner of the area to sort, nor does the end_location value have to be
            # the bottom right corner of the area; they are simply two corners of
            # the area to sort.
            #
            # The sort_cols argument specifies one or more columns to sort on.  Each
            # element in the list is the one-based index of a column in the region,
            # with 1 being the leftmost column in the region.  A column's index in
            # this list may be positive to sort in ascending order, or negative to
            # sort in descending order.  For example, to sort the region B3..J12 on
            # the first two columns, but with the second column in descending order,
            # one would specify sort_cols=[1, -2].
            #
            # The sorting implementation is a stable sort:  if two rows compare as
            # "equal" based on the sorting columns, then they will appear in the
            # final result in the same order as they are at the start.
            #
            # If multiple columns are specified, the behavior is as one would
            # expect:  the rows are ordered on the first column indicated in
            # sort_cols; when multiple rows have the same value for the first
            # column, they are then ordered on the second column indicated in
            # sort_cols; and so forth.
            #
            # No column may be specified twice in sort_cols; e.g. [1, 2, 1] or
            # [2, -2] are both invalid specifications.
            #
            # The sort_cols list may not be empty.  No index may be 0, or refer
            # beyond the right side of the region to be sorted.
            #
            # If the specified sheet name is not found, a KeyError is raised.
            # If any cell location is invalid, a ValueError is raised.
            # If the sort_cols list is invalid in any way, a ValueError is raised.

            # Check if the sheet_name does not exist
            if sheet_name.lower() not in self.sheet_to_idx.keys():
                raise KeyError("Invalid sheet name")

            # Check if target area exents out of bounds
            
            start_location = self.check_valid_location_str(sheet_name, start_location)
            end_location = self.check_valid_location_str(sheet_name, end_location)
            
            # Find cells to copy
            # Get start location index (minimum location) and end location index
            # (max location)
            loc1 = self.extract_row_col_from_loc(start_location)
            loc2 = self.extract_row_col_from_loc(end_location)
            start_loc = (min(loc1[0], loc2[0]), min(loc1[1], loc2[1]))
            end_loc = (max(loc1[0], loc2[0]), max(loc1[1], loc2[1]))

            start_loc = (start_loc[0], self.column_index_from_string(start_loc[1]))
            end_loc = (end_loc[0], self.column_index_from_string(end_loc[1]))

            copy_row_distance = end_loc[0] - start_loc[0]
            copy_column_distance = end_loc[1] - start_loc[1]

            # The sort_cols list may not be empty.
            if len(sort_cols) == 0:
                raise ValueError("sort_cols list cannot be empty")
            
            # All elements in sort_cols must be non-zero integers and
            # must not extend outside bounds of target area
            for col_ind in sort_cols:
                if not isinstance(col_ind, int):
                    raise ValueError("sort_cols takes list of integers only")
                elif col_ind == 0:
                    raise ValueError("cannot use 0 as an index")
                
                if abs(col_ind) > end_loc[1]:
                    raise ValueError("sort_cols contains index outside specified location")
                
            # No column may be specified twice
            abs_sort_cols = [abs(col_ind) for col_ind in sort_cols]
            abs_sort_cols_set = set(abs_sort_cols)
            if len(abs_sort_cols) != len(abs_sort_cols_set):
                raise ValueError("sort_cols may not contain duplicate indexes")
            

            rows_of_cell_locs_to_sort = []
            rows_of_cells_to_sort = []
            rows_of_original_cell_values = []
            sheet = self.get_sheet(sheet_name)

            for i in range(start_loc[0], start_loc[0] + copy_row_distance + 1):
                row_locs = []
                row_cells = []
                row_values = []
                for j in range(start_loc[1], start_loc[1] + copy_column_distance + 1):
                    cell_loc_str = self.cell_loc_str_from_indexes((i, j))
                    row_locs.append(cell_loc_str)
                    if cell_loc_str in sheet.get_cells().keys():
                        row_cells.append(sheet.get_cell(cell_loc_str))
                        row_values.append(sheet.get_cell(cell_loc_str).get_value())
                    else:
                        row_cells.append(Cell(sheet_name, (i,j), ""))
                        row_values.append(None)
                    
                
                rows_of_cell_locs_to_sort.append(row_locs)
                rows_of_cells_to_sort.append(row_cells)
                rows_of_original_cell_values.append(row_values)


            # Store rows as rowAdapterObjects
            row_adapter_objs = []
            for index, value_row in enumerate(rows_of_original_cell_values):
                row_adapter_objs.append(rowAdapterObject(index, value_row, sort_cols))

            # rowAdapterObjects sorted according to rules defined in
            # in sorter.py
            sorted_row_adapter_objs = sorted(row_adapter_objs)

            for sorted_row_ind, row_adapter_obj in enumerate(sorted_row_adapter_objs):
                original_ind = row_adapter_obj.row_ind

                row_shift = sorted_row_ind - original_ind

                # Shift of zero means no switch in ordering:
                if row_shift == 0: 
                    continue

                row_cells = rows_of_cells_to_sort[original_ind]



                updated_row_cell_contents = []
                for cell in row_cells:
                    new_contents = self.update_formula(cell, sheet_name, sheet_name, (0, row_shift))
                    updated_row_cell_contents.append(new_contents)

                row_cell_locs = rows_of_cell_locs_to_sort[sorted_row_ind]
                for i, cell_loc in enumerate(row_cell_locs):
                # Update formula of cell, then set contents
                    self.set_cell_contents(
                        sheet_name,
                        cell_loc,
                        updated_row_cell_contents[i], False)
                    
            self._notify_cells_helper(self.updated_cells)

                    

    ##################################################################################
    ################################ PRIVATE FUNCTIONS ###############################
    ##################################################################################
    def _update_formula(self, cell, sheet_name, new_sheet_name, move):
        # Helper function to rewrite the formula
        contents = cell.get_contents()
        # If the contents exist and is a formula
        if contents is not None and contents[0] == "=":
            try:
                tree = self.PARSER.parse(cell.get_contents())
                # Add outside quotes to the new sheet name if it needs it (has a
                # space)
                if ' ' in new_sheet_name:
                    new_sheet_name = "'" + new_sheet_name + "'"
                fe = FormulaEditor(self, sheet_name, new_sheet_name, move)
                parsed = fe.transform(tree)
                # Add the formula sign back to the contents
                return "=" + parsed.children[0]
            except (exceptions.LarkError, exceptions.UnexpectedCharacters) as e:
                return cell.get_contents()
        return contents
    
    def _new_default_sheet_num(self) -> int:
        # Get the new default num for sheet_name
        return next(
            filterfalse(
                self.default_sheet_nums.__contains__,
                count(1)))
    
    def _update_cells(self, update_cells, sheet):
        # Update all cells in the list update_cells for del_sheet
        sheet_name = sheet.get_name()

        # Iterate through all of the cells to be updated
        for loc in update_cells.keys():
            # Get the cell corresponding to the location
            cell = sheet.get_cell(loc)

            # Get the children and parents
            children, parents = cell.get_children_cells(), cell.get_parent_cells()

            # Per child, update the cell contents
            for child_loc in children:
                try:
                    child_sheet = self.get_sheet(child_loc[0])
                    child = child_sheet.get_cell(child_loc[1])
                    child_sheet_name, child_loc, contents = child.get_info()
                    self.set_cell_contents(
                        child_sheet_name, child_loc, contents, False)
                except (KeyError, ValueError):
                    pass

            # Per parent, update the cell contents
            for parent_loc in parents:
                try:
                    parent_sheet = self.get_sheet(parent_loc[0])
                    parent = parent_sheet.get_cell(parent_loc[1])
                    parent_sheet_name, parent_loc, contents = parent.get_info()

                    # Remove the current parent from the parent_cells list of the all children
                    # Update cell contents
                    c_parent = parent.get_children_cells()
                    c_parent.remove((sheet_name.lower(), loc.upper()))
                    parent.set_children_cells(c_parent)
                    self.set_cell_contents(
                        parent_sheet_name, parent_loc, contents, False)
                except (KeyError, ValueError):
                    pass
        self._notify_cells_helper(self.updated_cells)


    def _column_index_from_string(self, col):
        # get the number corresponding to the column string
        col = str(col)
        power = len(col) - 1
        index = 0
        for c in col.upper():
            val = ord(c)
            if  val < 91 and val > 64:
                index += (val - 64) * 26 ** power
            else:
                raise ValueError
            power -= 1
        return index

    def _cell_loc_str_from_indexes(self, loc):
        # Get the string location from a tuple of integers
        row, col = loc
        if row < 1 or col < 1 or col > 475254 or row > 9999:
            raise ValueError("Indexes provided are out of bounds")

        col_str = ""
        while col > 0:
            col, remainder = divmod(col - 1, 26)
            col_str = chr(65 + remainder) + col_str

        cell_loc_str = col_str + str(row)
        return cell_loc_str

    def _extract_row_col_from_loc(self, location):
        # Extract the row and col from a string location
        loc = ["".join(x)
               for _, x in itertools.groupby(location, key=str.isdigit)]
        return int(loc[1]), loc[0].upper()

    def _check_valid_location_str(self, sheet_name, location):
        # Find the sheet from dict
        sheet_loc = self.sheet_to_idx[sheet_name.lower()]
        sheet = self.sheets[sheet_loc]

        # Get the corresponding row and col from the location
        row, col = self._extract_row_col_from_loc(location)

        # Check loc does not exceed the max sheet size
        sheet_size = sheet.get_max_size()

        # Check if the row and col exceed the size of the sheet
        if sheet_size[0] < self._column_index_from_string(
                col) or sheet_size[1] < row:
            raise ValueError(f"{location} is an invalid location")

        return location

    def _check_loc(self, sheet_name: str, location):
        # Check if the given (sheet_name, location) tuple is valid
        # Check for the sheet_name
        if sheet_name.lower() not in list(self.sheet_to_idx.keys()):
            raise KeyError(f"{sheet_name} is not a valid sheet!")

        # Find the sheet from dict
        sheet_loc = self.sheet_to_idx[sheet_name.lower()]
        sheet = self.sheets[sheet_loc]

        # Extract location from input
        # loc has the form [column, row]
        location = self._check_valid_location_str(sheet_name, location)

        # Return the sheet and the location tuple
        return sheet, location.upper()
    
    def _check_circular_dependency(self, sheet_name, loc):
        # Check for circular dependencies in the newly set cell by iterating through children
        # until the first cell visited is encountered again
        circular_dependency_detected = False
        visited_children = []
        queue_of_cells = [(sheet_name, loc)]

        while len(queue_of_cells) > 0:
            curr_sheet_name, curr_cell_loc = queue_of_cells.pop(0)
            curr_sheet_idx = self.sheet_to_idx[curr_sheet_name.lower()]
            try:
                curr_cell = self.sheets[curr_sheet_idx].get_cell(curr_cell_loc)
                # If the current cell is an error, go to the next cell in the
                # queue
                if isinstance(curr_cell.get_value(), CellError) and curr_cell.get_value(
                ).get_type() == CellErrorType.CIRCULAR_REFERENCE:
                    continue

                # If we have revisited the first cell we visited, then we have
                # a circular dependency
                if len(
                        visited_children) > 0 and curr_cell == visited_children[0]:
                    circular_dependency_detected = True
                    break

                # Add the current cell to the list of visited children
                visited_children.append(curr_cell)

                # Add children of current cell to queue_of_cells
                for c in curr_cell.get_children_cells():
                    if c not in queue_of_cells:
                        queue_of_cells.append(c)

            except KeyError:
                continue

        return circular_dependency_detected, visited_children
    
    def _update_children_cells(self, new_cell, sheet_name, loc):
        # Update the children_cells of the parent cells
        for (p_sheet_name, p_loc) in new_cell.get_parent_cells():
            try:
                if p_sheet_name.lower()[
                        0] == "'" and p_sheet_name.lower()[-1] == "'":
                    p_sheet_name = p_sheet_name[1:-1]
                # Check if the cell exists. If it does, add the child to it
                parent_cell = self.sheets[self.sheet_to_idx[p_sheet_name.lower()]].get_cell(
                    p_loc.upper())
                if (sheet_name.lower(), loc.upper()) not in parent_cell.get_children_cells():
                    parent_cell.append_child_cell((sheet_name.lower(), loc.upper()))
            except KeyError:
                # If sheet name doesnt exist or cell doesn't exist:
                # Create a new cell with None contents, updating parent cells accordingly
                # Put this cell into floating cells
                if p_sheet_name.lower() not in self.sheet_to_idx:
                    # Sheet doesn't exist
                    new_cell.set_value(
                        CellError(
                            CellErrorType.BAD_REFERENCE,
                            detail="Sheet/cell does not exist!"))
                    parent_cell = Cell(p_sheet_name, p_loc, "")
                    parent_cell.append_child_cell((sheet_name.lower(), loc.upper()))
                    self.floating_cells[p_sheet_name.lower()].append(
                        parent_cell)
                # Sheet exists, cell doesn't exist:
                # Create a new cell with None contents, updating parent cells
                # accordingly
                else:
                    parent_sheet = self.sheets[self.sheet_to_idx[p_sheet_name.lower(
                    )]]
                    parent_cell = Cell(p_sheet_name, p_loc, "")
                    parent_cell.append_child_cell((sheet_name.lower(), loc.upper()))
                    parent_sheet.set_cell(p_loc, parent_cell)

    def _build_dependency_graph(self, start_cell, graph):
        # Build dependency_graph rooted at start_cell
        stack = [start_cell]
        visited = set()
        while stack:
            current_cell = stack.pop()
            for child_cell_tuple in current_cell.get_children_cells():
                child_cell_sheet_name = child_cell_tuple[0]
                child_cell_location = child_cell_tuple[1]
                child_cell_sheet_idx = self.sheet_to_idx[child_cell_sheet_name.lower(
                )]
                child_cell = self.sheets[child_cell_sheet_idx].get_cell(
                    child_cell_location)

                if child_cell not in visited:
                    visited.add(child_cell)
                    stack.append(child_cell)
                graph.add_edge(current_cell, child_cell)
        return graph
    
    def _notify_cells_helper(self, changed_cells):
        # Helper function to notifying cells
        # Applies each notification function stored to the list of cells

        changed_cells = list(set(changed_cells))
        for f in self.notification_functions:
            try:
                f(self, changed_cells)
            except BaseException:
                pass

        # Reset the updated cells
        self.updated_cells = []

    def _update_bad_ref(self, new_sheet_name):
        # Recompute the values of any cells that reference the new copied sheet
        # to_remove is a temporary list of cells that need to be removed from
        # the bad reference cells
        to_remove = []
        if new_sheet_name[0] == "'" and new_sheet_name[-1] == "'":
            new_sheet_name = new_sheet_name[1:-1]

        # For each cell in bad_ref_cells
        for (sheet_name, loc) in self.bad_ref_cells.keys():
            cell = self.bad_ref_cells[(sheet_name, loc)]
            for (parent_sheet_name, _) in cell.get_parent_cells():
                # If the cell has a parent that has a sheet_name equal to the
                # new_sheet_name, then update the cell again
                if parent_sheet_name.lower() == new_sheet_name.lower():
                    cell_sheet_name, cell_loc, cell_contents = cell.get_info()
                    self.set_cell_contents(cell_sheet_name, cell_loc, cell_contents)
                    to_remove.append((sheet_name, loc))

        # Remove all of the cells that were updated
        for (sheet_name, loc) in to_remove:
            self.bad_ref_cells.pop((sheet_name, loc))

    def _copy_and_move_helper(
            self,
            sheet_name: str,
            start_location: str,
            end_location: str,
            to_location: str,
            to_sheet: Optional[str] = None) -> None:
        # Since the only difference between copy and move is that the copied cells are
        # "deleted" by the move function. Thus, the code that obtains the destination cells
        # can be used by both functions. Returns three lists:
        # 1. A list of the Cell objects that are to be copied
        # 2. A list of strings with the destination cell locations
        # 3. A list of content-strings to be pasted into the destination cells

        # Check if the sheet_name does not exist
        if sheet_name.lower() not in self.sheet_to_idx.keys():
            raise KeyError("Invalid sheet name")

        # Check if target area exents out of bounds
        if to_sheet is None:
            to_sheet = sheet_name
            to_location = self._check_valid_location_str(
                sheet_name, to_location)
        else:
            to_location = self._check_valid_location_str(to_sheet, to_location)

        # Find cells to copy
        # Get start location index (minimum location) and end location index
        # (max location)
        loc1 = self._extract_row_col_from_loc(start_location)
        loc2 = self._extract_row_col_from_loc(end_location)
        start_loc = (min(loc1[0], loc2[0]), min(loc1[1], loc2[1]))
        end_loc = (max(loc1[0], loc2[0]), max(loc1[1], loc2[1]))

        start_loc = (start_loc[0], self._column_index_from_string(start_loc[1]))
        end_loc = (end_loc[0], self._column_index_from_string(end_loc[1]))

        copy_row_distance = end_loc[0] - start_loc[0]
        copy_column_distance = end_loc[1] - start_loc[1]

        # Get sheet, and the cells to be copied
        sheet = self.get_sheet(sheet_name)

        # Locs stored in numeric (row, col) form (i.e loc = (3,1) for cell "A3"
        cell_locs_to_copy = []
        cells_to_copy = []

        for i in range(start_loc[0], start_loc[0] + copy_row_distance + 1):
            for j in range(
                    start_loc[1],
                    start_loc[1] +
                    copy_column_distance +
                    1):
                cell_locs_to_copy.append((i, j))
                cell_loc_str = self._cell_loc_str_from_indexes((i, j))
                if cell_loc_str in sheet.get_cells().keys():
                    cells_to_copy.append(sheet.get_cell(cell_loc_str))
                else:
                    cells_to_copy.append(Cell(sheet_name, (i,j), ""))

        # Find the cells set to receive copied cells
        target_loc = self._extract_row_col_from_loc(to_location)
        target_loc = (
            target_loc[0],
            self._column_index_from_string(
                target_loc[1]))
        shift_row_distance = target_loc[0] - start_loc[0]
        shift_column_distance = target_loc[1] - start_loc[1]
        shift = (shift_column_distance, shift_row_distance)

        # Locs for cells that copied cells are pasted in
        destination_cell_locs = []
        for cell_loc in cell_locs_to_copy:
            destination_cell_locs.append((cell_loc[0] + shift_row_distance,
                                          cell_loc[1] + shift_column_distance))

        # Check if destination cells extends outside valid bounds
        # cell_loc_str_from_indexes() raises a value error if passed a loc with
        # indexes that are out of bounds
        destination_cell_loc_strings = []
        for cell_loc in destination_cell_locs:
            try:
                destination_cell_loc_strings.append(
                    self._cell_loc_str_from_indexes(cell_loc))
            except ValueError:
                raise ValueError(
                    "Destination cell not within permitted bounds")

        # Get the contents of cells that were originally copied
        destination_cell_contents = []
        for cell in cells_to_copy:
            contents = self._update_formula(cell, sheet_name, to_sheet, shift)
            destination_cell_contents.append(contents)

        return cells_to_copy, destination_cell_loc_strings, destination_cell_contents, shift
    

    def _clear_cell(self, sheet_name: str, location: str):
        # Clear the contents of a cell

        # Check if the cell is valid
        sheet, loc = self.check_loc(sheet_name, location.upper())
        if loc in sheet.get_cells().keys():
            cell = sheet.get_cell(loc)

            # Update the children cells of parents of the current cell
            for parent_loc in cell.get_parent_cells():
                parent_sheet = self.get_sheet(parent_loc[0])
                parent = parent_sheet.get_cell(parent_loc[1])
                parent.set_children_cells(
                    parent.get_children_cells().remove((sheet_name.lower(), loc.upper())))

            # Reset the values of the cell
            cell.set_parent_cells([])
            cell.set_content(None)
            cell.set_value(None)
            cell.set_formula_cell_flag = False

        sheet.update_sheet_extent()