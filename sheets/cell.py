import re
import itertools

from decimal import Decimal, InvalidOperation 


class Cell:
    def __init__(self, sheet_name, loc, contents=None):
        # For cell dependencies
        self._parent_cells = []  # Cells that new cell depends on
        self._children_cells = []  # Cells that depend on new cell
        self._formula_cell_flag = False
        self._contents = contents
        # Check if cell is empty string, and set cell to None type
        if not contents or contents == "" or len(contents) == 0:
            self.contents = None
        else:
            self.contents = contents.strip()
        self._value = None
        self._sheet_name = sheet_name
        self._loc = loc

        loc = ["".join(x) for _,x in itertools.groupby(loc,key=str.isdigit)]
        row = int(loc[1])
        col = self.column_index_from_string(loc[0])
        self._extent = (row, col)

    def get_extent(self):
        return self._extent

    def get_info(self):
        # Return information needed for set cell contents
        return self._sheet_name, self._loc, self._contents

    def set_sheet_name(self, sheet_name):
        self._sheet_name = sheet_name

    def set_content(self, content):
        self._contents = content

    def get_contents(self):
        return self._contents

    def get_parent_cells(self):
        return self._parent_cells

    def set_parent_cells(self, parent_cells):
        self._parent_cells = parent_cells

    def get_children_cells(self):
        return self._children_cells

    def set_children_cells(self, children_cells):
        self._children_cells = children_cells

    def append_child_cell(self, child):
        self._children_cells.append(child)

    def get_formula_cell_flag(self):
        return self._formula_cell_flag

    def set_formula_cell_flag(self, flag):
        self._formula_cell_flag = flag

    def get_value(self):
        return self._value

    def set_value(self, value):
        self._value = value

    def get_loc(self):
        return self._loc

    def parse_cell(self):
        # Parse the contents of a cell
        # If None, return None
        if not self.contents:
            return None

        # If the contents is denoted as a string
        if self.contents[0] == "'":
            self.contents = self.contents[1::]
            # If string can be decimal
            if self.contents != "" and re.match(
                    "^-?\\d*(\\.\\d+)?$", self.contents) is not None:
                return Decimal(self.contents).quantize(Decimal(self.contents))
            # All other strings
            else:
                return self.contents

        # For formulas
        if self.contents[0] == '=':
            self._formula_cell_flag = True
            return self.contents
        self._formula_cell_flag = False

        # For booleans
        if self.contents.lower() == 'false':
            return False

        if self.contents.lower() == 'true':
            return True

        # For literal errors
        if self.contents[0] == '#':
            return self.contents

        # Decimals should be left
        try:
            norm = Decimal(self.contents).normalize()
            _, _, exp = norm.as_tuple()
            return norm if exp <= 0 else norm.quantize(1)
        except InvalidOperation:
            return self.contents
        
    def column_index_from_string(self, col):
            # get the number corresponding to the column string
            col = str(col)
            power = len(col) - 1
            index = 0
            for c in col.upper():
                val = ord(c)
                if val < 91 and val > 64:
                    index += (val - 64) * 26 ** power
                else:
                    raise ValueError
                power -= 1
            return index