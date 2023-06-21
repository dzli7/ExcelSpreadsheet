import itertools
from lark.visitors import Transformer
from lark import Tree
from .cellerrortype import *
from .helper import column_index_from_string


class FormulaEditor(Transformer):
    # Parser to update formulas with new_sheet_name
    def __init__(self, wb, sheet_name, new_sheet_name, move):
        self._wb = wb
        self._default_sheet = sheet_name
        self._new_sheet_name = new_sheet_name
        self._move_cell = move

    def RULE(self, args):
        return args[0]

    def NUMBER(self, token):
        return str(token)

    def STRING(self, token):
        return str(token)

    def BOOL(self, token):
        return str(token)

    def parens(self, args):
        return Tree('STRING', [str("(" + str(args[0].children[0]) + ")")])

    def add_expr(self, args):
        # Return the concatenated string of the arguments
        l_val, op, r_val = args[0].children[-1], args[1], args[-1].children[-1]
        value = str(l_val) + " " + str(op) + " " + str(r_val)
        return Tree('STRING', [f"{(value)}"])

    def mul_expr(self, args):
        # Return the concatenated string of the arguments
        l_val, op, r_val = args[0].children[-1], args[1], args[-1].children[-1]
        value = str(l_val) + " " + str(op) + " " + str(r_val)
        return Tree('STRING', [f"{(value)}"])

    def concat_expr(self, args):
        # Get the left and right strings to concatenate
        l_val, r_val = args[0].children[-1], args[1].children[-1]

        # Concatenate string
        l_val = str(l_val) if l_val is not None else ""
        r_val = str(r_val) if r_val is not None else ""

        value = l_val + r_val
        return Tree('STRING', [f"{(value)}"])

    def unary_op(self, args):
        # Return the concatenated string of the arguments
        op = args[0]
        val = args[1].children[0]
        return Tree('STRING', [f"{(str(op) + str(val))}"])

    def function_expr(self, args):
        s = ""
        for a in args:
            s += a
        return s

    def arg_list(self, args):
        s = ""
        for a in args:
            s += a
        return a

    def CELLREF(self, token):
        return token

    def cell(self, args):
        if self._move_cell:
            return self.cell_move_cell(args)
        a = self.cell_rename_sheet(args)
        return a

    def cell_rename_sheet(self, args):
        # Default sheet of the cell
        sheet = self._default_sheet
        exclamation_point = False
        # Cet the cell and the Sheet if specified
        if len(args) == 1:
            cell = str(args[0])
        else:
            sheet = str(args[0])
            # Quoted sheet names
            if sheet.lower()[0] == "'" and sheet.lower()[-1] == "'":
                sheet = sheet[1:-1]
            cell = str(args[1])
            exclamation_point = True

        # Check if the sheet name is equal to the one that was changed
        # If it is, update the sheet name iwth the new one
        if sheet == self._default_sheet:
            sheet = self._new_sheet_name
            exclamation_point = True
        if exclamation_point:
            return Tree('STRING', [f"{sheet}!{cell}"])
        return Tree('STRING', [f"{cell}"])

    def cell_move_cell(self, args):
        # Default sheet of the cell
        sheet = self._default_sheet
        exclamation_point = False
        # Cet the cell and the Sheet if specified
        if len(args) == 1:
            cell = str(args[0])
        else:
            sheet = str(args[0])
            cell = str(args[1])
            exclamation_point = True
        try:
            row, col = self.helper_extract_row_col_from_loc(cell)
            if cell.count("$") == 0:
                row = row + self._move_cell[1]
                col = column_index_from_string(
                    col) + self._move_cell[0]
                col, row = self.helper_cell_loc_str_from_indexes((row, col))
                cell = col + row
            elif cell.count("$") == 1:
                # col is absolute
                if col[0] == "$":
                    row = row + self._move_cell[1]
                    cell = col + str(row)
                # row is absolute
                else:
                    # Remove $ from col and update it
                    col = column_index_from_string(
                        col.replace("$", "")) + self._move_cell[0]
                    col, row = self.helper_cell_loc_str_from_indexes(
                        (row, col))
                    cell = col + "$" + row
        except (ValueError):
            return Tree("String", ["#REF!"])
            # return Tree("cell_error", [CellError(CellErrorType.BAD_REFERENCE, detail=f"Invalid reference location")])
        # # Check if the sheet name is equal to the one that was changed
        # # If it is, update the sheet name iwth the new one
        # if sheet == self._default_sheet:
        #     sheet = self._new_sheet_name
        #     exclamation_point = True

        if exclamation_point:
            return Tree('STRING', [f"{sheet}!{cell}"])
        return Tree("STRING", [f"{cell}"])

    def helper_extract_row_col_from_loc(self, location):
        loc = ["".join(x)
               for _, x in itertools.groupby(location, key=str.isdigit)]
        return int(loc[1]), loc[0].upper()

    def helper_cell_loc_str_from_indexes(self, loc):
        row, col = loc
        if row < 1 or col < 1 or col > 475254 or row > 9999:
            raise ValueError("Indexes provided are out of bounds")

        col_str = ""
        while col > 0:
            col, remainder = divmod(col - 1, 26)
            col_str = chr(65 + remainder) + col_str

        return col_str, str(row)
