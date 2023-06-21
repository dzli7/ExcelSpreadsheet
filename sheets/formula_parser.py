import itertools
import re
from lark import Lark, Tree, Token
from lark.visitors import Transformer,  Interpreter
from decimal import Decimal, DecimalException
from .cellerror import CellError
from .cellerrortype import *
from .helper import compare_helper, column_index_from_string
from .functions import args_to_bool, args_to_num

# Illegal decimal values
_ILLEGAL_VALUES = [str(Decimal("Nan")), str(Decimal("-Nan")),
                   str(Decimal("Infinity")), str(Decimal("-Infinity"))]

# Map between the literal string error and the CellErrorType
_ERROR_TYPES = {"#ERROR!": CellErrorType.PARSE_ERROR,
                "#CIRCREF!": CellErrorType.CIRCULAR_REFERENCE,
                "#REF!": CellErrorType.BAD_REFERENCE,
                "#NAME?": CellErrorType.BAD_NAME,
                "#VALUE!": CellErrorType.TYPE_ERROR,
                "#DIV/0!": CellErrorType.DIVIDE_BY_ZERO
                }


class FormulaParser(Interpreter):
    # Parser to parse input formulas
    def __init__(self, wb):
        self._wb = wb
        self._default_sheet = ""
        # List of parent cells to the cell of the formula being parsed
        self._parent_cells = set()
        self._functions = wb.get_functions()

    def new_parsing(self, sheet_name, wb):
        self._default_sheet = sheet_name
        self._parent_cells = set()

    def get_parent_cells(self):
        return self._parent_cells

    def cell_range(self, tree):
        children = tree.children
        if len(children) == 1:
            sheet = self._default_sheet
            locs= children[0]
        else:
            sheet = children[0]
            locs = children[1]

        if sheet.lower()[0] == "'" and sheet.lower()[-1] == "'":
            sheet = sheet[1:-1]
        locs = locs.replace("$", "")
        (loc1, loc2) = locs.split(":")

        loc1 = self.extract_row_col_from_loc(loc1)
        loc2 = self.extract_row_col_from_loc(loc2)
        start_loc = (min(loc1[0], loc2[0]), min(loc1[1], loc2[1]))
        end_loc = (max(loc1[0], loc2[0]), max(loc1[1], loc2[1]))

        start_loc = (start_loc[0], self.column_index_from_string(start_loc[1]))
        end_loc = (end_loc[0], self.column_index_from_string(end_loc[1]))

        copy_row_distance = end_loc[0] - start_loc[0]
        copy_column_distance = end_loc[1] - start_loc[1]
        
        cells_to_visit = []
        for i in range(start_loc[0], start_loc[0] + copy_row_distance + 1):
            for j in range(start_loc[1],start_loc[1] + copy_column_distance +1):
                cell_loc = self.cell_loc_str_from_indexes((i,j))
                t = Tree(Token('RULE', 'cell'), [Token('SHEET_NAME', sheet), Token('CELLREF', cell_loc)])
                cells_to_visit.append(t)

        return cells_to_visit

    def add_expr(self, tree):
        l_val, op, r_val = self.visit_children(tree)

        if isinstance(l_val, Tree):
            l_val = self.visit_children(l_val)
        if isinstance(r_val, Tree):
            r_val = self.visit_children(r_val)

        # If either the left of right are CellError, continue to return CellError
        if isinstance(l_val, CellError):
            return l_val
        if isinstance(r_val, CellError):
            return r_val

        # Check if there are any literal errors in the left or right
        if str(r_val) in _ERROR_TYPES.keys():
            return CellError(_ERROR_TYPES[str(r_val)], detail=f"{str(r_val)}")
        if str(l_val) in _ERROR_TYPES.keys():
            return CellError(_ERROR_TYPES[str(l_val)], detail=f"{str(l_val)}")

        # Make the left and right values Decimals
        try:
            l_val = Decimal(l_val) if l_val is not None else Decimal(0)
            r_val = Decimal(r_val) if r_val is not None else Decimal(0)

            # If either value is Nan or Inf, raise ValueError
            if str(l_val) in _ILLEGAL_VALUES or str(r_val) in _ILLEGAL_VALUES:
                raise ValueError("Cannot use Inf/Nan to compute!")
        # If either value cannot be converted to decimal or inf/nan is used,
        # raise CellError with Type_Error
        except (DecimalException, ValueError):
            return CellError(CellErrorType.TYPE_ERROR, detail="+ or - with illegal values")

        # Return the sum or difference of the arguments
        value = l_val + r_val if op == "+" else l_val - r_val
        return Decimal(value)

    def mul_expr(self, tree):
        l_val, op, r_val = self.visit_children(tree)
        if isinstance(l_val, Tree):
            l_val = self.visit_children(l_val)
        if isinstance(r_val, Tree):
            r_val = self.visit_children(r_val)
        # If either the left of right are CellError, continue to return
        # CellError
        if isinstance(l_val, CellError):
            return l_val
        if isinstance(r_val, CellError):
            return r_val

        # Check if there are any literal errors in the left or right
        if str(r_val) in _ERROR_TYPES.keys():
            return CellError(_ERROR_TYPES[str(r_val)], detail=f"{str(r_val)}")
        if str(l_val) in _ERROR_TYPES.keys():
            return CellError(_ERROR_TYPES[str(l_val)], detail=f"{str(l_val)}")

        # Make the left and right values Decimals
        try:
            l_val = Decimal(l_val) if l_val is not None else Decimal(0)
            r_val = Decimal(r_val) if r_val is not None else Decimal(0)

            # If either value is Nan or Inf, raise ValueError
            if str(l_val) in _ILLEGAL_VALUES or str(r_val) in _ILLEGAL_VALUES:
                raise ValueError("Cannot use Inf/Nan to compute!")
        # If either value cannot be converted to decimal or inf/nan is used,
        # raise CellError with Type_Error
        except (DecimalException, ValueError):
            return CellError(CellErrorType.TYPE_ERROR, detail="+ or - with illegal values")

        if r_val == 0 and op == '/':
            return CellError(CellErrorType.DIVIDE_BY_ZERO, detail="can't divide by 0")

        # Return the sum or difference of the arguments
        value = l_val * r_val if op == "*" else l_val / r_val
        return Decimal(value)

    def concat_expr(self, tree):
        l_val, r_val = self.visit_children(tree)
        # If either the left of right are CellError, continue to return
        # CellError
        if isinstance(l_val, CellError):
            return l_val
        if isinstance(r_val, CellError):
            return r_val

        # Make the arguments strings
        l_val = str(l_val) if l_val is not None else ""
        r_val = str(r_val) if r_val is not None else ""

        false_pattern = re.compile("[Ff][Aa][Ll][Ss][Ee]")
        true_pattern = re.compile("[Tt][Rr][Uu][Ee]")
        upper_true = "TRUE"
        upper_false = "FALSE"

        if re.match(false_pattern, l_val) is not None:
            l_val = re.sub(false_pattern, upper_false, l_val)

        if re.match(false_pattern, r_val) is not None:
            r_val = re.sub(false_pattern, upper_false, r_val)
        
        if re.match(true_pattern, l_val) is not None:
            l_val = re.sub(true_pattern, upper_true, l_val)

        if re.match(true_pattern, r_val) is not None:
            r_val = re.sub(true_pattern, upper_true, r_val)

        # regex match l_val/r_val with true/false
        # Concatenate string
        value = l_val + r_val
        return str(value)
    
    def compare_expr(self, tree):
        l_val, op, r_val = self.visit_children(tree)
        if isinstance(l_val, Tree):
            l_val = self.visit_children(l_val)
        if isinstance(r_val, Tree):
            r_val = self.visit_children(r_val)
        # If either the left of right are CellError, continue to return
        # CellError
        if isinstance(l_val, CellError):
            return l_val
        if isinstance(r_val, CellError):
            return r_val

        # Check if there are literal errors in the left or right
        if str(r_val) in _ERROR_TYPES.keys():
            return CellError(_ERROR_TYPES[str(r_val)], detail=f"{str(r_val)}")
        if str(l_val) in _ERROR_TYPES.keys():
            return CellError(_ERROR_TYPES[str(l_val)], detail=f"{str(l_val)}")
        
        return compare_helper(l_val, r_val, op)

    def unary_op(self, tree):
        op, val = self.visit_children(tree)

        if isinstance(val, CellError):
            return val
        # Try converting it to decimal
        try:
            val = Decimal(val) if val is not None else Decimal(0)
            # If either value is Nan or Inf, raise ValueError
            if str(val) in _ILLEGAL_VALUES:
                raise ValueError("Cannot use Inf/Nan to compute!")

        # If either value cannot be converted to decimal or inf/nan is used,
        # raise CellError with Type_Error
        except (DecimalException, ValueError):
            return CellError(CellErrorType.TYPE_ERROR, detail="unary_op with illegal values")

        # Compute the unary op
        value = -1 * val if op == '-' else val
        return value

    def function_expr(self, tree):
        func = tree.children[0].upper()
        arg_list = tree.children[1]
        if func == "IF":
            if isinstance(arg_list, Tree):
                return self.IF(arg_list)
            else:
                return self.IF(arg_list.children)
        elif func == "IFERROR":
            if isinstance(arg_list, Tree):
                return self.IFERROR(arg_list)
            else:
                return self.IFERROR(arg_list.children)
        elif func == "CHOOSE":
            return self.CHOOSE(arg_list.children)
        elif func == "INDIRECT":
            return self.INDIRECT(arg_list.children)
        elif func == "VLOOKUP":
            return self.VLOOKUP(arg_list.children)
        elif func == "HLOOKUP":
            return self.HLOOKUP(arg_list.children)
        else:
            if arg_list.data == "arg_list":
                try:
                    if len(arg_list.children) == 1:
                        args = [self.visit(arg_list)]
                    else:
                        args = self.visit_children(arg_list)
                    return self._functions[func.upper()](*args)
                except (ValueError, KeyError) as e:
                    if isinstance(e, KeyError):
                        return CellError(CellErrorType.BAD_NAME, detail = e)
                    else: 
                        return CellError(CellErrorType.TYPE_ERROR, detail = e)
            elif arg_list.data == "cell_range":
                try:
                    args = Tree(Token('RULE', 'arg_list'), self.visit(arg_list))
                    args = self.visit_children(args)
                    return self._functions[func.upper()](*args)
                except (ValueError, KeyError) as e:
                    if isinstance(e, KeyError):
                        return CellError(CellErrorType.BAD_NAME, detail = e)
                    else: 
                        return CellError(CellErrorType.TYPE_ERROR, detail = e)
            else:
                try:
                    args = self.visit(arg_list)
                    if isinstance(args, list):
                        args = [self.visit(a) for a in args]
                    else:
                        args = [args]
                    return self._functions[func.upper()](*args)
                except (ValueError, KeyError) as e:
                    if isinstance(e, KeyError):
                        return CellError(CellErrorType.BAD_NAME, detail = e)
                    else: 
                        return CellError(CellErrorType.TYPE_ERROR, detail = e)

        
    def VLOOKUP(self, arg_list):
        if len(arg_list) != 3:
            return CellError(CellErrorType.TYPE_ERROR, 
                             detail="HLOOKUP takes exactly three arguments")
        
        key = self.visit(arg_list[0])

        try:
            idx = int(self.visit(arg_list[2]))
        except:
            return CellError(CellErrorType.TYPE_ERROR,
                             detail=f"Could not convert {arg_list[2].children[0]} to number")

        cell_range = arg_list[1].children
        if len(cell_range) == 1:
            sheet = self._default_sheet
            locs= cell_range[0]
        else:
            sheet = cell_range[0]
            locs = cell_range[1]

        if sheet.lower()[0] == "'" and sheet.lower()[-1] == "'":
            sheet = sheet[1:-1]
        locs = locs.replace("$", "")
        (loc1, loc2) = locs.split(":")

        loc1 = self.extract_row_col_from_loc(loc1)
        loc2 = self.extract_row_col_from_loc(loc2)
        if idx > abs(self.column_index_from_string(loc1[1]) \
                     - self.column_index_from_string(loc2[1])) + 1:
            return CellError(CellErrorType.TYPE_ERROR,
                             detail=f"{idx} is not within the provided grid of cells!")

        start_loc = (min(loc1[0], loc2[0]), min(loc1[1], loc2[1]))
        end_loc = (max(loc1[0], loc2[0]), max(loc1[1], loc2[1]))

        start_loc = (start_loc[0], self.column_index_from_string(start_loc[1]))
        end_loc = (end_loc[0], self.column_index_from_string(end_loc[1]))

        col = start_loc[1]
        copy_row_distance = end_loc[0] - start_loc[0]
        copy_col_distance = end_loc[1] - start_loc[1]
        
        cells_to_visit = []
        for i in range(start_loc[1],start_loc[1] + copy_row_distance +1):
            cell_loc = self.cell_loc_str_from_indexes((i,col))
            t = Tree(Token('RULE', 'cell'), [Token('SHEET_NAME', sheet), Token('CELLREF', cell_loc)])
            cells_to_visit.append(t)
        args = Tree(Token('RULE', 'arg_list'), cells_to_visit)
        col_vals = self.visit_children(args)

        row = -1
        for i, val in enumerate(col_vals):
            if val == key:
                row = i + 1

        if row == -1:
            return CellError(CellErrorType.TYPE_ERROR,
                             detail=f"Could not find {key} in provided rows")
        
        cells_to_visit = []
        for i in range(start_loc[0],start_loc[0] + copy_col_distance +1):
            cell_loc = self.cell_loc_str_from_indexes((row,i))
            t = Tree(Token('RULE', 'cell'), [Token('SHEET_NAME', sheet), Token('CELLREF', cell_loc)])
            cells_to_visit.append(t)
        args = Tree(Token('RULE', 'arg_list'), cells_to_visit)
        col_vals = self.visit_children(args)
        return col_vals[idx-1]

        


    def HLOOKUP(self, arg_list):
        if len(arg_list) != 3:
            return CellError(CellErrorType.TYPE_ERROR, 
                             detail="HLOOKUP takes exactly three arguments")
        
        key = self.visit(arg_list[0])

        try:
            idx = int(self.visit(arg_list[2]))
        except:
            return CellError(CellErrorType.TYPE_ERROR,
                             detail=f"Could not convert {arg_list[2].children[0]} to number")

        cell_range = arg_list[1].children
        if len(cell_range) == 1:
            sheet = self._default_sheet
            locs= cell_range[0]
        else:
            sheet = cell_range[0]
            locs = cell_range[1]

        if sheet.lower()[0] == "'" and sheet.lower()[-1] == "'":
            sheet = sheet[1:-1]
        locs = locs.replace("$", "")
        (loc1, loc2) = locs.split(":")

        loc1 = self.extract_row_col_from_loc(loc1)
        loc2 = self.extract_row_col_from_loc(loc2)
        if idx > abs(loc1[0] - loc2[0]) + 1:
            return CellError(CellErrorType.TYPE_ERROR,
                             detail=f"{idx} is not within the provided grid of cells!")

        start_loc = (min(loc1[0], loc2[0]), min(loc1[1], loc2[1]))
        end_loc = (max(loc1[0], loc2[0]), max(loc1[1], loc2[1]))

        start_loc = (start_loc[0], self.column_index_from_string(start_loc[1]))
        end_loc = (end_loc[0], self.column_index_from_string(end_loc[1]))

        row = start_loc[0]
        copy_row_distance = end_loc[0] - start_loc[0]
        copy_col_distance = end_loc[1] - start_loc[1]
        
        cells_to_visit = []
        for i in range(start_loc[0],start_loc[0] + copy_col_distance +1):
            cell_loc = self.cell_loc_str_from_indexes((row,i))
            t = Tree(Token('RULE', 'cell'), [Token('SHEET_NAME', sheet), Token('CELLREF', cell_loc)])
            cells_to_visit.append(t)

        args = Tree(Token('RULE', 'arg_list'), cells_to_visit)
        col_vals = self.visit_children(args)

        col = -1
        for i, val in enumerate(col_vals):
            if val == key:
                col = i+1

        if col == -1:
            return CellError(CellErrorType.TYPE_ERROR,
                             detail=f"Could not find {key} in provided rows")
        
        
        cells_to_visit = []
        for i in range(start_loc[1],start_loc[1] + copy_row_distance +1):
            cell_loc = self.cell_loc_str_from_indexes((i,col))
            t = Tree(Token('RULE', 'cell'), [Token('SHEET_NAME', sheet), Token('CELLREF', cell_loc)])
            cells_to_visit.append(t)
        args = Tree(Token('RULE', 'arg_list'), cells_to_visit)
        col_vals = self.visit_children(args)
        return col_vals[idx-1]

    def IF(self, arg_list):
        if arg_list.data == "cell":
            arg_list = [arg_list]
        else:
            arg_list = arg_list.children
        if len(arg_list) != 2 and len(arg_list) != 3:
            return CellError(CellErrorType.TYPE_ERROR, 
                             detail="IF function takes exactly two or three arguments")
        
        try:
            val = self.visit(arg_list[0])
            cond = args_to_bool([val])[0]
        except ValueError as e:
            return CellError(CellErrorType.TYPE_ERROR,
                             detail=f"Cannot convert {arg_list[0]} to boolean!")

        if cond:
            return self.visit(arg_list[1])
        else:
            if len(arg_list) == 3:
                return self.visit(arg_list[2])
            else:
                return False
        
    def IFERROR(self, arg_list):
        if arg_list.data == "cell":
            arg_list = [arg_list]
        else:
            arg_list = arg_list.children
        
        if len(arg_list) != 1 and len(arg_list) != 2:
            return CellError(CellErrorType.TYPE_ERROR, 
                             detail="IF function takes exactly one or two arguments")
        val = self.visit(arg_list[0])
        cond = isinstance(val, CellError)

        if cond:
            if len(arg_list) == 2:
                return self.visit(arg_list[1])
            else:
                return ""
        else:
            return val

    def CHOOSE(self, arg_list):
        if len(arg_list) < 2:
            return CellError(CellErrorType.TYPE_ERROR, 
                             detail="CHOOSE function takes more than 2 arguments")
        try:
            val = self.visit(arg_list[0])
            idx = args_to_num([val])[0]
        except ValueError as e:
            return CellError(CellErrorType.TYPE_ERROR,
                             detail=f"Cannot convert {arg_list[0]} to numeric!")

        if idx < 1 or idx > len(arg_list) - 1:
            return CellError(CellErrorType.TYPE_ERROR,detail=f"Index out of bounds")
        return self.visit(arg_list[int(idx)])

    def INDIRECT(self, arg_list):
        arg = arg_list[0]
        if len(arg) < 2 or (arg[0] != '"' and arg[-1] != '"'):
            return CellError(CellErrorType.PARSE_ERROR,
                                        detail = "INDIRECT takes one argument surrounded by quotes")
        if len(arg_list) != 1:
            return CellError(CellErrorType.TYPE_ERROR,
                                        detail = "INDIRECT takes one argument")
        arg = arg.replace('"', "")
        if "!" in arg:
            arg = arg.split("!")
        
        if isinstance(arg, list):
            tree = Tree(Token('RULE', 'cell'), [Token('SHEET_NAME', arg[0]), Token('CELLREF', arg[1])])
        else:
            tree = Tree(Token('RULE', 'cell'), [Token('CELLREF', arg)])
        val = self.cell(tree)
        return val

    def number(self, tree):
        return Decimal(tree.children[0].value)

    def string(self, tree):
        s = str(tree.children[0].value)
        if len(s) > 2 and s[0] == '"' and s[-1] == '"':
            return s[1:-1]
        elif s == '""':
            return ""
        elif s[0] == "'":
            return s[1::]
        else:
            return s
        
    def bool(self, tree):
        return False if tree.children[0].lower() == "false" else True

    def error(self, tree):
        return CellError(_ERROR_TYPES[tree.children[0]], detail=tree.children[0].value)
    
    def parens(self, tree):
        return self.visit(tree.children[0])

    def cell(self, tree):
        if len(tree.children) == 2:
            sheet = str(tree.children[0])
            if sheet.lower()[0] == "'" and sheet.lower()[-1] == "'":
                sheet = sheet[1:-1]
            cell = str(tree.children[1])
        else:
            sheet = self._default_sheet
            cell = str(tree.children[0])
        # Replace absolute value sign with just string so we can get
        cell = cell.replace("$", "")

        # Check the location of the referenced cell
        loc = ["".join(x) for _, x in itertools.groupby(cell, key=str.isdigit)]
        loc[0] = loc[0].upper()
        loc[1] = int(loc[1])

        if 475254 < column_index_from_string(loc[0]) or 9999 < loc[1]:
            return CellError(CellErrorType.BAD_REFERENCE, detail=f"{cell} is out of range")
        
        self._parent_cells.add((sheet.lower(), cell.upper()))
        # Append to parent cell list and get the value of the cell
        try:
            val = self._wb.get_cell_value(sheet, cell)
        except (KeyError, ValueError) as e:
            val = CellError(CellErrorType.BAD_REFERENCE, detail=f"{e}")
        # If the value is a literal cell error, return the respective error
        try:
            p = str(val).upper()
            if p in _ERROR_TYPES.keys():
                val = CellError(_ERROR_TYPES[p], detail=f"{p}")
        except (KeyError, AttributeError):
            pass

        return val

    def cell_loc_str_from_indexes(self, loc):
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
    
    def extract_row_col_from_loc(self, location):
        # Extract the row and col from a string location
        loc = ["".join(x)
               for _, x in itertools.groupby(location, key=str.isdigit)]
        return int(loc[1]), loc[0].upper()
    
    def column_index_from_string(self, col):
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