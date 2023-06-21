from decimal import Decimal
from sheets.cellerror import CellError
from sheets.cellerrortype import CellErrorType

# Helper functions

def column_index_from_string(col):
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

def empty_cell_conversion(val):
    if type(val) == str: return ""
    elif type(val) == Decimal: return Decimal("0")
    elif type(val) == bool: return False

def compare_bools(l_val, r_val, op):
    if op in ["=", "=="]:
        return l_val == r_val
    elif op == ">":
        return l_val > r_val
    elif op == "<":
        return l_val < r_val
    elif op == ">=":
        return l_val >= r_val
    elif op == "<=":
        return l_val <= r_val
    elif op in ["<>", "!="]:
        return l_val != r_val

def compare_helper(l_val, r_val, op):
    if l_val is None and r_val is not None:
        l_val = empty_cell_conversion(r_val)
    elif l_val is not None and r_val is None:
        r_val = empty_cell_conversion(l_val)
    elif l_val is None and r_val is None:
        l_val, r_val, = 0, 0

    if type(l_val) == type(r_val):
        if type(l_val) == str:
            l_val, r_val = l_val.lower(), r_val.lower()
        return compare_bools(l_val, r_val, op)
    
    elif type(l_val) == bool and type(r_val) in [str, Decimal]:
        if op in [">=", ">", "!=", "<>"]:
            return True
        return False
    
    elif type(l_val) == str and type(r_val) == Decimal:
        if op in [">=", ">", "!=", "<>"]:
            return True
        return False
    
    elif type(r_val) == bool and type(l_val) in [str, Decimal]:
        if op in ["<=", "<", "!=", "<>"]:
            return True
        return False
    
    elif type(r_val) == str and type(l_val) == Decimal:
        if op in ["<=", "<", "!=", "<>"]:
            return True
        return False

    return False

def column_string_from_index(loc):
    col = loc
    if col < 1 or col > 475254:
        raise ValueError("Indexes provided are out of bounds")

    col_str = ""
    while col > 0:
        col, remainder = divmod(col - 1, 26)
        col_str = chr(65 + remainder) + col_str

    return col_str

def correct_error(error, actual_error):
    # Check for specific error type matches
    if isinstance(error, CellError):
        if error.get_type() == actual_error:
            return True
    return False