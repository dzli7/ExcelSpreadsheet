from decimal import Decimal
from functools import reduce
from sheets.cellerror import CellError
import sheets
from lark import Token

# Illegal decimal values
_ILLEGAL_VALUES = [Decimal("Nan"), Decimal("-Nan"),
                   Decimal("Infinity"), Decimal("-Infinity")]

def args_to_bool(args):
    '''Convert any non-boolean types in arguments to bool
    so that they may be used by boolean functions'''
    converted_args = []
    for arg in args:
        if isinstance(arg, list):
            arg = arg[0]
        if isinstance(arg, bool):
            converted_args.append(arg)
        else:
            if isinstance(arg, Decimal):
                # bool of any non-zero number returns True
                converted_args.append(bool(arg))
            elif isinstance(arg, str):

                if arg.lower() == "false":
                    converted_args.append(False)
                elif arg.lower() == "true":
                    converted_args.append(True)
                else:
                    raise ValueError(str(arg) + " cannot be converted to type BOOL")
            elif arg is None:
                converted_args.append(False)
            # Argument cannot be converted so raise error
            else:
                raise ValueError(str(arg) + " cannot be converted to type BOOL")
    return converted_args
            
def args_to_str(args):
    '''Convert any non-string types in arguments to string
    so that they may be used by string functions'''
    converted_args = []
    for arg in args:
        if isinstance(arg, list):
            arg = arg[0]
        if isinstance(arg, Token):
            arg = arg.value
            if arg.upper() == "TRUE":
                arg = True
            elif arg.upper() == "FALSE":
                arg = False
        if isinstance(arg, str):
            converted_args.append(arg)
        else:
            if isinstance(arg, Decimal):
                converted_args.append(str(arg))
            elif isinstance(arg, bool):
                if not arg:
                    converted_args.append("FALSE")
                else:
                    converted_args.append("TRUE")
            elif arg is None:
                converted_args.append("")
            # Argument cannot be converted so raise error
            else:
                raise ValueError(str(arg) + " cannot be converted to type string")
    return converted_args

def args_to_num(args):
    '''Convert any non-number types in arguments to number
    so that they may be used by functions that require numbers'''
    converted_args = []
    for arg in args:
        if isinstance(arg, list):
            arg = arg[0]
        if isinstance(arg, Decimal):
            converted_args.append(arg)
        else:
            if isinstance(arg, str):
                try:
                    num = Decimal(arg)
                    if num in _ILLEGAL_VALUES:
                        raise ValueError(str(num) + "is an illegal decimal")
                    else:
                        converted_args.append(num)
                except:
                    raise ValueError(str(num) + " cannot be converted to number")
                
            elif isinstance(arg, bool):
                if not arg:
                    converted_args.append(Decimal(0))
                else:
                    converted_args.append(Decimal(1))
            elif arg is None:
                converted_args.append(Decimal(0))
            # Argument cannot be converted so raise error
            else:
                raise ValueError(str(arg) + " cannot be converted to type number")
    return converted_args
            
def AND(*args):
    if len(args) < 1:
        raise ValueError("AND function takes one or more arguments")
    try:
        args = args_to_bool(args)
    except ValueError as e:
        raise ValueError(e)
    return reduce(lambda arg1, arg2: arg1 and arg2, args)

def OR(*args):
    if len(args) < 1:
        raise ValueError("OR function takes one or more arguments")
    try:
        args = args_to_bool(args)
    except ValueError as e:
        raise ValueError(e)
    return reduce(lambda arg1, arg2: arg1 or arg2, args)

def NOT(*args):
    if len(args) != 1:
        raise ValueError("NOT function takes exactly one argument")
    try:
        args = args_to_bool(args)
    except ValueError as e:
        raise ValueError(e)
    return not args[0]

def XOR(*args):
    if len(args) < 1:
        raise ValueError("XOR function takes one or more arguments")
    try:
        args = args_to_bool(args)
    except ValueError as e:
        raise ValueError(e)
    return reduce(lambda arg1, arg2: arg1 ^ arg2, args)

def EXACT(*args):
    if len(args) != 2:
        raise ValueError("EXACT function takes exactly two argument")
    try:
        args = args_to_str(args)
    except ValueError as e:
        raise ValueError(e)
    return args[0] == args[1]

# def IF(*args):
#     if len(args) != 2 and len(args) != 3:
#         raise ValueError("IF function takes exactly two or three arguments")
#     try:
#         cond = args_to_bool([args[0]])[0]
#     except ValueError as e:
#         raise ValueError(f"Cannot convert {args[0]} to boolean!")

#     if isinstance(cond, CellError):
#         raise ValueError("IF function takes exactly two or three arguments")
#     if cond:
#         return args[1]
#     else:
#         if len(args) == 3:
#             return args[2]
#         else:
#             return False

# def IFERROR(*args):
#     if len(args) != 1 and len(args) != 2:
#         raise ValueError("IFERROR function takes exactly one or two arguments")
    
#     if not isinstance(args[0], CellError):
#         return args[0]
#     else:
#         if len(args) == 2:
#             return args[1]
#         else:
#             return ""
        
# def CHOOSE(*args):
#     if len(args) < 2:
#         raise ValueError("CHOOSE function requires two or more arguments")
#     try:
#         index = args_to_num(args)
#     except ValueError as e:
#         raise ValueError(e)
#     if index < 1 or index > len(args) - 1:
#         raise ValueError("Index out of bounds")
    
#     return args[index]

def ISBLANK(*args):
    if len(args) != 1:
        raise ValueError("ISBLANK function takes exactly one argument")
    args = args[0]
    if args == None:
        return True
    return False

def ISERROR(*args):
    if len(args) != 1:
        raise ValueError("ISERROR function takes exactly one argument")

    return isinstance(args[0], CellError)

def SUM(*args):
    if len(args) < 0:
        raise ValueError("SUM takes at least one argument")
    try:
        args = args_to_num(args)
    except ValueError as e:
        raise ValueError(e)
    return sum(args)

def MIN(*args):
    if len(args) < 0:
        raise ValueError("MIN takes at least one argument")
    try:
        args = args_to_num(args)
    except ValueError as e:
        raise ValueError(e)
    return min(args)

def MAX(*args):
    if len(args) < 0:
        raise ValueError("MAX takes at least one argument")
    try:
        args = args_to_num(args)
    except ValueError as e:
        raise ValueError(e)
    return max(args)

def AVERAGE(*args):
    if len(args) < 0:
        raise ValueError("AVERAGE takes at least one argument")
    try:
        args = args_to_num(args)
    except ValueError as e:
        raise ValueError(e)
    return sum(args) / len(args)

def VERSION(*args):
    if len(args) != 0:
        raise ValueError("VERSION does not take any arguments")
    
    return sheets.version

def functions():
    " Returns a dictionary so functions can be accessed by their name as str"
    " I.E. { str : func() }"
    name_to_function_dict = {}
    name_to_function_dict["AND"] = AND
    name_to_function_dict["OR"] = OR
    name_to_function_dict["NOT"] = NOT
    name_to_function_dict["XOR"] = XOR
    name_to_function_dict["EXACT"] = EXACT
    # name_to_function_dict["IF"] = IF
    # name_to_function_dict["IFERROR"] = IFERROR
    # name_to_function_dict["CHOOSE"] = CHOOSE
    name_to_function_dict["ISBLANK"] = ISBLANK
    name_to_function_dict["ISERROR"] = ISERROR
    name_to_function_dict["VERSION"] = VERSION
    name_to_function_dict["SUM"] = SUM
    name_to_function_dict["MIN"] = MIN
    name_to_function_dict["MAX"] = MAX
    name_to_function_dict["AVERAGE"] = AVERAGE


    return name_to_function_dict








        




