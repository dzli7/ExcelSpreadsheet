import enum


class CellErrorType(enum.Enum):
    '''
    This enum specifies the kinds of errors that spreadsheet cells can hold.
    '''

    # A formula doesn't parse successfully ("#ERROR!")
    PARSE_ERROR = 1

    # A cell is part of a circular reference ("#CIRCREF!")
    CIRCULAR_REFERENCE = 2

    # A cell-reference is invalid in some way ("#REF!")
    BAD_REFERENCE = 3

    # Unrecognized function name ("#NAME?")
    BAD_NAME = 4

    # A value of the wrong type was encountered during evaluation ("#VALUE!")
    TYPE_ERROR = 5

    # A divide-by-zero was encountered during evaluation ("#DIV/0!")
    DIVIDE_BY_ZERO = 6
