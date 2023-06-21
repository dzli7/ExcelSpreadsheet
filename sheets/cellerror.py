from typing import Optional
from .cellerrortype import CellErrorType


class CellError:
    '''
    This class represents an error value from user input, cell parsing, or
    evaluation.
    '''

    def __init__(self, error_type: CellErrorType, detail: str,
                 exception: Optional[Exception] = None):
        self._error_type = error_type
        self._detail = detail
        self._exception = exception

    def get_type(self) -> CellErrorType:
        ''' The category of the cell error. '''
        return self._error_type

    def get_detail(self) -> str:
        ''' More detail about the cell error. '''
        return self._detail

    def get_exception(self) -> Optional[Exception]:
        '''
        If the cell error was generated from a raised exception, this is the
        exception that was raised.  Otherwise this will be None.
        '''
        return self._exception

    def __str__(self) -> str:
        return f'ERROR[{self._error_type}, "{self._detail}"]'

    def __repr__(self) -> str:
        return self.__str__()
