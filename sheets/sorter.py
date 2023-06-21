from functools import total_ordering
from decimal import Decimal
from .cellerror import CellError
from .cellerrortype import CellErrorType

@total_ordering
class rowAdapterObject:
    def __init__(self, row_ind, row_values, sort_cols):
        self.row_ind = row_ind
        self.row_values = row_values
        self.sort_cols = sort_cols 

    def _less_than(self, obj1, obj2):
        # Helper function to do less than comparisons on objects of varying types
        # Comparison heirarchy as follows: 
        # None < Error < Decimal < String < Bool

        # Check if either or both objects are None 
        if obj1 is None and obj2 is None:
            return False
        if obj1 is None and obj2 is not None:
            return True
        if obj1 is not None and obj2 is None:
            return False
        
        # Check if either or both objects are Cell Error
        if isinstance(obj1, CellError) and isinstance(obj2, CellError):
            return obj1.get_type().value < obj2.get_type().value
        if isinstance(obj1, CellError) and not isinstance(obj2, CellError):
            return True
        if not isinstance(obj1, CellError) and isinstance(obj2, CellError):
            return False
        
        # Check if objects are same type
        if type(obj1) == type(obj2):
            if isinstance(obj1, str):
                # Case insesitive comparisons for strings
                return obj1.lower() < obj2.lower()
            return obj1 < obj2
        
        if isinstance(obj1, Decimal) and isinstance(obj2, str):
            return True
        if isinstance(obj1, str) and isinstance(obj2, Decimal):
            return False
        
        if (isinstance(obj1, Decimal) or isinstance(obj1, str)) \
            and isinstance(obj2, bool):
            return True
        
        if (isinstance(obj2, Decimal) or isinstance(obj2, str)) \
            and isinstance(obj1, bool):
            return False
        else:
            raise ValueError("Comparison invalid")
        
    def _equals(self, obj1, obj2):
        if isinstance(obj1, CellError) and isinstance(obj2, CellError):
            return obj1.get_type().value == obj2.get_type().value
        if isinstance(obj1, str) and isinstance(obj2, str):
            return obj1.lower() == obj2.lower()
        return obj1 == obj2
    
    def __eq__(self, other): 
        for index in self.sort_cols:
            # Ascending vs descending does not matter for equals
            index = abs(index)
            # Columns 1-indexed
            if not self._equals(self.row_values[index - 1], other.row_values[index - 1]):
                return False
            
        # All possible column comparisons made
        return True
                
    def __lt__(self, other):
        for index in self.sort_cols:
            # Sort ascending or descending dependent on sign of index
            if index > 0:
                obj1 = self.row_values[index - 1]
                obj2 = other.row_values[index - 1]
            else:
                obj1 = other.row_values[(-index) - 1] 
                obj2 = self.row_values[(-index) - 1]

            if self._less_than(obj1, obj2):
                return True
            # Move onto next column comparison if equality found
            elif self._equals(obj1, obj2):
                continue
            else:
                return False

        # If continued until end of loop, all obj equal
        return False
