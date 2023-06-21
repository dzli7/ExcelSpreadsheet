import itertools


class Sheet:
    def __init__(self, idx=-1, name=None):
        # Initialize a new empty workbook.
        self._index = idx
        self._name = name
        # Cells in the sheet. Keys are "col row"
        self._cells = {}
        # Current size of sheet
        self._size = (0, 0)
        # Max size in [column, row]
        self._max_size = [475254, 9999]

    def get_max_size(self):
        # Return max size of sheet
        return self._max_size

    def update_sheet_extent(self):
        # Update the sheet extent
        max_col, max_row = 0, 0
        # Iterate through each location
        for cell in self._cells.values():
            # Get the cell in that location
            # cell = self._cells[loc]
            (row, col) = cell.get_extent()
            # If the cell contents is not None, update the max of col and row
            if cell.get_contents():
                max_col = max(col, max_col)
                max_row = max(row, max_row)
        self._size = (max_col, max_row)

    def get_extent(self):
        # Return current size of the sheet
        return self._size

    def get_name(self):
        # Return the name of the sheet
        return self._name

    def set_name(self, name):
        self._name = name

    def get_index(self):
        return self._index

    def set_index(self, idx):
        self._index = idx

    def set_cell(self, loc, cell):
        self._cells[loc] = cell

    def get_cell(self, loc):
        return self._cells[loc]

    def get_cells(self):
        return self._cells

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
