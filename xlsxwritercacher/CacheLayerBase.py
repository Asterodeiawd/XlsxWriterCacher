from enum import Enum
from itertools import product

import xlsxwriter


class LayerType(Enum):
    DATA = 0
    STYLE = 1
    MIXED = 2


class CacheLayer:
    def __init__(self, type):
        self._type = type
        self._data = {}

    def cache_at_cell(self, row, col, props, update=False):
        """
        props values should be plain! check later

        Params:
            row: 0-based
            col: 0-based
        """
        if row < 0 or col < 0:
            raise IndexError('row and col must be greater than 0')

        cell = xlsxwriter.worksheet.xl_rowcol_to_cell_fast(
            row, col
        )  # 将行列号转换成A1这样的格式

        old_format = self._data.get(cell, {})

        self._data[cell] = {**old_format, **props}

    def cache_at_range(
        self, start_row, start_col, end_row, end_col, props, update=False
    ):
        """
        Params:
            row: 0-based
            col: 0-based

            end_row: not included
            end_col: not included
        """
        for row, col in product(range(start_row, end_row), range(start_col, end_col)):
            self.cache_at_cell(row, col, props)
