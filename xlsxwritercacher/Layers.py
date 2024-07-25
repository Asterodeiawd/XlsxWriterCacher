from xlsxwriter.worksheet import xl_cell_to_rowcol

from .CacheLayerBase import CacheLayer, LayerType


class DataCacheLayer(CacheLayer):
    def __init__(self):
        super().__init__(LayerType.DATA)

    def write_at(self, row, col, value, value_type=str):
        if row < 0 or col < 0:
            raise IndexError('row and col must be greater than 0')

        self.cache_at_cell(row, col, {"type": value_type, "data": value})

    def write_range(
        self, start_row, start_col, end_row, end_col, value, value_type=str
    ):
        """
        Params:
            row: 0-based
            col: 0-based

            end_row: not included
            end_col: not included
        """
        self.cache_at_range(
            start_row, start_col, end_row, end_col, {"type": value_type, "data": value}
        )


class FormatCacheLayer(CacheLayer):
    def __init__(self):
        super().__init__(LayerType.STYLE)

    def merge_layer(self, other):
        result = FormatCacheLayer()

        this_keys, other_keys = set(self._data.keys()), set(other._data.keys())

        for key in this_keys.symmetric_difference(other_keys):
            row, col = xl_cell_to_rowcol(key)
            result.cache_at_cell(row, col, self._data.get(key, other._data.get(key)))

        for key in this_keys.intersection(other_keys):
            row, col = xl_cell_to_rowcol(key)
            result.cache_at_cell(row, col, {**self._data[key], **other._data[key]})

        return result

    def write_at(self, row, col, props):
        """
        props values should be plain! check later

        Params:
            row: 0-based
            col: 0-based
        """

        if row < 0 or col < 0:
            raise IndexError('row and col must be greater than 0')

        self.cache_at_cell(row, col, props)

    def write_range(self, start_row, start_col, end_row, end_col, props):
        """
        Params:
            row: 0-based
            col: 0-based

            end_row: not included
            end_col: not included
        """
        self.cache_at_range(start_row, start_col, end_row, end_col, props)
