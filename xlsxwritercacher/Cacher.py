from functools import reduce

from .Layers import DataCacheLayer, FormatCacheLayer


class FormatCacheLayerContainer(list):
    def get_projection(self):
        return reduce(lambda l1, l2: l1.merge_layer(l2), self, FormatCacheLayer())

    def add_layer(self, index=None):
        """
        Add a layer in the cache, the later format layer will overwrite
        previous ones if property are equal, otherwise add to format cache.

        Params:
            self: XlsxWriterCache Instance
            order: insert layer into the format cache layer array at the given position

        Returns:
            FormatCacheLayer()

        """

        layer = FormatCacheLayer()

        if index is not None:
            self.insert(index, layer)
        else:
            self.append(layer)

        return layer


class XlsxWriterCacher:
    def __init__(self):
        self._data_layer = DataCacheLayer()
        self._format_layer_container = FormatCacheLayerContainer()

    def add_format_layer(self, index=None):
        return self._format_layer_container.add_layer(index)

    def get_data_layer(self):
        return self._data_layer

    def render(self, workbook, sheet_name):
        worksheet = workbook.get_worksheet_by_name(sheet_name)

        format_layer = self._format_layer_container.get_projection()
        data_layer = self._data_layer

        format_keys, data_keys = set(format_layer._data.keys()), set(
            data_layer._data.keys()
        )

        for key in format_keys.difference(data_keys):
            worksheet.write(key, None, workbook.add_format(format_layer._data[key]))

        for key in data_keys.difference(format_keys):
            # only support write string now
            worksheet.write(key, data_layer._data[key]["data"])

        for key in format_keys.intersection(data_keys):
            worksheet.write(
                key,
                data_layer._data[key]["data"],
                workbook.add_format(format_layer._data[key]),
            )
