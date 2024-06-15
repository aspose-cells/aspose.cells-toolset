"""
    Data Conversion
"""

from tkinter import W
from aspose.cells import Workbook
from aspose.cells import Worksheet
from aspose.cells import Cells
from aspose.cells import Cell
from aspose.cells import Range
from aspose.cells import Name
from aspose.cells import CellsHelper
from aspose.cells import CellValueType
from aspose.cells import ProtectionType
from aspose.cells.tables import ListObject
import numpy as np
import pandas as pd


## cells object to python object
def worksheet_to_list(worksheet: Worksheet) -> list:
    """
    Convert Worksheet to list.
    :param Worksheet worksheet:  (required)
    :return list:
    """
    max_row_index = worksheet.cells.max_row
    max_column_index = worksheet.cells.max_column
    table = []
    for row_index in range(0, max_row_index):
        row = []
        for column_index in range(0, max_column_index):
            cur_cell = worksheet.cells.check_cell(row_index, column_index)
            if cur_cell != None:
                row.append(cur_cell.value)
        table.append(row)
    return table


def worksheet_to_tuple(worksheet: Worksheet) -> tuple:
    """
    Convert Worksheet to tuple.
    :param Worksheet worksheet:  (required)
    :return tuple:
    """
    max_row_index = worksheet.cells.max_row
    max_column_index = worksheet.cells.max_column
    table = []
    for row_index in range(0, max_row_index):
        row = []
        for column_index in range(0, max_column_index):
            cur_cell = worksheet.cells.check_cell(row_index, column_index)
            if cur_cell != None:
                row.append(cur_cell.value)
        table.append(row)
    return tuple(table)


def worksheet_to_ndarray(worksheet: Worksheet) -> np.ndarray:
    """
    Convert Worksheet to ndarray.
    :param Worksheet worksheet:  (required)
    :return ndarray:
    """
    ##
    max_row_index = worksheet.cells.max_row
    max_column_index = worksheet.cells.max_column
    # worksheet.cells.get(0,0).type
    # table =np.full(max_row_index, max_column_index)
    table = []
    for row_index in range(0, max_row_index):
        row = []
        for column_index in range(0, max_column_index):
            cur_cell = worksheet.cells.check_cell(row_index, column_index)
            if cur_cell == None:
                row.append(None)
            else:
                row.append(cur_cell.value)
                # if cur_cell.type == CellValueType.IS_NUMERIC:
                #     row.append(cur_cell.value)
                # else :
                #     row.append(0)
        table.append(row)
    return np.asarray(table)


def worksheet_to_dataframe(worksheet: Worksheet) -> pd.DataFrame:
    """
    Convert Worksheet to DataFrame.
    :param Worksheet worksheet:  (required)
    :return DataFrame:
    """
    max_row_index = worksheet.cells.max_row
    max_column_index = worksheet.cells.max_column
    show_table_header = __has_table_header(
        worksheet.cells, 0, 0, max_row_index, max_column_index
    )
    return __get_dataframe(
        worksheet.cells, 0, 0, max_row_index, max_column_index, show_table_header, False
    )


def list_object_to_list(list_object: ListObject) -> list:
    """
    Convert listobject to list.
    :param ListObject list_object:  (required)
    :return list:
    """
    cells = list_object.data_range.worksheet.cells
    table = []
    for row_index in range(list_object.start_row, list_object.end_row + 1):
        row = []
        for column_index in range(list_object.start_column, list_object.end_column + 1):
            cur_cell = cells.check_cell(row_index, column_index)
            if cur_cell == None:
                row.append("")
            else:
                row.append(cur_cell.value)
        table.append(row)
    return table


def list_object_to_tuple(list_object: ListObject) -> tuple:
    """
    Convert ListObject to tuple.
    :param ListObject list_object:  (required)
    :return tuple:
    """
    cells = list_object.data_range.worksheet.cells
    table = []
    for row_index in range(list_object.start_row, list_object.end_row + 1):
        row = []
        for column_index in range(list_object.start_column, list_object.end_column + 1):
            cur_cell = cells.check_cell(row_index, column_index)
            if cur_cell == None:
                row.append("")
            else:
                row.append(cur_cell.value)
        table.append(row)
    return tuple(table)


def list_object_to_ndarray(list_object: ListObject) -> np.ndarray:
    """
    Convert ListObject to ndarray.
    :param ListObject list_object:  (required)
    :return ndarray:
    """
    cells = list_object.data_range.worksheet.cells
    table = []
    for row_index in range(list_object.data_range.first_row, list_object.end_row + 1):
        row = []
        for column_index in range(list_object.start_column, list_object.end_column + 1):
            cur_cell = cells.check_cell(row_index, column_index)
            if cur_cell == None:
                row.append(0)
            else:
                row.append(cur_cell.value)
                # if cur_cell.type == CellValueType.IS_NUMERIC:
                #     row.append(cur_cell.value)
                # else :
                #     row.append(0)
        table.append(row)
    return np.asarray(table)


def list_object_to_dataframe(list_object: ListObject) -> pd.DataFrame:
    """
    Convert ListObject to DataFrame.
    :param ListObject list_object:  (required)
    :return DataFrame:
    """
    cells = list_object.data_range.worksheet.cells
    show_table_header = __has_table_header(
        cells,
        list_object.start_row,
        list_object.start_column,
        list_object.end_row,
        list_object.end_column,
    )
    return __get_dataframe(
        cells,
        list_object.start_row,
        list_object.start_column,
        list_object.end_row,
        list_object.end_column,
        show_table_header,
        False,
    )


def range_to_list(range_name: Range) -> list:
    """
    Convert Range to list.
    :param Range range_name:  (required)
    :return list:
    """
    cells = range_name.worksheet.cells
    table = []
    for row_index in range(
        range_name.first_row, range_name.first_row + range_name.row_count
    ):
        row = []
        for column_index in range(
            range_name.first_column, range_name.first_column + range_name.column_count
        ):
            cur_cell = cells.check_cell(row_index, column_index)
            if cur_cell == None:
                row.append("")
            else:
                row.append(cur_cell.value)
        table.append(row)
    return table


def range_to_tuple(range_name: Range) -> tuple:
    """
    Convert Range to tuple.
    :param Range range:  (required)
    :return tuple:
    """
    cells = range_name.worksheet.cells
    table = []
    for row_index in range(
        range_name.first_row, range_name.first_row + range_name.row_count
    ):
        row = []
        for column_index in range(
            range_name.first_column, range_name.first_column + range_name.column_count
        ):
            cur_cell = cells.check_cell(row_index, column_index)
            if cur_cell == None:
                row.append("")
            else:
                row.append(cur_cell.value)
        table.append(row)
    return tuple(table)


def range_to_ndarray(range_name: Range) -> np.ndarray:
    """
    Convert Range to ndarray.
    :param Range range:  (required)
    :return ndarray:
    """
    cells = range_name.worksheet.cells
    table = []
    for row_index in range(
        range_name.first_row, range_name.first_row + range_name.row_count
    ):
        row = []
        for column_index in range(
            range_name.first_column, range_name.first_column + range_name.column_count
        ):
            cur_cell = cells.check_cell(row_index, column_index)
            if cur_cell == None:
                row.append(0)
            else:
                row.append(cur_cell.value)
                # if cur_cell.type == CellValueType.IS_NUMERIC:
                #     row.append(cur_cell.value)
                # else :
                #     row.append(0)
        table.append(row)
    return np.asarray(table)


def range_to_dataframe(range_name: Range) -> pd.DataFrame:
    """
    Convert Range to DataFrame.
    :param Range range:  (required)
    :return DataFrame:
    """
    cells = range_name.worksheet.cells
    show_table_header = __has_table_header(
        cells,
        range_name.first_row,
        range_name.first_column,
        range_name.first_row + range_name.row_count - 1,
        range_name.first_column + range_name.column_count - 1,
    )
    return __get_dataframe(
        cells,
        range_name.first_row,
        range_name.first_column,
        range_name.first_row + range_name.row_count - 1,
        range_name.first_column + range_name.column_count - 1,
        show_table_header,
        False,
    )


def name_to_list(name: Name) -> list:
    """
    Convert Range to list.
    :param Range range_name:  (required)
    :return list:
    """
    range_name = name.get_range()
    cells = range_name.worksheet.cells
    table = []
    for row_index in range(
        range_name.first_row, range_name.first_row + range_name.row_count
    ):
        row = []
        for column_index in range(
            range_name.first_column, range_name.first_column + range_name.column_count
        ):
            cur_cell = cells.check_cell(row_index, column_index)
            if cur_cell == None:
                row.append("")
            else:
                row.append(cur_cell.value)
        table.append(row)
    return table


def name_to_tuple(name: Name) -> tuple:
    """
    Convert Range to tuple.
    :param Range range:  (required)
    :return tuple:
    """
    range_name = name.get_range()
    cells = range_name.worksheet.cells
    table = []
    for row_index in range(
        range_name.first_row, range_name.first_row + range_name.row_count
    ):
        row = []
        for column_index in range(
            range_name.first_column, range_name.first_column + range_name.column_count
        ):
            cur_cell = cells.check_cell(row_index, column_index)
            if cur_cell == None:
                row.append("")
            else:
                row.append(cur_cell.value)
        table.append(row)
    return tuple(table)


def name_to_ndarray(name: Name) -> np.ndarray:
    """
    Convert Range to ndarray.
    :param Range range:  (required)
    :return ndarray:
    """
    range_name = name.get_range()
    cells = range_name.worksheet.cells
    table = []
    for row_index in range(
        range_name.first_row, range_name.first_row + range_name.row_count
    ):
        row = []
        for column_index in range(
            range_name.first_column, range_name.first_column + range_name.column_count
        ):
            cur_cell = cells.check_cell(row_index, column_index)
            if cur_cell == None:
                row.append(0)
            else:
                row.append(cur_cell.value)
                # if cur_cell.type == CellValueType.IS_NUMERIC:
                #     row.append(cur_cell.value)
                # else :
                #     row.append(0)
        table.append(row)
    return np.asarray(table)


def name_to_dataframe(name: Name) -> pd.DataFrame:
    """
    Convert Range to DataFrame.
    :param Range range:  (required)
    :return DataFrame:
    """
    range_name = name.get_range()
    cells = range_name.worksheet.cells
    show_table_header = __has_table_header(
        cells,
        range_name.first_row,
        range_name.first_column,
        range_name.first_row + range_name.row_count - 1,
        range_name.first_column + range_name.column_count - 1,
    )
    return __get_dataframe(
        cells,
        range_name.first_row,
        range_name.first_column,
        range_name.first_row + range_name.row_count - 1,
        range_name.first_column + range_name.column_count - 1,
        show_table_header,
        False,
    )


##  python object to cells object
def list_to_worksheet(data: list, worksheet: Worksheet, **kwargs) -> Worksheet:
    """
    import list data into worksheet.
    :param list data:  (required)
    :param Worksheet worksheet: . (required)
    :param int begin_row_index: The row index of worksheet indicating the position in the imported data workbook. If the index is None, the default index is 0. (optional)
    :param int begin_column_index: The column index of worksheet indicating the position in the imported data workbook. If the index is None, the default index is 0. (optional)
    :param bool is_vertical: Indicate whether the data is inserted vertically. The default value is False. (optional)
    :return Worksheet:
    """
    is_vertical = False
    if kwargs.get("is_vertical") is not None:
        is_vertical = kwargs.get("is_vertical")
    begin_row_index = 0
    if kwargs.get("begin_row_index") is not None:
        begin_row_index = kwargs.get("begin_row_index")
    begin_column_index = 0
    if kwargs.get("begin_column_index") is not None:
        begin_column_index = kwargs.get("begin_column_index")
    __import_table_data_into_workbook(
        worksheet.cells, data, begin_row_index, begin_column_index, is_vertical
    )

    return worksheet


def tuple_to_worksheet(data: tuple, worksheet: Worksheet, **kwargs) -> Worksheet:
    """
    import tuple data into worksheet.
    :param tuple data:  (required)
    :param Worksheet worksheet: . (required)
    :param int begin_row_index: The row index of worksheet indicating the position in the imported data workbook. If the index is None, the default index is 0. (optional)
    :param int begin_column_index: The column index of worksheet indicating the position in the imported data workbook. If the index is None, the default index is 0. (optional)
    :param bool is_vertical: Indicate whether the data is inserted vertically. The default value is False. (optional)
    :param bool only_ready: Indicate whether the data is only read data. The default value is True. (optional)
    :return Worksheet:
    """
    is_vertical = False
    if kwargs.get("is_vertical") is not None:
        is_vertical = kwargs.get("is_vertical")
    only_ready = True
    if kwargs.get("only_ready") is not None:
        only_ready = kwargs.get("only_ready")
    begin_row_index = 0
    if kwargs.get("begin_row_index") is not None:
        begin_row_index = kwargs.get("begin_row_index")
    begin_column_index = 0
    if kwargs.get("begin_column_index") is not None:
        begin_column_index = kwargs.get("begin_column_index")
    __import_table_data_into_workbook(
        worksheet.cells, data, begin_row_index, begin_column_index, is_vertical
    )

    if only_ready:
        worksheet.protect(ProtectionType.CONTENTS)
        worksheet.protection.password = ""

    return worksheet


def ndarray_to_worksheet(data: np.ndarray, worksheet: Worksheet, **kwargs) -> Worksheet:
    """
    import ndarray data into worksheet.
    :param ndarray data:  (required)
    :param Worksheet worksheet: . (required)
    :param int begin_row_index: The row index of worksheet indicating the position in the imported data workbook. If the index is None, the default index is 0. (optional)
    :param int begin_column_index: The column index of worksheet indicating the position in the imported data workbook. If the index is None, the default index is 0. (optional)
    :param bool is_vertical: Indicate whether the data is inserted vertically. The default value is False. (optional)
    :param bool only_ready: Indicate whether the data is only read data. The default value is True. (optional)
    :return Worksheet:
    """
    is_vertical = False
    if kwargs.get("is_vertical") is not None:
        is_vertical = kwargs.get("is_vertical")
    only_ready = False
    if kwargs.get("only_ready") is not None:
        only_ready = kwargs.get("only_ready")
    begin_row_index = 0
    if kwargs.get("begin_row_index") is not None:
        begin_row_index = kwargs.get("begin_row_index")
    begin_column_index = 0
    if kwargs.get("begin_column_index") is not None:
        begin_column_index = kwargs.get("begin_column_index")
    __import_table_data_into_workbook(
        worksheet.cells, data, begin_row_index, begin_column_index, is_vertical
    )

    if only_ready:
        worksheet.protect(ProtectionType.CONTENTS)
        worksheet.protection.password = ""

    return worksheet


def dataframe_to_worksheet(
    data: pd.DataFrame, worksheet: Worksheet, **kwargs
) -> Worksheet:
    """
    import dataframe data into worksheet.
    :param DataFrame data:  (required)
    :param Worksheet worksheet: . (required)
    :param int begin_row_index: The row index of worksheet indicating the position in the imported data workbook. If the index is None, the default index is 0. (optional)
    :param int begin_column_index: The column index of worksheet indicating the position in the imported data workbook. If the index is None, the default index is 0. (optional)
    :param bool is_vertical: Indicate whether the data is inserted vertically. The default value is False. (optional)
    :param bool only_ready: Indicate whether the data is only read data. The default value is True. (optional)
    :return Worksheet:
    """
    is_vertical = False
    if kwargs.get("is_vertical") is not None:
        is_vertical = kwargs.get("is_vertical")
    begin_row_index = 0
    if kwargs.get("begin_row_index") is not None:
        begin_row_index = kwargs.get("begin_row_index")
    begin_column_index = 0
    if kwargs.get("begin_column_index") is not None:
        begin_column_index = kwargs.get("begin_column_index")
    cells = worksheet.cells
    df_row_index = begin_row_index
    df_column_index = begin_column_index
    for column_name in data.columns:
        df_row_index = begin_row_index
        __put_value_to_cell(cells, column_name, df_row_index, df_column_index + 1)
        for df_value in data[column_name]:
            __put_value_to_cell(cells, df_value, df_row_index + 1, df_column_index + 1)
            df_row_index = df_row_index + 1
        df_column_index = df_column_index + 1
    df_row_index = begin_row_index
    df_column_index = begin_column_index
    for df_row_name in data.index.values:
        __put_value_to_cell(cells, df_row_name, df_row_index + 1, df_column_index)
        df_row_index = df_row_index + 1

    return worksheet


def list_to_range(data: list, worksheet: Worksheet, **kwargs) -> Range:
    """
    convert list data to range in the worksheet.
    :param list data:  (required)
    :param Worksheet worksheet: . (required)
    :param int begin_row_index: The row index of worksheet indicating the position in the imported data workbook. If the index is None, the default index is 0. (optional)
    :param int begin_column_index: The column index of worksheet indicating the position in the imported data workbook. If the index is None, the default index is 0. (optional)
    :param bool is_vertical: Indicate whether the data is inserted vertically. The default value is False. (optional)
    :return Range:
    """
    is_vertical = False
    if kwargs.get("is_vertical") is not None:
        is_vertical = kwargs.get("is_vertical")
    begin_row_index = 0
    if kwargs.get("begin_row_index") is not None:
        begin_row_index = kwargs.get("begin_row_index")
    begin_column_index = 0
    if kwargs.get("begin_column_index") is not None:
        begin_column_index = kwargs.get("begin_column_index")
    (begin_row_index, begin_column_index, end_row_index, end_column_index) = (
        __import_table_data_into_workbook(
            worksheet.cells, data, begin_row_index, begin_column_index, is_vertical
        )
    )

    return worksheet.cells.create_range(
        begin_row_index,
        begin_column_index,
        end_row_index - begin_row_index + 1,
        end_column_index - begin_column_index + 1,
    )


def tuple_to_range(data: tuple, worksheet: Worksheet, **kwargs) -> Range:
    """
    convert tuple data to range in the worksheet.
    :param tuple data:  (required)
    :param Worksheet worksheet: . (required)
    :param int begin_row_index: The row index of worksheet indicating the position in the imported data workbook. If the index is None, the default index is 0. (optional)
    :param int begin_column_index: The column index of worksheet indicating the position in the imported data workbook. If the index is None, the default index is 0. (optional)
    :param bool is_vertical: Indicate whether the data is inserted vertically. The default value is False. (optional)
    :param bool only_ready: Indicate whether the data is only read data. The default value is True. (optional)
    :return Range:
    """
    is_vertical = False
    if kwargs.get("is_vertical") is not None:
        is_vertical = kwargs.get("is_vertical")
    only_ready = True
    if kwargs.get("only_ready") is not None:
        only_ready = kwargs.get("only_ready")
    begin_row_index = 0
    if kwargs.get("begin_row_index") is not None:
        begin_row_index = kwargs.get("begin_row_index")
    begin_column_index = 0
    if kwargs.get("begin_column_index") is not None:
        begin_column_index = kwargs.get("begin_column_index")
    (begin_row_index, begin_column_index, end_row_index, end_column_index) = (
        __import_table_data_into_workbook(
            worksheet.cells, data, begin_row_index, begin_column_index, is_vertical
        )
    )

    if only_ready:
        worksheet.protect(ProtectionType.CONTENTS)
        worksheet.protection.password = ""

    return worksheet.cells.create_range(
        begin_row_index,
        begin_column_index,
        end_row_index - begin_row_index + 1,
        end_column_index - begin_column_index + 1,
    )


def ndarray_to_range(data: np.ndarray, worksheet: Worksheet, **kwargs) -> Range:
    """
    convert ndarray data to range in the worksheet.
    :param ndarray data:  (required)
    :param Worksheet worksheet: . (required)
    :param int begin_row_index: The row index of worksheet indicating the position in the imported data workbook. If the index is None, the default index is 0. (optional)
    :param int begin_column_index: The column index of worksheet indicating the position in the imported data workbook. If the index is None, the default index is 0. (optional)
    :param bool is_vertical: Indicate whether the data is inserted vertically. The default value is False. (optional)
    :param bool only_ready: Indicate whether the data is only read data. The default value is True. (optional)
    :return Range:
    """
    is_vertical = False
    if kwargs.get("is_vertical") is not None:
        is_vertical = kwargs.get("is_vertical")
    only_ready = False
    if kwargs.get("only_ready") is not None:
        only_ready = kwargs.get("only_ready")
    begin_row_index = 0
    if kwargs.get("begin_row_index") is not None:
        begin_row_index = kwargs.get("begin_row_index")
    begin_column_index = 0
    if kwargs.get("begin_column_index") is not None:
        begin_column_index = kwargs.get("begin_column_index")
    (begin_row_index, begin_column_index, end_row_index, end_column_index) = (
        __import_table_data_into_workbook(
            worksheet.cells, data, begin_row_index, begin_column_index, is_vertical
        )
    )

    if only_ready:
        worksheet.protect(ProtectionType.CONTENTS)
        worksheet.protection.password = ""

    return worksheet.cells.create_range(
        begin_row_index,
        begin_column_index,
        end_row_index - begin_row_index + 1,
        end_column_index - begin_column_index + 1,
    )


def dataframe_to_range(data: pd.DataFrame, worksheet: Worksheet, **kwargs) -> Range:
    """
    convert dataframe data to range in the worksheet.
    :param DataFrame data:  (required)
    :param Worksheet worksheet: . (required)
    :param int begin_row_index: The row index of worksheet indicating the position in the imported data workbook. If the index is None, the default index is 0. (optional)
    :param int begin_column_index: The column index of worksheet indicating the position in the imported data workbook. If the index is None, the default index is 0. (optional)
    :param bool is_vertical: Indicate whether the data is inserted vertically. The default value is False. (optional)
    :param bool only_ready: Indicate whether the data is only read data. The default value is True. (optional)
    :return Range:
    """
    is_vertical = False
    if kwargs.get("is_vertical") is not None:
        is_vertical = kwargs.get("is_vertical")
    begin_row_index = 0
    if kwargs.get("begin_row_index") is not None:
        begin_row_index = kwargs.get("begin_row_index")
    begin_column_index = 0
    if kwargs.get("begin_column_index") is not None:
        begin_column_index = kwargs.get("begin_column_index")
    cells = worksheet.cells
    df_row_index = begin_row_index
    df_column_index = begin_column_index
    for column_name in data.columns:
        df_row_index = begin_row_index
        __put_value_to_cell(cells, column_name, df_row_index, df_column_index + 1)
        for df_value in data[column_name]:
            __put_value_to_cell(cells, df_value, df_row_index + 1, df_column_index + 1)
            df_row_index = df_row_index + 1
        df_column_index = df_column_index + 1
    df_row_index = begin_row_index
    df_column_index = begin_column_index
    for df_row_name in data.index.values:
        __put_value_to_cell(cells, df_row_name, df_row_index + 1, df_column_index)
        df_row_index = df_row_index + 1
    return worksheet.cells.create_range(
        begin_row_index,
        begin_column_index,
        df_row_index - begin_row_index + 1,
        df_column_index - begin_column_index + 1,
    )


def list_to_name(data: list, worksheet: Worksheet, **kwargs) -> Name:
    """
    convert list data to name in the worksheet.
    :param list data:  (required)
    :param Worksheet worksheet: . (required)
    :param int begin_row_index: The row index of worksheet indicating the position in the imported data workbook. If the index is None, the default index is 0. (optional)
    :param int begin_column_index: The column index of worksheet indicating the position in the imported data workbook. If the index is None, the default index is 0. (optional)
    :param bool is_vertical: Indicate whether the data is inserted vertically. The default value is False. (optional)
    :return Name:
    """
    is_vertical = False
    if kwargs.get("is_vertical") is not None:
        is_vertical = kwargs.get("is_vertical")
    begin_row_index = 0
    if kwargs.get("begin_row_index") is not None:
        begin_row_index = kwargs.get("begin_row_index")
    begin_column_index = 0
    if kwargs.get("begin_column_index") is not None:
        begin_column_index = kwargs.get("begin_column_index")
    (begin_row_index, begin_column_index, end_row_index, end_column_index) = (
        __import_table_data_into_workbook(
            worksheet.cells, data, begin_row_index, begin_column_index, is_vertical
        )
    )

    name_text = "Name_" + str(len(worksheet.workbook.worksheets.names))
    name_refers_to = (
        "="
        + worksheet.name
        + "!$"
        + CellsHelper.column_index_to_name(begin_column_index)
        + "$"
        + str(begin_row_index + 1)
        + ":$"
        + CellsHelper.column_index_to_name(end_column_index)
        + "$"
        + str(end_row_index + 1)
    )
    position = worksheet.workbook.worksheets.names.add(name_text)
    name = worksheet.workbook.worksheets.names[position]
    name.refers_to = name_refers_to
    return name


def tuple_to_name(data: tuple, worksheet: Worksheet, **kwargs) -> Name:
    """
    convert tuple data to name in the worksheet.
    :param tuple data:  (required)
    :param Worksheet worksheet: . (required)
    :param int begin_row_index: The row index of worksheet indicating the position in the imported data workbook. If the index is None, the default index is 0. (optional)
    :param int begin_column_index: The column index of worksheet indicating the position in the imported data workbook. If the index is None, the default index is 0. (optional)
    :param bool is_vertical: Indicate whether the data is inserted vertically. The default value is False. (optional)
    :param bool only_ready: Indicate whether the data is only read data. The default value is True. (optional)
    :return Name:
    """
    is_vertical = False
    if kwargs.get("is_vertical") is not None:
        is_vertical = kwargs.get("is_vertical")
    only_ready = True
    if kwargs.get("only_ready") is not None:
        only_ready = kwargs.get("only_ready")
    begin_row_index = 0
    if kwargs.get("begin_row_index") is not None:
        begin_row_index = kwargs.get("begin_row_index")
    begin_column_index = 0
    if kwargs.get("begin_column_index") is not None:
        begin_column_index = kwargs.get("begin_column_index")
    (begin_row_index, begin_column_index, end_row_index, end_column_index) = (
        __import_table_data_into_workbook(
            worksheet.cells, data, begin_row_index, begin_column_index, is_vertical
        )
    )

    if only_ready:
        worksheet.protect(ProtectionType.CONTENTS)
        worksheet.protection.password = ""

    name_text = "Name_" + str(len(worksheet.workbook.worksheets.names))
    name_refers_to = (
        "="
        + worksheet.name
        + "!$"
        + CellsHelper.column_index_to_name(begin_column_index)
        + "$"
        + str(begin_row_index + 1)
        + ":$"
        + CellsHelper.column_index_to_name(end_column_index)
        + "$"
        + str(end_row_index + 1)
    )
    position = worksheet.workbook.worksheets.names.add(name_text)
    name = worksheet.workbook.worksheets.names[position]
    name.refers_to = name_refers_to
    return name


def ndarray_to_name(data: np.ndarray, worksheet: Worksheet, **kwargs) -> Name:
    """
    convert ndarray data to name in the worksheet.
    :param ndarray data:  (required)
    :param Worksheet worksheet: . (required)
    :param int begin_row_index: The row index of worksheet indicating the position in the imported data workbook. If the index is None, the default index is 0. (optional)
    :param int begin_column_index: The column index of worksheet indicating the position in the imported data workbook. If the index is None, the default index is 0. (optional)
    :param bool is_vertical: Indicate whether the data is inserted vertically. The default value is False. (optional)
    :param bool only_ready: Indicate whether the data is only read data. The default value is True. (optional)
    :return Name:
    """
    is_vertical = False
    if kwargs.get("is_vertical") is not None:
        is_vertical = kwargs.get("is_vertical")
    only_ready = False
    if kwargs.get("only_ready") is not None:
        only_ready = kwargs.get("only_ready")
    begin_row_index = 0
    if kwargs.get("begin_row_index") is not None:
        begin_row_index = kwargs.get("begin_row_index")
    begin_column_index = 0
    if kwargs.get("begin_column_index") is not None:
        begin_column_index = kwargs.get("begin_column_index")
    (begin_row_index, begin_column_index, end_row_index, end_column_index) = (
        __import_table_data_into_workbook(
            worksheet.cells, data, begin_row_index, begin_column_index, is_vertical
        )
    )

    if only_ready:
        worksheet.protect(ProtectionType.CONTENTS)
        worksheet.protection.password = ""

    name_text = "Name_" + str(len(worksheet.workbook.worksheets.names))
    name_refers_to = (
        "="
        + worksheet.name
        + "!$"
        + CellsHelper.column_index_to_name(begin_column_index)
        + "$"
        + str(begin_row_index + 1)
        + ":$"
        + CellsHelper.column_index_to_name(end_column_index)
        + "$"
        + str(end_row_index + 1)
    )
    position = worksheet.workbook.worksheets.names.add(name_text)
    name = worksheet.workbook.worksheets.names[position]
    name.refers_to = name_refers_to
    return name


def dataframe_to_name(data: pd.DataFrame, worksheet: Worksheet, **kwargs) -> Name:
    """
    convert dataframe data to name in the worksheet.
    :param DataFrame data:  (required)
    :param Worksheet worksheet: . (required)
    :param int begin_row_index: The row index of worksheet indicating the position in the imported data workbook. If the index is None, the default index is 0. (optional)
    :param int begin_column_index: The column index of worksheet indicating the position in the imported data workbook. If the index is None, the default index is 0. (optional)
    :param bool is_vertical: Indicate whether the data is inserted vertically. The default value is False. (optional)
    :param bool only_ready: Indicate whether the data is only read data. The default value is True. (optional)
    :return Name:
    """
    is_vertical = False
    if kwargs.get("is_vertical") is not None:
        is_vertical = kwargs.get("is_vertical")
    begin_row_index = 0
    if kwargs.get("begin_row_index") is not None:
        begin_row_index = kwargs.get("begin_row_index")
    begin_column_index = 0
    if kwargs.get("begin_column_index") is not None:
        begin_column_index = kwargs.get("begin_column_index")
    cells = worksheet.cells
    df_row_index = begin_row_index
    df_column_index = begin_column_index
    for column_name in data.columns:
        df_row_index = begin_row_index
        __put_value_to_cell(cells, column_name, df_row_index, df_column_index + 1)
        for df_value in data[column_name]:
            __put_value_to_cell(cells, df_value, df_row_index + 1, df_column_index + 1)
            df_row_index = df_row_index + 1
        df_column_index = df_column_index + 1
    df_row_index = begin_row_index
    df_column_index = begin_column_index
    for df_row_name in data.index.values:
        __put_value_to_cell(cells, df_row_name, df_row_index + 1, df_column_index)
        df_row_index = df_row_index + 1
    name_text = "Name_" + str(len(worksheet.workbook.worksheets.names))
    name_refers_to = (
        "="
        + worksheet.name
        + "!$"
        + CellsHelper.column_index_to_name(begin_column_index)
        + "$"
        + str(begin_row_index + 1)
        + ":$"
        + CellsHelper.column_index_to_name(df_column_index)
        + "$"
        + str(df_row_index + 1)
    )
    position = worksheet.workbook.worksheets.names.add(name_text)
    name = worksheet.workbook.worksheets.names[position]
    name.refers_to = name_refers_to
    return name


def list_to_list_object(data: list, worksheet: Worksheet, **kwargs) -> ListObject:
    """
    create list object with list data on the worksheet.
    :param list data:  (required)
    :param Worksheet worksheet: . (required)
    :param int begin_row_index: The row index of worksheet indicating the position in the imported data workbook. If the index is None, the default index is 0. (optional)
    :param int begin_column_index: The column index of worksheet indicating the position in the imported data workbook. If the index is None, the default index is 0. (optional)
    :param bool is_vertical: Indicate whether the data is inserted vertically. The default value is False. (optional)
    :param bool has_table_header: Indicate whether has table header. The default value is True. (optional)
    :return Worksheet:
    """
    is_vertical = False
    if kwargs.get("is_vertical") is not None:
        is_vertical = kwargs.get("is_vertical")
    has_table_header = True
    
    if kwargs.get("has_table_header") is not None:
        if kwargs.get("has_table_header") == False:
            has_table_header = False
    begin_row_index = 0
    if kwargs.get("begin_row_index") is not None:
        begin_row_index = kwargs.get("begin_row_index")
    begin_column_index = 0
    if kwargs.get("begin_column_index") is not None:
        begin_column_index = kwargs.get("begin_column_index")
    (begin_row_index, begin_column_index, end_row_index, end_column_index) = (
        __import_table_data_into_workbook(
            worksheet.cells, data, begin_row_index, begin_column_index, is_vertical
        )
    )
    
    index = worksheet.list_objects.add(
        begin_row_index, begin_column_index, end_row_index, end_column_index, has_table_header
    )
    return worksheet.list_objects[index]
    pass


def tuple_to_list_object(data: tuple, worksheet: Worksheet, **kwargs) -> ListObject:
    """
    create list object with tuple data on the worksheet.
    :param tuple data:  (required)
    :param Worksheet worksheet: . (required)
    :param int begin_row_index: The row index of worksheet indicating the position in the imported data workbook. If the index is None, the default index is 0. (optional)
    :param int begin_column_index: The column index of worksheet indicating the position in the imported data workbook. If the index is None, the default index is 0. (optional)
    :param bool is_vertical: Indicate whether the data is inserted vertically. The default value is False. (optional)
    :param bool only_ready: Indicate whether the data is only read data. The default value is True. (optional)
    :return ListObject:
    """
    is_vertical = False
    if kwargs.get("is_vertical") is not None:
        is_vertical = kwargs.get("is_vertical")
    only_ready = True
    if kwargs.get("only_ready") is not None:
        only_ready = kwargs.get("only_ready")
    begin_row_index = 0
    if kwargs.get("begin_row_index") is not None:
        begin_row_index = kwargs.get("begin_row_index")
    begin_column_index = 0
    if kwargs.get("begin_column_index") is not None:
        begin_column_index = kwargs.get("begin_column_index")
    (begin_row_index, begin_column_index, end_row_index, end_column_index) = (
        __import_table_data_into_workbook(
            worksheet.cells, data, begin_row_index, begin_column_index, is_vertical
        )
    )

    if only_ready:
        worksheet.protect(ProtectionType.CONTENTS)
        worksheet.protection.password = ""
    index = worksheet.list_objects.add(
        begin_row_index, begin_column_index, end_row_index, end_column_index, True
    )
    return worksheet.list_objects[index]


def ndarray_to_list_object(
    data: np.ndarray, worksheet: Worksheet, **kwargs
) -> ListObject:
    """
    create list object with ndarray data on the worksheet.
    :param ndarray data:  (required)
    :param Worksheet worksheet: . (required)
    :param int begin_row_index: The row index of worksheet indicating the position in the imported data workbook. If the index is None, the default index is 0. (optional)
    :param int begin_column_index: The column index of worksheet indicating the position in the imported data workbook. If the index is None, the default index is 0. (optional)
    :param bool is_vertical: Indicate whether the data is inserted vertically. The default value is False. (optional)
    :param bool only_ready: Indicate whether the data is only read data. The default value is True. (optional)
    :return Worksheet:
    """
    is_vertical = False
    if kwargs.get("is_vertical") is not None:
        is_vertical = kwargs.get("is_vertical")
    only_ready = False
    if kwargs.get("only_ready") is not None:
        only_ready = kwargs.get("only_ready")
    begin_row_index = 0
    if kwargs.get("begin_row_index") is not None:
        begin_row_index = kwargs.get("begin_row_index")
    begin_column_index = 0
    if kwargs.get("begin_column_index") is not None:
        begin_column_index = kwargs.get("begin_column_index")
    (begin_row_index, begin_column_index, end_row_index, end_column_index) = (
        __import_table_data_into_workbook(
            worksheet.cells, data, begin_row_index, begin_column_index, is_vertical
        )
    )
    index = worksheet.list_objects.add(
        begin_row_index, begin_column_index, end_row_index, end_column_index, True
    )

    return worksheet.list_objects[index]


def dataframe_to_list_object(
    data: pd.DataFrame, worksheet: Worksheet, **kwargs
) -> ListObject:
    """
    create list object with dataframe data on the worksheet.
    :param DataFrame data:  (required)
    :param Worksheet worksheet: . (required)
    :param int begin_row_index: The row index of worksheet indicating the position in the imported data workbook. If the index is None, the default index is 0. (optional)
    :param int begin_column_index: The column index of worksheet indicating the position in the imported data workbook. If the index is None, the default index is 0. (optional)
    :param bool is_vertical: Indicate whether the data is inserted vertically. The default value is False. (optional)
    :param bool only_ready: Indicate whether the data is only read data. The default value is True. (optional)
    :return ListObject:
    """
    is_vertical = False
    if kwargs.get("is_vertical") is not None:
        is_vertical = kwargs.get("is_vertical")
    begin_row_index = 0
    if kwargs.get("begin_row_index") is not None:
        begin_row_index = kwargs.get("begin_row_index")
    begin_column_index = 0
    if kwargs.get("begin_column_index") is not None:
        begin_column_index = kwargs.get("begin_column_index")
    cells = worksheet.cells
    df_row_index = begin_row_index
    df_column_index = begin_column_index
    for column_name in data.columns:
        df_row_index = begin_row_index
        __put_value_to_cell(cells, column_name, df_row_index, df_column_index + 1)
        for df_value in data[column_name]:
            __put_value_to_cell(cells, df_value, df_row_index + 1, df_column_index + 1)
            df_row_index = df_row_index + 1
        df_column_index = df_column_index + 1
    df_row_index = begin_row_index
    df_column_index = begin_column_index
    for df_row_name in data.index.values:
        __put_value_to_cell(cells, df_row_name, df_row_index + 1, df_column_index)
        df_row_index = df_row_index + 1
    index = worksheet.list_objects.add(
        begin_row_index, begin_column_index, df_row_index, df_column_index, True
    )

    return worksheet.list_objects[index]


def __get_dataframe(
    cells: Cells,
    begin_row_index: int,
    begin_column_index: int,
    end_row_index: int,
    end_column_index: int,
    has_header: bool,
    has_total: bool,
):
    column_title_list = []
    row_index = 0
    cells_helper = CellsHelper
    if has_header:
        row_index = begin_row_index
    for column_index in range(begin_column_index, end_column_index + 1):
        if has_header:
            column_title_list.append(
                cells.get(row_index, column_index).display_string_value
            )
        else:
            column_title_list.append(cells_helper.column_index_to_name(column_index))

    start_row = 0
    end_row = 0
    if has_header:
        start_row = begin_row_index + 1
    else:
        start_row = begin_row_index

    if has_total:
        end_row = end_row_index
    else:
        end_row = end_row_index + 1

    position = 0
    data = {}
    for column_index in range(begin_column_index, end_column_index + 1):
        column_data = []
        for row_index in range(start_row, end_row):
            column_data.append(cells.get(row_index, column_index).value)
        data[column_title_list[position]] = column_data
        position = position + 1
    return pd.DataFrame(data)


def __has_table_header(
    cells: Cells,
    begin_row_index: int,
    begin_column_index: int,
    end_row_index: int,
    end_column_index: int,
):
    has_header = True
    for column_index in range(begin_column_index, end_column_index + 1):
        cell = cells.get(begin_row_index, column_index)
        if cell.type != CellValueType.IS_STRING:
            has_header = False
            break
        sen_cell = cells.get(begin_row_index + 1, column_index)
        if cell.type != sen_cell.type:
            break
    return has_header


def __import_table_data_into_workbook(
    cells, table_data, row_index, column_index, is_vertical
):    
    table_row_index = row_index
    table_column_index = column_index    
    row_count = len(table_data)
    column_count = 0
    for table_row in table_data:
        for table_cell in table_row:
            column_count = len(table_row)
            __put_value_to_cell(cells, table_cell, table_row_index, table_column_index)
            if is_vertical:
                table_row_index = table_row_index + 1
            else:
                table_column_index = table_column_index + 1
        if is_vertical:
            table_row_index = row_index
            table_column_index = table_column_index + 1
        else:
            table_column_index = column_index
            table_row_index = table_row_index + 1

    if is_vertical:
        end_row_index = row_index + row_count - 1
        end_column_index = table_column_index - 1
    else:
        end_row_index = table_row_index - 1
        end_column_index = column_index + column_count - 1

    return (row_index, column_index, end_row_index, end_column_index)


def __put_value_to_cell(cells, raw_value, row, column):
    cell = cells.get(row, column)
    dtype = type(raw_value)
    match dtype:
        case np.bool_:
            value = bool(raw_value)
        case np.int_:
            value = int(raw_value)
        case np.intc:
            value = int(raw_value)
        case np.intp:
            value = int(raw_value)
        case np.int8:
            value = int(raw_value)
        case np.int16:
            value = int(raw_value)
        case np.int32:
            value = int(raw_value)
        case np.int64:
            value = int(raw_value)
        case np.uint8:
            value = int(raw_value)
        case np.uint16:
            value = int(raw_value)
        case np.uint32:
            value = int(raw_value)
        case np.uint64:
            value = int(raw_value)
        case np.float_:
            value = int(raw_value)
        case np.float16:
            value = float(raw_value)
        case np.float32:
            value = float(raw_value)
        case np.float64:
            value = float(raw_value)
        case np.single:
            value = float(raw_value)
        case np.double:
            value = float(raw_value)
        case np.datetime64:
            ts = pd.to_datetime(str(raw_value))
            value = ts.strftime("%Y.%m.%d")
        case _:
            value = raw_value
    cell.put_value(value)
    pass