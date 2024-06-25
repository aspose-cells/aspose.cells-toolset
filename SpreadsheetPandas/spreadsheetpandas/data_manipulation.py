"""
    Data Manipulation
"""

from aspose.cells import Workbook
from aspose.cells import Worksheet
from aspose.cells import Cells
from aspose.cells import Cell
from aspose.cells import Range
from aspose.cells import Name
from aspose.cells import CellsHelper
from aspose.cells.tables import ListObject
import numpy as np
import pandas as pd

from spreadsheetpandas.data_conversion import (
    list_object_to_list,
    list_object_to_list_dict,
)


## cells object to python object
def pivot_column(
    table: ListObject, pivot_column: str, value_column: str, aggregation: str, **kwargs
) -> list:
    """
    List Object
    :param ListObject table:  (required)
    :param str pivot_column:  (required)
    :param str value_column:  (required)
    :param str aggregation:  (required)
    :param list out_fields :  A set of output fields that are retained. (optional)
    :param dict date_to_string_fields : A date field convert string format output. (optional)
    :return list:
    """
    # 1. Get table data range and table column index
    cells = table.data_range.worksheet.cells
    pivot_column_index = 0
    value_column_index = -1
    out_fields = None
    out_fields_index = {}
    date_to_string_fields = None
    date_to_string_fields_index = {}
    table_rows = {}
    column_index = 0
    if kwargs.get("out_fields") is not None:
        out_fields = kwargs.get("out_fields")
    if kwargs.get("date_to_string_fields") is not None:
        date_to_string_fields = kwargs.get("date_to_string_fields")

    for column in table.list_columns:
        if column.name == pivot_column:
            pivot_column_index = column_index
        if column.name == value_column:
            value_column_index = column_index
        if out_fields is not None:
            if column.name in out_fields:
                out_fields_index[column_index] = column.name
        if date_to_string_fields is not None:
            if column.name in date_to_string_fields:
                date_to_string_fields_index[column_index] = date_to_string_fields[
                    column.name
                ]

        column_index = column_index + 1
    # 2. table to dict
    table_data_begin_row_index = table.data_range.first_row
    table_data_end_row_index = table.data_range.first_row + table.data_range.row_count
    table_data_begin_column_index = table.data_range.first_column
    table_data_end_column_index = (
        table.data_range.first_column + table.data_range.column_count
    )

    cur_pivot_column_value = None
    cur_value_column_value = None
    cur_row_cell_value = None
    cur_row = None
    column_value_dict = {}
    for row_index in range(table_data_begin_row_index, table_data_end_row_index):
        IsFirstCell = True
        for column_index in range(
            table_data_begin_column_index, table_data_end_column_index
        ):
            if column_index == pivot_column_index:
                cur_pivot_column_value = cells.get(row_index, column_index).value
                if cur_pivot_column_value not in column_value_dict:
                    column_value_dict[cur_pivot_column_value] = cur_pivot_column_value
            elif column_index == value_column_index:
                cur_value_column_value = cells.get(row_index, column_index).value
            else:
                if column_index in date_to_string_fields_index:
                    string_format = date_to_string_fields_index[column_index]
                    cur_row_cell_value = cells.get(
                        row_index, column_index
                    ).value.strftime(string_format)
                else:
                    cur_row_cell_value = cells.get(row_index, column_index).value
                if out_fields is None:
                    pass
                else:
                    if column_index in out_fields_index:
                        pass
                    else:
                        continue
                if IsFirstCell:
                    if cur_row_cell_value in table_rows:
                        cur_row = table_rows[cur_row_cell_value]
                    else:
                        table_rows[cur_row_cell_value] = {}
                        cur_row = table_rows[cur_row_cell_value]
                    IsFirstCell = False
                else:
                    if cur_row_cell_value in cur_row:
                        cur_row = table_rows[cur_row_cell_value]
                    else:
                        cur_row[cur_row_cell_value] = {}
                        cur_row = cur_row[cur_row_cell_value]

        if not bool(cur_row):
            if cur_value_column_value is None:
                cur_row[cur_pivot_column_value] = 1
            else:
                cur_row[cur_pivot_column_value] = cur_value_column_value
        else:
            if cur_value_column_value is None:
                if cur_pivot_column_value in cur_row:
                    cur_row[cur_pivot_column_value] = (
                        cur_row[cur_pivot_column_value] + 1
                    )
                else:
                    cur_row[cur_pivot_column_value] = 1
            else:
                if cur_pivot_column_value in cur_row:
                    cur_row[cur_pivot_column_value] = (
                        cur_row[cur_pivot_column_value] + cur_value_column_value
                    )
                else:
                    cur_row[cur_pivot_column_value] = cur_value_column_value

    # 3. dict to list
    column_value_list = sorted(list(column_value_dict.keys()))
    result = []
    row = []
    table_head = []
    if out_fields is None:
        for column in table.list_columns:
            if column.name == pivot_column:
                pass
            elif column.name == value_column:
                pass
            else:
                table_head.append(column.name)
    else:
        for new_column_name in out_fields:
            table_head.append(new_column_name)
    for new_column_name in column_value_list:
        table_head.append(new_column_name)
    __dict_to_list__(
        table_rows, row, result, 0, len(out_fields_index), column_value_list
    )
    result.insert(0, table_head)
    return result


def unpivot_column(
    table: ListObject, column_names: list, column_map_name: str, value_map_name: str
) -> list:
    """
    List Object
    :param ListObject table:  (required)
    :param str pivot_column:  (required)
    :param str value_column:  (required)
    :param str aggregation:  (required)
    :param list out_fields :  A set of output fields that are retained. (optional)
    :param dict date_to_string_fields : A date field convert string format output. (optional)
    :return list:
    """
    # 1. get basic data.
    rows = []
    cells = table.data_range.worksheet.cells
    column_name_map_column_index = {}
    cur_column_index = 0
    row_head = []
    for column in table.list_columns:
        if column.name in column_names:
            column_name_map_column_index[cur_column_index] = column.name
        else:
            row_head.append(column.name)
        cur_column_index = cur_column_index + 1
    row_head.append(column_map_name)
    row_head.append(value_map_name)
    rows.append(row_head)
    table_data_begin_row_index = table.data_range.first_row
    table_data_end_row_index = table.data_range.first_row + table.data_range.row_count
    table_data_begin_column_index = table.data_range.first_column
    table_data_end_column_index = (
        table.data_range.first_column + table.data_range.column_count
    )

    # 2.table to dict
    for row_index in range(table_data_begin_row_index, table_data_end_row_index):
        row = []
        column_value_list = []
        for column_index in range(
            table_data_begin_column_index, table_data_end_column_index
        ):
            if column_index in column_name_map_column_index:
                column_value_list.append(
                    [
                        column_name_map_column_index[column_index],
                        cells.get(row_index, column_index).value,
                    ]
                )
            else:
                row.append(cells.get(row_index, column_index).value)

        for fields in column_value_list:
            new_row = row.copy()
            new_row.append(fields[0])
            new_row.append(fields[1])
            rows.append(new_row)

    return rows


def left_join(
    table1: ListObject, table2: ListObject, table1_field: str, table2_field: str
) -> list:
    list1 = list_object_to_list(table1)
    list2 = list_object_to_list(table2)
    return __list_join_list(list1, list2, table1_field, table2_field)


def right_join(
    table1: ListObject, table2: ListObject, table1_field: str, table2_field: str
) -> list:
    list1 = list_object_to_list(table1)
    list2 = list_object_to_list(table2)
    return __list_join_list(list2, list1, table2_field, table1_field)


def inner_join(
    table1: ListObject, table2: ListObject, table1_field: str, table2_field: str
) -> list:
    result = []
    list1 = list_object_to_list(table1)
    list2 = list_object_to_list(table2)
    table1_column_index = 0
    table2_column_index = 0
    column_index = 0
    list1_length = len(list1[0])
    list2_length = len(list2[0])
    for column_name in list1[0]:
        if column_name == table1_field:
            table1_column_index = column_index
        column_index = column_index + 1
    column_index = 0
    for column_name in list2[0]:
        if column_name == table1_field:
            table2_column_index = column_index
        else:
            list1[0].append(column_name)
        column_index = column_index + 1

    result.append(list1[0])

    for row in list1[1 : list1_length - 1]:
        field_value = row[table1_column_index]
        loop_break = False
        for table2_row in list2[1 : list2_length - 1]:
            field2_value = table2_row[table2_column_index]
            if field_value == field2_value:
                column_index = 0
                for column_value in table2_row:
                    if column_index == table2_column_index:
                        pass
                    else:
                        row.append(column_value)
                    column_index = column_index + 1
                loop_break = True
                break
        if loop_break:
            result.append(row)

    return result


def merge_table_with_appending_non_matching_rows(
    main_table: ListObject, lookup_main: ListObject, match_column: str
) -> list:
    main_table_data = list_object_to_list_dict(main_table)
    lookup_main_data = list_object_to_list_dict(lookup_main)
    result = []
    result.extend(main_table_data)
    for lookup_row in lookup_main_data:
        key_value = lookup_row[match_column]
        is_match = False
        for main_row in main_table_data:
            if main_row[match_column] == key_value:
                is_match = True
                break
        if not is_match:
            result.append(lookup_row)
    return result


def merge_table_with_appending_additional_matching_rows(
    main_table: ListObject, lookup_main: ListObject, match_column: str
) -> list:
    """
    Merge two tables with appending additional matching rows.
    :param ListObject main_table:  (required)
    :param ListObject lookup_main:  (required)
    :param str match_column:  (required)
    :return list:
    """
    result = []
    main_table_data = list_object_to_list_dict(main_table)
    lookup_main_data = list_object_to_list_dict(lookup_main)
    key_dict = {}
    all_match_rows = {}
    main_table_len = len(main_table_data)
    lookup_table_len = len(lookup_main_data)
    for main_row_index in range(0, main_table_len):
        main_row = main_table_data[main_row_index]
        key_value = main_row[match_column]
        match_rows = []
        match_row_index_list = []
        has_all_match = False
        for lookup_row_index in range(0, lookup_table_len):
            lookup_row = lookup_main_data[lookup_row_index]
            if lookup_row[match_column] == key_value:
                all_match = True
                for column_name in main_row:
                    if main_row[column_name] != lookup_row[column_name]:
                        all_match = False
                        break
                if all_match:
                    has_all_match = True
                else:
                    match_rows.append(lookup_row)
                match_row_index_list.append(lookup_row_index)
        if has_all_match:
            result.extend(match_rows)
        else:
            match_rows_count = len(match_rows)
            for position in range(0, match_rows_count):
                if position == 0:
                    main_table_data[main_row_index] = match_rows[position]
                else:
                    result.append(match_rows[position])

        for position in range(len(match_row_index_list), 0):
            lookup_main_data.remove(match_row_index_list[position - 1])
    main_table_data.extend(lookup_main_data)
    main_table_data.extend(result)

    pass


def merge_table_with_inserting_additional_matching_rows(
    main_table: ListObject, lookup_main: ListObject, match_column: str
) -> list:
    """
    Merge two tables with inserting additional matching rows.
    :param ListObject main_table:  (required)
    :param ListObject lookup_main:  (required)
    :param str match_column:  (required)
    :return list:
    """
    result = []
    main_table_data = list_object_to_list_dict(main_table)
    lookup_main_data = list_object_to_list_dict(lookup_main)
    key_dict = {}
    all_match_rows = {}
    main_table_len = len(main_table_data)
    lookup_table_len = len(lookup_main_data)
    for main_row_index in range(main_table_len, 0):
        main_row = main_table_data[main_row_index]
        key_value = main_row[match_column]
        match_rows = []
        match_row_index_list = []
        has_all_match = False
        for lookup_row_index in range(0, lookup_table_len):
            lookup_row = lookup_main_data[lookup_row_index]
            if lookup_row[match_column] == key_value:
                all_match = True
                for column_name in main_row:
                    if main_row[column_name] != lookup_row[column_name]:
                        all_match = False
                        break
                if all_match:
                    has_all_match = True
                else:
                    match_rows.append(lookup_row)
                match_row_index_list.append(lookup_row_index)
        if has_all_match:
            if main_row_index == len( main_table_data) -1 :
                main_table_data.extend(match_rows)
            else:
                main_table_data.insert(main_row_index + 1 , match_rows)
        else:
            match_rows_count = len(match_rows)
            for position in range(0, match_rows_count):
                if position == 0:
                    main_table_data[main_row_index] = match_rows[position]
                else:
                    result.insert(main_row_index + 1 ,match_rows[position])

        for position in range(len(match_row_index_list), 0):
            lookup_main_data.remove(match_row_index_list[position - 1])
    main_table_data.extend(lookup_main_data)
    pass


def __list_join_list(
    list1: list, list2: list, table1_field: str, table2_field: str
) -> list:
    table1_column_index = 0
    table2_column_index = 0
    column_index = 0
    list1_length = len(list1[0])
    list2_length = len(list2[0])
    for column_name in list1[0]:
        if column_name == table1_field:
            table1_column_index = column_index
        column_index = column_index + 1
    column_index = 0
    for column_name in list2[0]:
        if column_name == table1_field:
            table2_column_index = column_index
        else:
            list1[0].append(column_name)
        column_index = column_index + 1

    list_null = []
    for key in range(0, list2_length):
        list_null.append(None)

    for row in list1[1 : list1_length - 1]:
        field_value = row[table1_column_index]
        loop_break = False
        for table2_row in list2[1 : list2_length - 1]:
            field2_value = table2_row[table2_column_index]
            if field_value == field2_value:
                column_index = 0
                for column_value in table2_row:
                    if column_index == table2_column_index:
                        pass
                    else:
                        row.append(column_value)
                    column_index = column_index + 1
                loop_break = True
                break
        if loop_break:
            pass
        else:
            row.extend(list_null)
    return list1


def __dict_to_list__(
    dict_data: dict,
    row: list,
    result: list,
    cur_level: int,
    deep_level: int,
    value_map_column_list: list,
):
    if cur_level == deep_level:
        new_row = row.copy()
        for column in value_map_column_list:
            if column in dict_data:
                new_row.append(dict_data[column])
            else:
                new_row.append(0)
        result.append(new_row)
        pass
    else:
        for key in dict_data:
            new_row = row.copy()
            new_row.append(key)
            __dict_to_list__(
                dict_data[key],
                new_row,
                result,
                cur_level + 1,
                deep_level,
                value_map_column_list,
            )
    pass
