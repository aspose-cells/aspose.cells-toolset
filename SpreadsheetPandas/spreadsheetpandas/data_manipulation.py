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
from aspose.cells import CellValueType
from aspose.cells import ProtectionType
from aspose.cells.tables import ListObject
import numpy as np
import pandas as pd

## cells object to python object
def pivot_column( table: ListObject , pivot_column: str , value_column:str , aggregation: str) ->list :
    """
    List Object 
    :param ListObject table:  (required)
    :param str pivot_column:  (required)
    :param str value_column:  (required)
    :param str aggregation:  (required)
    :return list: 
    """

    cells = table.data_range.worksheet.cells
    pivot_column_index = 0
    value_column_index = 0
    column_index_map_name = {}
    column_name_map_index = {}
    table_rows = {} 
    column_index = 0
    for column in table.list_columns:
        if column.name == pivot_column :
            pivot_column_index = column_index
        if column.name == value_column :
            value_column_index = column_index
        column_index_map_name[column_index] = column.name
        column_name_map_index[column.name] = column_index
        column_index = column_index  + 1
    table_data_begin_row_index =  table.data_range.first_row 
    table_data_end_row_index =  table.data_range.first_row  +  table.data_range.row_count
    table_data_begin_column_index =  table.data_range.first_column  
    table_data_end_column_index =  table.data_range.first_column  +  table.data_range.column_count

    for row_index in range( table_data_begin_row_index, table_data_end_row_index):
        cur_pivot_column_value = None
        cur_value_column_value = None
        cur_row = None
        IsFirstCell = True
        for column_index in range(table_data_begin_column_index ,table_data_end_column_index ):
            if column_index == pivot_column_index:
                cur_pivot_column_value = cells[row_index,column_index].value
            elif column_index == value_column_index :
                cur_value_column_value = cells[row_index,column_index].value
            else:        
                cell_value = cells[row_index,column_index].value
                if IsFirstCell == True :                           
                    if  cell_value in  table_rows:
                        cur_row = table_rows[cell_value]
                    else:
                        table_rows[cell_value] ={}
                        cur_row = table_rows[cell_value]
                    IsFirstCell = False
                else:
                    if cell_value in cur_row :
                        cur_row = cur_row[cell_value]
                    else :
                        cur_row[cell_value] ={}
                        cur_row = cur_row[cell_value]                             
                
        cur_row[cur_pivot_column_value] = cur_value_column_value
    return table_rows
  
    
