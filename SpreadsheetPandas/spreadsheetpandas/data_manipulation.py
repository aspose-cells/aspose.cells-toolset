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
    # 1. Get table data range and table column index 
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
    # 2. table to dict
    column_value_dict = {}
    for row_index in range( table_data_begin_row_index, table_data_end_row_index):
        cur_pivot_column_value = None
        cur_value_column_value = None
        cur_row = None
        IsFirstCell = True
        for column_index in range(table_data_begin_column_index ,table_data_end_column_index ):
            if column_index == pivot_column_index:
                cur_pivot_column_value = cells[row_index,column_index].value
                if cur_pivot_column_value not in column_value_dict :
                    column_value_dict[cur_pivot_column_value] = cur_pivot_column_value
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
    #3. dict to list 
    column_value_list = sorted( list(column_value_dict.keys())) 
    result =[]
    row = []
    __dict_to_list__(table_rows,row,result,0, len(ListObject.list_columns)-2,column_value_list )
    return result
  
def unpivot_column( table: ListObject ,column_names : list , column_map_name : str , value_map_name :str ) ->list:
    
    #1. get basic data.
    rows = [] 
    cells = table.data_range.worksheet.cells    
    column_name_map_column_index = {}
    cur_column_index = 0
    row_head = []
    for column in table.list_columns :
        if column.name in column_names :
            column_name_map_column_index[cur_column_index] = column.name
        else :
            row_head.append( column.name)
        cur_column_index = cur_column_index + 1
    row_head.append( column_map_name )
    row_head.append( value_map_name )
    rows.append(row_head)
    table_data_begin_row_index =  table.data_range.first_row 
    table_data_end_row_index =  table.data_range.first_row  +  table.data_range.row_count
    table_data_begin_column_index =  table.data_range.first_column  
    table_data_end_column_index =  table.data_range.first_column  +  table.data_range.column_count
    # 2. table to dict

    #2. 
    for row_index in range( table_data_begin_row_index, table_data_end_row_index): 
        row = []
        column_value_list = []
        for column_index in range(table_data_begin_column_index ,table_data_end_column_index ):
            if column_index in column_name_map_column_index :
                column_value_list.append ( [ column_name_map_column_index[column_index] , cells[row_index,column_index].value])
            else:
                row.append(cells[row_index,column_index].value)
        
        for fields in column_value_list:
            new_row = row.copy()
            new_row.append(fields[0])
            new_row.append(fields[1])
            rows.append(new_row)

    return rows

def __dict_to_list__(dict_data :dict, row :list, result : list, cur_level: int , deep_level :int, value_map_column_list :list  ) :    
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
        for  key in  dict_data :                
            new_row = row.copy()
            new_row.append (key)
            print(cur_level , new_row)
            __dict_to_list__(dict_data[key],new_row ,result, cur_level +1 ,deep_level,value_map_column_list)    
    pass

