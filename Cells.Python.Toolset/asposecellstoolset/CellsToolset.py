from tkinter import W
from aspose.cells import Workbook
from asposecellstoolset.CellsImportUtility import CellsImportUtility
from asposecellstoolset.CellsExportUtility import CellsExportUtility
import numpy as np
import pandas as pd
import datetime
import os




def import_data_into_file( path : str , data ,  **kwargs):
    """
    mport data into a file.
    :param str path: The file path. (required)
    :param any data: Imported data, which can be in the form of ndarray, tuple, list, set, DataFrame, dict, or other data. (required)
    :param int sheet_index: The worksheet index indicating the position in the imported data workbook. The default value is active sheet index. (optional)
    :param int row_index: The row index of worksheet indicating the position in the imported data workbook. The default value is 0.(optional)
    :param int column_index: The column index of worksheet indicating the position in the imported data workbook. The default value is 0.(optional)
    :param bool is_vertical: Indicate whether the data is inserted vertically. The default value is False. (optional)
    :param bool one_sheet: Indicates whether the data is inserted into a table. "The default value is False."(optional)
    :return: 
    """
    if os.path.exists(path):
        workbook = Workbook(path)
    else:
        workbook = Workbook()
    
    import_tool = CellsImportUtility()
    import_tool.import_data_into_workbook(workbook,data,**kwargs)
    
    workbook.save(path)
    
    pass    

def import_ndarray_into_file( path :str ,data : np.ndarray , sheet_index : int ,row_index :int, column_index : int ,is_vertical : bool, one_sheet:bool):
    """
    Import ndarray data into a workbook.
    :param str path: The file path. (required)
    :param ndarray data: Imported ndarray data. (required)
    :param int sheet_index: The worksheet index indicating the position in the imported data workbook. The default value is active sheet index. (required)
    :param int row_index: The row index of worksheet indicating the position in the imported data workbook. The default value is 0.(required)
    :param int column_index: The column index of worksheet indicating the position in the imported data workbook. The default value is 0.(required)
    :param bool is_vertical: Indicate whether the data is inserted vertically. The default value is False. (required)
    :param bool one_sheet: Indicates whether the data is inserted into a table. "The default value is False."(required)
    :return: 
    """    
    if os.path.exists(path):
        workbook = Workbook(path)
    else:
        workbook = Workbook()
    
    import_tool = CellsImportUtility()    
    import_tool.import_ndarray_into_workbook( workbook.worksheets[sheet_index].cells,data,row_index,column_index,is_vertical )    
    workbook.save(path)

def import_dict_into_file( path :str ,data : dict , sheet_index : int ,row_index :int, column_index : int ,is_vertical : bool):
    """
    Import dict data into a workbook.
    :param str path: The file path. (required)
    :param dict data: Imported dict data. (required)
    :param int sheet_index: The worksheet index indicating the position in the imported data workbook. The default value is active sheet index. (required)
    :param int row_index: The row index of worksheet indicating the position in the imported data workbook. The default value is 0.(required)
    :param int column_index: The column index of worksheet indicating the position in the imported data workbook. The default value is 0.(required)
    :param bool is_vertical: Indicate whether the data is inserted vertically. The default value is False. (required)
    :return: 
    """    
    if os.path.exists(path):
        workbook = Workbook(path)
    else:
        workbook = Workbook()
    
    import_tool = CellsImportUtility()    
    import_tool.import_dict_into_workbook( workbook.worksheets[sheet_index].cells,data,row_index,column_index,is_vertical )    
    workbook.save(path)

def import_dataframe_into_file ( path :str ,data : pd.DataFrame , sheet_index : int ,row_index :int, column_index : int ,is_vertical : bool):
    """
    Import dataframe data into a workbook.
    :param str path: The file path. (required)
    :param dataframe data: Imported dataframe data. (required)
    :param int sheet_index: The worksheet index indicating the position in the imported data workbook. The default value is active sheet index. (required)
    :param int row_index: The row index of worksheet indicating the position in the imported data workbook. The default value is 0.(required)
    :param int column_index: The column index of worksheet indicating the position in the imported data workbook. The default value is 0.(required)
    :param bool is_vertical: Indicate whether the data is inserted vertically. The default value is False. (required)
    :return: 
    """    
    if os.path.exists(path):
        workbook = Workbook(path)
    else:
        workbook = Workbook()   
    import_tool = CellsImportUtility()   
    import_tool.import_dataframe_into_workbook( workbook.worksheets[sheet_index].cells,data,row_index,column_index,is_vertical )    
    workbook.save(path)    

def import_aggregate_data_into_file ( path :str ,data :list or set or tuple , sheet_index : int ,row_index :int, column_index : int ,is_vertical : bool):
    """
    Import dataframe data into a workbook.
    :param str path: The file path. (required)
    :param dataframe data: Imported dataframe data. (required)
    :param int sheet_index: The worksheet index indicating the position in the imported data workbook. The default value is active sheet index. (required)
    :param int row_index: The row index of worksheet indicating the position in the imported data workbook. The default value is 0.(required)
    :param int column_index: The column index of worksheet indicating the position in the imported data workbook. The default value is 0.(required)
    :param bool is_vertical: Indicate whether the data is inserted vertically. The default value is False. (required)
    :return: 
    """    
    if os.path.exists(path):
        workbook = Workbook(path)
    else:
        workbook = Workbook()   
    import_tool = CellsImportUtility()   
    import_tool.import_aggregate_data_into_workbook( workbook.worksheets[sheet_index].cells,data,row_index,column_index,is_vertical )    
    workbook.save(path)   

def export_worksheet_data(path : str , sheet_index : int) -> list:
    """
    Export worksheet data from a file.
    :param str path: The file path. (required)
    :param int sheet_index: The worksheet index indicates the position in the exported data workbook.  (required)
    :return list data: 
    """    
    workbook = Workbook(path)
    export_tool = CellsExportUtility()    
    return export_tool.export_data( workbook ,sheet_index=sheet_index )
    pass

def export_list_object_data(path : str , sheet_index : int, list_object_index : int) -> list:
    """
    Export list object data from a file.
    :param str path: The file path. (required)
    :param int sheet_index: The worksheet index indicates the position in the exported data workbook.  (required)
    :param int list_object_index: The list object index indicates the position in the exported data workbook.  (required)
    :return list data: 
    """    
    workbook = Workbook(path)
    export_tool = CellsExportUtility()    
    return export_tool.export_data( workbook ,sheet_index=sheet_index,list_object_index=list_object_index)
    pass

def export_list_object_data_to_dataframe(path : str , sheet_index : int, list_object_index : int) -> pd.DataFrame :
    """
    Export list object data from a file.
    :param str path: The file path. (required)
    :param int sheet_index: The worksheet index indicates the position in the exported data workbook.  (required)
    :param int list_object_index: The list object index indicates the position in the exported data workbook.  (required)
    :return list data: 
    """    
    workbook = Workbook(path)
    export_tool = CellsExportUtility()    
    return pd.DataFrame( export_tool.export_data( workbook ,sheet_index=sheet_index,list_object_index=list_object_index))

    pass
def export_pivot_table_data(path : str , sheet_index : int, pivot_table_index : int) -> list:
    """
    Export pivot table data from a file.
    :param str path: The file path. (required)
    :param int sheet_index: The worksheet index indicates the position in the exported data workbook.  (required)
    :param int pivot_table_index: The pivot table index indicates the position in the exported data workbook.  (required)
    :return list data: 
    """    
    workbook = Workbook(path)
    export_tool = CellsExportUtility()    
    return export_tool.export_data( workbook ,sheet_index=sheet_index,pivot_table_index=pivot_table_index)
    pass

def export_range_data(path : str , sheet_index : int, range_name : str) -> list:
    """
    Export range data from a file.
    :param str path: The file path. (required)
    :param int sheet_index: The worksheet index indicates the position in the exported data workbook.  (required)
    :param str range_name: The range_name indicates the position in the exported data workbook.  (required)
    :return list data: 
    """      
    workbook = Workbook(path)
    export_tool = CellsExportUtility()
    return export_tool.export_data( workbook ,sheet_index=sheet_index, range_name=range_name)
    pass
