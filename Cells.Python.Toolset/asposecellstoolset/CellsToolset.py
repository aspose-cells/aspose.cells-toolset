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
    
    """
    if os.path.exists(path):
        workbook = Workbook(path)
    else:
        workbook = Workbook()
    
    import_tool = CellsImportUtility()
    import_tool.import_data_into_workbook(workbook,data,**kwargs)
    
    workbook.save(path)
    
    pass    

def export_worksheet_data(path : str , sheet_index : int) -> list:
    workbook = Workbook(path)
    export_tool = CellsExportUtility()    
    return export_tool.export_data( workbook ,sheet_index=sheet_index )
    pass

def export_list_object_data(path : str , sheet_index : int, list_object_index : int) -> list:
    workbook = Workbook(path)
    export_tool = CellsExportUtility()    
    return export_tool.export_data( workbook ,sheet_index=sheet_index,list_object_index=list_object_index)
    pass
def export_pivot_table_data(path : str , sheet_index : int, pivot_table_index : int) -> list:
    workbook = Workbook(path)
    export_tool = CellsExportUtility()    
    return export_tool.export_data( workbook ,sheet_index=sheet_index,pivot_table_index=pivot_table_index)
    pass

def export_range_data(path : str , sheet_index : int, range_name : str) -> list:
    
    workbook = Workbook(path)
    export_tool = CellsExportUtility()
    return export_tool.export_data( workbook ,sheet_index=sheet_index, range_name=range_name)
    pass
