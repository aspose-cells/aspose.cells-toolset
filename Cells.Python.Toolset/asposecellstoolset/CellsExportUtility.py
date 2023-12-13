from __future__ import absolute_import
from aspose.cells import Workbook
from aspose.cells import Worksheet
from aspose.cells.tables import ListObject
from aspose.cells import CellsHelper
import numpy as np
import pandas as pd
import datetime

class CellsExportUtility(object):
    
    def __init__(self):
        self.sheet_index = None
        self.shape_index = None
        self.picture_index = None
        self.list_object_index = None
        self.pivot_table_index = None
        self.chart_index = None
        self.range_name = None
        pass

    def export_data(self, workbook , **kwargs):
        self.__init_parameters(**kwargs)

        if self.sheet_index is None:
            self.sheet_index =  workbook.worksheets.active_sheet_index
            
        if self.chart_index is not None:
            pass
        elif self.list_object_index is not None:
            return self.__export_list_object(workbook.worksheets[self.sheet_index],self.list_object_index)
            pass
        elif self.picture_index is not None:
            pass
        elif self.pivot_table_index is not  None:
            return  self.__export_pivot_table(workbook.worksheets[self.sheet_index],self.pivot_table_index)
            pass
        elif self.shape_index is not  None:
            pass 
        elif self.range_name is not  None:
            return self.__export_range(workbook.worksheets[self.sheet_index], self.range_name)
            pass

        return self.__export_worksheet(workbook.worksheets[self.sheet_index])    
        pass
    
    def export_worksheet(self, workbook : Workbook , sheet_index : int) -> list :
        return self.__export_worksheet(workbook.worksheets[sheet_index])
    
    def export_worksheet(self, worksheet : Worksheet) -> list :
        return self.__export_worksheet(worksheet)
    
    def export_list_object(self, workbook : Workbook , sheet_index : int, list_object_index : int ) -> list :
        return self.__export_list_object(workbook.worksheets[sheet_index],list_object_index)
    

    def __init_parameters(self, **kwargs):
        # parameter initialize
        if kwargs.get("sheet_index") is not None:
            self.sheet_index = kwargs.get("sheet_index")
            
        if kwargs.get("shape_index") is not None:
            self.shape_index = kwargs.get("shape_index")
        if kwargs.get("picture_index") is not None:
            self.picture_index = kwargs.get("picture_index")
        if kwargs.get("list_object_index") is not None:
            self.list_object_index = kwargs.get("list_object_index")
        if kwargs.get("pivot_table_index") is not None:
            self.pivot_table_index = kwargs.get("pivot_table_index")
        if kwargs.get("chart_index") is not None:
            self.chart_index = kwargs.get("chart_index")
        if kwargs.get("range_name") is not None:
            self.range_name = kwargs.get("range_name")
            
        pass    
    
    def __export_worksheet(self, worksheet ):
        max_row_index  = worksheet.cells.max_row
        max_column_index  = worksheet.cells.max_column
        table =[]
        for row_index in range(0,max_row_index):
            row  =[]            
            for column_index in range(0,max_column_index):
                row.append(  worksheet.cells.get(row_index,column_index).value)
            table.append(row)
        return table
    
    def __export_range(self, worksheet, range_name):
        pos = range_name.find(":")
        table = []
        
        temp_rows = []
        temp_columns = []
        if pos > -1 :            
            CellsHelper.cell_name_to_index(range_name[0:pos], temp_rows,temp_columns)
            begin_row_index = temp_rows[0]
            begin_column_index = temp_columns[0]
            CellsHelper.cell_name_to_index(range_name[pos+1:], temp_rows,temp_columns)
            end_row_index = temp_rows[0]
            end_column_index = temp_columns[0]
            for row_index in range(begin_row_index,end_row_index):
                row  =[]
                for column_index in range(begin_column_index, end_column_index):
                    row.append(  worksheet.cells.get(row_index,column_index).value)
                table.append(row)
            
        else:
            CellsHelper.cell_name_to_index(range_name, temp_rows,temp_columns)
            begin_row_index = temp_rows[0]
            begin_column_index = temp_columns[0]
            row  =[]
            table.append(row.append( worksheet.cells.get(begin_row_index,begin_column_index).value))
        return table

    def __export_list_object(self, worksheet ,list_object_index):
        list_object = worksheet.list_objects[list_object_index]
        table =[]
        for row_index in range(list_object.start_row , list_object.end_row):
            row  =[]            
            for column_index in range(list_object.start_column, list_object.end_column):
                row.append(  worksheet.cells.get(row_index,column_index).value)
            table.append(row)
        return table    
        pass
    
    def __export_pivot_table(self, worksheet ,pivot_table_index):
        pivot_table = worksheet.pivot_tables[pivot_table_index]
        cellarea = pivot_table.table_range2
        cells = worksheet.cells
        table =[]
        for row_index in range(cellarea.start_row , cellarea.end_row):
            row  =[]            
            for column_index in range(cellarea.start_column, cellarea.end_column):
                row.append(  cells.get(row_index,column_index).value)
            table.append(row)
        return table    
        pass
    
    def __export_chart(self, worksheet , chart_index):
        chart = worksheet.charts[chart_index]
        category_data = chart.n_series.category_data
        cells = worksheet.cells
        table =[]
        for row_index in range(cellarea.start_row , cellarea.end_row):
            row  =[]            
            for column_index in range(cellarea.start_column, cellarea.end_column):
                row.append(  cells.get(row_index,column_index).value)
            table.append(row)
        return table    
        pass
    