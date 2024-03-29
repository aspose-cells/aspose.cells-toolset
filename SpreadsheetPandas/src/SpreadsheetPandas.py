from __future__ import absolute_import
from pathlib import Path
import pandas as pd
import re
import io
from aspose.cells import Workbook
from aspose.cells import Worksheet
from aspose.cells import Cells
from aspose.cells.tables import ListObject
from aspose.cells import CellsHelper
from aspose.cells import CellValueType


class SpreadsheetPandas(object):
    
    def __init__(self):
        
        pass
    
        """
        read data form spreadsheet which is include of Excel, cvs, txt, ods, iCalc and so on.
        :param str path:  (required)
        :param int sheet_index: The worksheet index indicates the position in the spreadsheet. (optional)
        :param int list_object_index: The list object index indicates the position in the spreadsheet. If the worksheet index is None, the default worksheet index is the active worksheet index. (optional)
        :param int pivot_table_index: The worksheet index indicates the position in the spreadsheet. If the worksheet index is None, the default worksheet index is the active worksheet index. (optional)
        :param int chart_index: The worksheet index indicates the position in the spreadsheet. If the worksheet index is None, the default worksheet index is the active worksheet index. (optional)
        :param int cell_area: The worksheet cell_area indicates the position in the spreadsheet. If the worksheet index is None, the default worksheet index is the active worksheet index. (optional)
        :param int name_text: The workbook name indicates the position in the spreadsheet. If the worksheet index is None, the default worksheet index is the active worksheet index. (optional)
        :return DataFrame: 
        """                
    def read_spreadsheet( self , path: str , **kwargs )-> pd.DataFrame:
        workbook = Workbook(path)
        sheet_index = None
        if kwargs.get("sheet_index") is not None:
            sheet_index = kwargs.get("sheet_index")
        else:
            sheet_index =  workbook.worksheets.active_sheet_index
            
        list_object_index = None    
        if kwargs.get("list_object_index") is not None:
            list_object_index = kwargs.get("list_object_index")
        pivot_table_index = None    
        if kwargs.get("pivot_table_index") is not None:
            pivot_table_index = kwargs.get("pivot_table_index")
        chart_index = None
        if kwargs.get("chart_index") is not None:
            chart_index = kwargs.get("chart_index")
        cell_area = None
        if kwargs.get("cell_area") is not None:
            cell_area = kwargs.get("cell_area")        
        name_text = None
        if kwargs.get("name_text") is not None:
            name_text = kwargs.get("name_text") 
            
        if list_object_index is not None:
            worksheet = workbook.worksheets[sheet_index]
            table  = worksheet.list_objects[list_object_index]
            return self.__get_data_to_dataframe(worksheet.cells ,table.start_row,table.start_column,table.end_row,table.end_column,table.show_header_row,table.show_totals )
        
        if pivot_table_index is not None:
            worksheet = workbook.worksheets[sheet_index]
            pivot_table = worksheet.pivot_tables[pivot_table_index]
            cellarea = pivot_table.table_range2
            return self.__get_data_to_dataframe( worksheet.cells ,cellarea.start_row,cellarea.start_column,cellarea.end_row,cellarea.end_column,False,False )
        
        if chart_index is not None:
            return self.__get_chart_data_to_dataframe( workbook,sheet_index,chart_index )

        if cell_area is not None:
            tuple_cell_area = self.__parse_cell_area(cell_area)
            cells = workbook.worksheets[sheet_index].cells
            has_table_header = self.__has_table_header(cells,tuple_cell_area[0],tuple_cell_area[1],tuple_cell_area[2],tuple_cell_area[3])
            return self.__get_data_to_dataframe(cells,tuple_cell_area[0],tuple_cell_area[1],tuple_cell_area[2],tuple_cell_area[3],has_table_header,False )
        
        if name_text is not None:
            name_range = workbook.worksheets.get_range_by_name(name_text)
            
            cells = name_range.worksheet.cells
            begin_row_index = name_range.first_row
            begin_column_index = name_range.first_column
            end_row_index = name_range.first_row + name_range.row_count -1
            end_column_index = name_range.first_column + name_range.column_count -1            
            has_table_header = self.__has_table_header(cells,begin_row_index,begin_column_index,end_row_index,end_column_index)
            return self.__get_data_to_dataframe(cells,begin_row_index,begin_column_index,end_row_index,end_column_index,has_table_header,False )
        
        cells = workbook.worksheets[sheet_index].cells
        has_table_header = self.__has_table_header(cells,cells.min_data_row,cells.min_data_column,cells.max_data_row,cells.max_data_column)
        return self.__get_data_to_dataframe(cells,cells.min_data_row,cells.min_data_column,cells.max_data_row,cells.max_data_column,has_table_header,False )

        pass
    

        """
        write data form spreadsheet which is include of Excel, cvs, txt, ods, iCalc and so on.
        :param str path: If the file exists, it will be appended to the file. If the file does not exist, a new file is created and written.  (required)
        :param int sheet_index: The worksheet index indicates the position in the spreadsheet. (optional)
        :param int list_object_index: The list object index indicates the position in the spreadsheet. If the worksheet index is None, the default worksheet index is the active worksheet index. (optional)
        :param int pivot_table_index: The worksheet index indicates the position in the spreadsheet. If the worksheet index is None, the default worksheet index is the active worksheet index. (optional)
        :param int chart_index: The worksheet index indicates the position in the spreadsheet. If the worksheet index is None, the default worksheet index is the active worksheet index. (optional)
        :param int cell_area: The worksheet cell_area indicates the position in the spreadsheet. If the worksheet index is None, the default worksheet index is the active worksheet index. (optional)
        :param int name_text: The workbook name indicates the position in the spreadsheet. If the worksheet index is None, the default worksheet index is the active worksheet index. (optional)
        :return  
        """  
        def write_spreadsheet( self , path: str, data : pd.DataFrame , **kwargs ) :
            if Path.exists(path) :
                workbook = Workbook(path)
            else :
                workbook = Workbook()
            

            pass
    
    def __get_data_to_dataframe( self , cells : Cells , begin_row_index : int , begin_column_index : int , end_row_index : int , end_column_index : int , has_header: bool, has_total : bool)->pd.DataFrame:        
        column_title_list =[]
        row_index = 0
        cells_helper = CellsHelper
        if has_header :
            row_index = begin_row_index
        for column_index in range(begin_column_index , end_column_index + 1 ):
            if has_header :
                column_title_list.append (cells.get(row_index,column_index).display_string_value )
            else:
                column_title_list.append (cells_helper.column_index_to_name(column_index) )                       

        start_row = 0
        end_row = 0              
        if has_header :
            start_row = begin_row_index + 1
        else:
            start_row = begin_row_index
        
        if has_total:
            end_row = end_row_index 
        else:
            end_row = end_row_index + 1
                 
        position = 0
        data = {}
        for column_index in range(begin_column_index , end_column_index + 1 ):
            column_data = []
            for row_index in range(start_row ,end_row ):
                column_data.append(cells.get(row_index,column_index).value)
            data[column_title_list[position]] = column_data
            position = position + 1
            
        return pd.DataFrame(data)
    def __has_table_header(self, cells: Cells, begin_row_index :int, begin_column_index:int, end_row_index :int, end_column_index:int ):
        has_header = True
        for column_index in range(begin_column_index , end_column_index +1) :
            cell = cells.get(begin_row_index,column_index)
            if cell.type != CellValueType.IS_STRING :
                has_header = False
                break
            sen_cell = cells.get(begin_row_index+1,column_index)
            if cell.type != sen_cell.type :
                break
        return has_header
        pass        
    def __get_chart_data_to_dataframe( self , workbook : Workbook, sheet_index : int , chart_index : int )->pd.DataFrame:
        chart = workbook.worksheets[sheet_index].charts[chart_index]
        data = {}
        series = self.__parse_data_source( chart.n_series.category_data )
        cells = workbook.worksheets.get(series[0]).cells
        column_index  = series[2]
        column_data = []
        for row_index in range (series[1],series[3] +1):
            column_data.append(cells.get(row_index , column_index).value)
        
        xName = ""
        if cells.get(series[1] -1  , column_index).type == CellValueType.IS_NULL :
            xName = CellsHelper.column_index_to_name(series[1])
            
        else:
            xName = cells.get(series[1] -1  , column_index).value
            
        data[xName] = column_data
        yNames = []
        for index in range( 0,len (chart.n_series)):
            values = self.__parse_data_source( chart.n_series[index].values)
            values_data = []
            for row_index in range (values[1],values[3] +1):
                values_data.append(cells.get(row_index , column_index + index + 1).value)
            data[chart.n_series[index].display_name] = values_data
            yNames.append(chart.n_series[index].display_name)
        return pd.DataFrame(data)
            
    def __parse_data_source( self , value : str):        
        matchObj = re.match( r'^=(.*)!\$(.*)\$(\d+):\$(.*)\$(\d+)', value, re.M|re.I)
        if matchObj == None :
            return None
        
        return (matchObj.group(1) ,int( matchObj.group(3)) - 1 ,int( CellsHelper.column_name_to_index (matchObj.group(2))), int( matchObj.group(5)) -1 , int(  CellsHelper.column_name_to_index (matchObj.group(4))) )
            
    def __parse_cell_area( self , value : str):        
        matchObj = re.match( r'^([a-zA-Z]*)(\d+):([a-zA-Z]*)(\d+)', value, re.M|re.I)
        if matchObj == None :
            return None
        return (int( matchObj.group(2)) - 1 ,int( CellsHelper.column_name_to_index (matchObj.group(1))), int( matchObj.group(4)) -1 , int(  CellsHelper.column_name_to_index (matchObj.group(3))) )
