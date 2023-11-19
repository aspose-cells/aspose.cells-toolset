from __future__ import absolute_import
from aspose.cells import Workbook
import numpy as np
import pandas as pd
import datetime

class CellsImportUtility(object):
    
    def __init__(self):
        self.sheet_index = None
        self.is_vertical = False
        self.row_index = 0
        self.column_index = 0
        self.one_sheet = False
        pass
    
    def import_data_into_workbook(self , workbook, data , **kwargs):
        self.__init_parameters(**kwargs)
        #
        if self.sheet_index is None:
            self.sheet_index = workbook.worksheets.active_sheet_index        
        row = self.row_index
        column = self.column_index                
        cells = workbook.worksheets[self.sheet_index].cells        
        #
        dtype = type(data)
        
        if dtype is np.ndarray :                   
            self.__import_ndarray_into_workbook( cells ,data, row, column, self.is_vertical)
        elif dtype is tuple or dtype is list or dtype is set :
            self.__import_data_into_workbook( cells ,data, row, column, self.is_vertical)
        elif dtype is dict :
            self.__import_dict_data_into_workbook(cells, data, row, column, self.is_vertical )
        elif dtype is pd.DataFrame:
            self.__import_dataframe_into_workbook(cells, data, row, column, self.is_vertical )
        else:
            self.__put_value_to_cell(cells, data, row, column)
        pass
    
    def __init_parameters(self, **kwargs):
        # parameter initialize
        if kwargs.get("sheet_index") is not None:
            self.sheet_index = kwargs.get("sheet_index")
        if kwargs.get("is_vertical") is not None:
            self.is_vertical = kwargs.get("is_vertical")
        if kwargs.get("row_index") is not None:
            self.row_index = kwargs.get("row_index")
        if kwargs.get("column_index") is not None:
            self.column_index = kwargs.get("column_index")
        if kwargs.get("one_sheet") is not None:
            self.one_sheet = kwargs.get("one_sheet")
            
        pass
    
    def __import_dict_data_into_workbook(self, cells, data, row_index, column_index, is_vertical):
        for key in data:
            if is_vertical:
                self.__put_value_to_cell(cells,key,row_index,column_index)
                self.__put_value_to_cell(cells,data[key],row_index,column_index + 1)
                row_index = row_index + 1
            else:
                self.__put_value_to_cell(cells,key,row_index,column_index)
                self.__put_value_to_cell(cells,data[key],row_index + 1,column_index)
                column_index = column_index + 1
        pass
    
    def __is_tablix(self, data):
        if (type(data) is tuple) or (type(data) is list) :
            for val in data:
                if (type(val) is tuple) or (type(val) is list) :
                    return True
        return False
        pass
    
    def __import_data_into_workbook(self, cells,data,row_index,column_index,is_vertical):
        if self.__is_tablix (data):
            self.__import_table_data_into_workbook(cells,data,row_index,column_index,is_vertical)
        else:
            for val in data:
               self.__put_value_to_cell(cells,val,row_index,column_index)
               if is_vertical:
                   row_index = row_index + 1 
               else:
                   column_index = column_index  + 1
        pass

    def __import_table_data_into_workbook(self, cells,table_data,row_index,column_index,is_vertical):
        table_row_index = row_index
        table_column_index = column_index
        for table_row in table_data:
            for table_cell in table_row:
               self.__put_value_to_cell(cells,table_cell,table_row_index,table_column_index)
               if is_vertical :
                   table_row_index = table_row_index + 1
               else:
                   table_column_index = table_column_index + 1
            if is_vertical :
                table_row_index = row_index
                table_column_index = table_column_index + 1
            else:
                table_column_index = column_index
                table_row_index = table_row_index + 1                   
        pass
    
    def __import_dataframe_into_workbook(self, cells, data, row_index, column_index, is_vertical):
        df_row_index = row_index
        df_column_index = column_index
        for column_name in data.columns:
            df_row_index = row_index
            self.__put_value_to_cell(cells, column_name, df_row_index ,df_column_index + 1)
            for df_value in data[column_name]:
                self.__put_value_to_cell(cells, df_value, df_row_index + 1 ,df_column_index + 1)
                df_row_index = df_row_index + 1
            df_column_index = df_column_index + 1                
        df_row_index = row_index
        df_column_index = column_index
        for df_row_name in  data.index.values:
            self.__put_value_to_cell(cells, df_row_name, df_row_index + 1 ,df_column_index)
            df_row_index = df_row_index + 1
        pass

    def __import_ndarray_into_workbook(self, cells, data, row_index, column_index, is_vertical):
        
        if data.ndim == 1 :
            if self.is_vertical :                
                self.__import_ndarray_data_into_column( cells ,data, row_index,column_index)
            else:
                self.__import_ndarray_data_into_row( cells ,data, row_index,column_index)
            pass
        elif data.ndim == 2 :            
            self.__import_ndarray_data_into_table(cells,data,row_index,column_index,is_vertical)
            pass
        elif data.ndim == 3 :
            if self.one_sheet :
                new_data = self.__reshape(data,self.one_sheet)
                self.__import_ndarray_data_into_table(cells,new_data,row_index,column_index,is_vertical)
                pass
            else:
                dim_index = 1
                for sheet_data in data:
                    sheet_row_index = row_index
                    sheet_column_index = column_index
                    self.__import_ndarray_data_into_table(cells,sheet_data,sheet_column_index,sheet_column_index,is_vertical)
                    dim_index = dim_index + 1
                    if(dim_index <= 3):
                        sheet_index = cells.first_cell.worksheet.workbook.worksheets.add()
                        cells = cells.first_cell.worksheet.workbook.worksheets[sheet_index].cells
            pass
        else:
            new_data = self.__reshape(data,self.one_sheet)
            if self.one_sheet :                
                self.__import_ndarray_data_into_table(cells,new_data,row_index,column_index,is_vertical)
                pass 
            else:
                dim_index = 1
                for sheet_data in new_data:
                    sheet_row_index = row_index
                    sheet_column_index = column_index
                    self.__import_ndarray_data_into_table(cells,sheet_data,sheet_column_index,sheet_column_index,is_vertical)
                    dim_index = dim_index + 1
                    if(dim_index <= 3):
                        sheet_index = cells.first_cell.worksheet.workbook.worksheets.add()
                        cells = cells.first_cell.worksheet.workbook.worksheets[sheet_index].cells 
                pass                 
        pass
    
    def __reshape(self,data,one_sheet):
        new_dim_numbers = 1
        if one_sheet:
            end_dim_numbers = data.shape[data.ndim - 1]
            dim_count = data.ndim - 1
            for pos in range(dim_count):
                new_dim_numbers = new_dim_numbers * data.shape[pos]
            new_data  = data.reshape(new_dim_numbers,end_dim_numbers)            
        else:
            end1_dim_numbers = data.shape[data.ndim - 1]
            end2_dim_numbers = data.shape[data.ndim - 2]
            dim_count = data.ndim - 2
            for pos in range(dim_count):
                new_dim_numbers = new_dim_numbers * data.shape[pos]
            new_data  = data.reshape(new_dim_numbers,end2_dim_numbers,end1_dim_numbers)    
        return new_data

    def __import_ndarray_data_into_table(self, cells,data,row_index,column_index,is_vertical):
        for row_data in data :
            if is_vertical:
                self.__import_ndarray_data_into_column( cells ,row_data, row_index,column_index)
                column_index = column_index + 1
            else:
                self.__import_ndarray_data_into_row( cells ,row_data, row_index,column_index)
                row_index = row_index + 1       
        pass
                    
    def __import_ndarray_data_into_row(self, cells, data, row , column  ):
        for val in data:
            self.__put_value_to_cell(cells,val,row,column)   
            column = column + 1                
        pass

    def __import_ndarray_data_into_column(self, cells, data, row , column ):
        for val in data:
            self.__put_value_to_cell(cells,val,row,column)            
            row = row + 1
        pass
    
    def __put_value_to_cell(self , cells, raw_value, row , column):
        cell = cells.get(row , column)
        dtype = type(raw_value)
        match dtype:
            case np.bool_ :
                value = bool(raw_value)
            case np.int_ :
                value = int(raw_value)
            case np.intc :
                value = int(raw_value)
            case np.intp :
                value = int(raw_value)
            case np.int8 :
                value = int(raw_value)
            case np.int16 :
                value = int(raw_value)
            case np.int32 :
                value = int(raw_value)
            case np.int64 :
                value = int(raw_value)
            case np.uint8 :
                value = int(raw_value)
            case np.uint16 :
                value = int(raw_value)
            case np.uint32 :
                value = int(raw_value)
            case np.uint64 :
                value = int(raw_value)
            # case np.byte :
            #     value = byte(raw_value)
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
            case np.datetime64 :
                ts = pd.to_datetime(str(raw_value))
                value = ts.strftime('%Y.%m.%d')
            case _:
                 value = raw_value
        # if dtype is np.int32  :            
        #     value = int(raw_value)
        # elif dtype is np.float128:
        #     value = float(raw_value)
        # elif dtype is np.datetime64:
        #     ts = pd.to_datetime(str(raw_value))
        #     value = ts.strftime('%Y.%m.%d')
        # else:
        #     value = raw_value
        cell.put_value(value)
        pass 
