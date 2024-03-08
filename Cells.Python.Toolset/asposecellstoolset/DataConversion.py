import pandas as pd
from aspose.cells.tables import ListObject
from aspose.cells import CellsHelper
from aspose.cells import CellValueType
from aspose.cells import Cells
from aspose.cells import Range
from aspose.cells import Worksheet

class DataConversion(object):
    
    def __init__(self):

        pass

    def listobject_to_dataframe( self , table : ListObject ):        
        return self.__get_dataframe(table.data_range.worksheet.cells ,table.start_row,table.start_column,table.end_row,table.end_column,table.show_header_row,table.show_totals )
        pass
    
    def range_to_dataframe( self , range_name : Range):
        begin_row_index = range_name.first_row
        begin_column_index = range_name.first_column
        end_row_index = range_name.first_row + range_name.row_count -1
        end_column_index = range_name.first_column + range_name.column_count -1
        cells = range_name.worksheet.cells
        has_header = self.__has_table_header(cells,begin_row_index,begin_column_index,end_row_index,end_column_index)
        return self.__get_dataframe(cells ,begin_row_index,begin_column_index,end_row_index,end_column_index,has_header,False )    
        pass
    
    def worksheet_to_dataframe( self , worksheet : Worksheet):
        cells = worksheet.cells
        begin_row_index = cells.min_data_row
        begin_column_index = cells.min_data_column
        end_row_index = cells.max_data_row 
        end_column_index = cells.max_data_column         
        has_header = self.__has_table_header(cells,begin_row_index,begin_column_index,end_row_index,end_column_index)
        return self.__get_dataframe(cells ,begin_row_index,begin_column_index,end_row_index,end_column_index,has_header,False )    
        pass

    def dataframe_to_listobject( self ,dataframe: pd.DataFrame, cells: Cells , first_row : int , first_column: int ):
        cells_area = self.__dataframe_import_cells(dataframe ,cells,first_row ,first_column)
        return cells.first_cell.worksheet.list_objects.add(cells_area[0], cells_area[1], cells_area[0] + cells_area[2] - 1, cells_area[1] + cells_area[3] -1 ,True) 
        pass

    def dataframe_to_range(self,dataframe: pd.DataFrame , cells: Cells , first_row : int , first_column: int):
        cells_area = self.__dataframe_import_cells(dataframe ,cells,first_row ,first_column)
        return cells.create_range(cells_area[0],cells_area[1],cells_area[2],cells_area[3])        
        pass
    
    def dataframe_to_worksheet(self,dataframe: pd.DataFrame , cells: Cells , first_row : int , first_column: int):
        cells_area = self.__dataframe_import_cells(dataframe ,cells,first_row ,first_column)
        return cells.first_cell.worksheet       
        pass
    
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
        pass
    
    def __dataframe_import_cells(self,dataframe: pd.DataFrame , cells: Cells , first_row : int , first_column: int):
        column_count = 0
        row_count = 0
        column_index = first_column        
        for column_name in dataframe.columns:  
            column_count = column_count + 1
            row_index = first_row
            cell = cells.get(row_index , column_index )
            cell.put_value(column_name)
            
            row_index = row_index + 1
            for column_value in dataframe[column_name]:
                cell = cells.get(row_index , column_index )
                cell.put_value(column_value)
                row_index = row_index + 1
            column_index = column_index +1
            row_count = row_index - first_row
        return (first_row,first_column,row_count, column_count)
        pass
    
    def __get_dataframe(self , cells : Cells , begin_row_index : int , begin_column_index : int , end_row_index : int , end_column_index : int , has_header: bool, has_total : bool):
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
        pass


        # column_title_list =[]
        # row_index = 0
        # begin_row_index = 0
        # end_row_index = 0
        # cells_helper = CellsHelper
        # cells = table.data_range.worksheet.cells
        # if table.show_header_row :
        #     row_index = table.start_row
        # for column_index in range(table.start_column , table.end_column +1):
        #     if table.show_header_row :
        #         column_title_list.append (cells.get(row_index,column_index).display_string_value )
        #     else:
        #         column_title_list.append (cells_helper.column_index_to_name(column_index) )                       
              
        # if table.show_header_row :
        #     begin_row_index = table.start_row + 1
        # else:
        #     begin_row_index = table.start_row
        
        # if table.show_totals:
        #     end_row_index = table.end_row - 1
        # else:
        #     end_row_index = table.end_row
            
        # position = 0
        # data = {}
        # for column_index in range(table.start_column , table.end_column +1):
        #     column_data = []
        #     for row_index in range(begin_row_index,end_row_index):
        #         column_data.append(cells.get(row_index,column_index).value)
        #     data[column_title_list[position]] = column_data
        #     position = position + 1
        # return pd.DataFrame(data)