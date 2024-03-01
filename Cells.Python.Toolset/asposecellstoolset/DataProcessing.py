from xmlrpc.client import Boolean
from aspose.cells import Workbook
from aspose.cells import Worksheet
from aspose.cells import FileFormatType
from aspose.cells import CellValueType

class DataProcessing(object):
    
    def __init__(self):

        pass
     
    def data_cleansing(self , workbook : Workbook, need_fill_data : bool , **kwargs):        
        self.data_deduplication(workbook ,kwargs)
        if need_fill_data :
            self.data_fill(workbook ,kwargs)
        
        pass
    
    def data_deduplication(self , workbook : Workbook,**kwargs):        
        entireSheet = workbook.file_format in [ FileFormatType.CSV , FileFormatType.TSV , FileFormatType.HTML ,FileFormatType.M_HTML ]
        if entireSheet :
            for sheet in workbook.worksheets:
                sheet.cells.remove_duplicates()
        else :            
            for namerange in workbook.worksheets.get_named_ranges() :
                worksheet = workbook.worksheets.get(namerange.worksheet.name() );
                worksheet.cells.remove_duplicates(namerange.first_row, namerange.first_column, namerange.first_row + namerange.row_count - 1, namerange.first_column + namerange.column_count - 1);
            for sheet in workbook.worksheets:
                if sheet.list_objects.count > 0 :
                    for table in sheet.list_objects:
                        sheet.cells.remove_duplicates(table.start_row, table.start_column, table.end_row, table.end_column)
                        for row in range(table.end_row , table.start_row):
                            needBreak = False
                            for column in range(table.start_column , table.end_column):
                                cell = sheet.cells.get(row , column)
                                if cell.type == CellValueType.IsNull :
                                    continue
                                else :
                                    needBreak = True
                                    break
                            if needBreak :
                                table.resize(table.start_row, table.start_column, row, table.end_column, table.show_header_row)
                                break
                else:
                    sheet.cells.remove_duplicates()
        pass
    
    def data_fill(self , workbook : Workbook, **kwargs):
        entireSheet = workbook.file_format in [ FileFormatType.CSV , FileFormatType.TSV , FileFormatType.HTML ,FileFormatType.M_HTML ]
        if entireSheet :
            for worksheet in workbook.worksheets:
                for column_index in range ( 0, worksheet.cells.max_data_column):
                    for row_index in range( 0 , worksheet.cells.max_data_row):                    
                        cell = worksheet.cells.get( row_index , column_index )
                        if cell.type == CellValueType.IS_NULL :
                            cell.put_value(0)                        
        else:
            for worksheet in workbook.worksheets:
                if worksheet.list_objects.count > 0 :
                    for table in worksheet.list_objects:
                        for column_index in range ( table.start_column , table.end_column ):
                            for row_index in range(table.end_row , table.start_row):
                                cell = worksheet.cells.get( row_index , column_index )
                                if cell.type == CellValueType.IsNull :
                                    cell.put_value(0)    
                else:
                    for column_index in range ( 0, worksheet.cells.max_data_column):
                        for row_index in range( 0 , worksheet.cells.max_data_row):                    
                            cell = worksheet.cells.get( row_index , column_index )
                            if cell.type == CellValueType.IS_NULL :
                                cell.put_value(0)       
        pass
    
    def get_column_main_data_type(self, worksheet :Worksheet , column_index : int ,**kwargs ):
        cell_type_set ={}
        
        for row_index in range(0, worksheet.cells.max_data_row):
            cell = worksheet.cells.get( row_index , column_index )
            value_type = cell.type 
            if value_type in [ CellValueType.IS_NULL ,CellValueType.IS_NULL , CellValueType.IS_NULL ] :
                continue
            else:
                if cell.type in cell_type_set :
                    count =  cell_type_set[cell.type] 
                    cell_type_set[cell.type] = count + 1
                else:
                    cell_type_set[cell.type] =  1
        count = 0
        last_cell_value_type =  CellValueType.IS_NULL
        for cell_value_type in cell_type_set:           
            if count < cell_value_type.value :
                count = cell_value_type.value
                last_cell_value_type = cell_value_type.key
        return last_cell_value_type
        pass 


                

            