from tkinter import ALL
from aspose.cells import Workbook
from aspose.cells import FileFormatType
from aspose.cells import CellValueType

class DataProcessing(object):
    
    def __init__(self):

        pass
     
    def data_cleansing(self , workbook : Workbook,  **kwargs):
        
        self.data_deduplication(workbook ,kwargs)
        self.data_fill(workbook ,kwargs)
        
        pass
    
    def data_deduplication(self , workbook : Workbook,**kwargs):        
        entireSheet = workbook.file_format() in [ FileFormatType.CSV() , FileFormatType.TSV() , FileFormatType.HTML() ,FileFormatType.M_HTML()]
        if entireSheet :
            for sheet in workbook.worksheets:
                sheet.cells.remove_duplicates()
        else :            
            for namerange in workbook.worksheets.get_named_ranges() :
                worksheet = workbook.worksheets.get(namerange.worksheet.name() );
                worksheet.cells.remove_duplicates(namerange.first_row, namerange.first_column, namerange.first_row + namerange.row_count - 1, namerange.first_column + namerange.column_count - 1);
            for sheet in workbook.worksheets:
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
        
        pass
    
    def data_fill(self , workbook : Workbook,**kwargs):
        
        pass


                

            