from tkinter import ALL
from aspose.cells import Workbook
from aspose.cells import FileFormatType


class DataProcessing(object):
    
    def __init__(self):

        pass
     
    def data_cleansing(self , workbook : Workbook,  **kwargs):
        

            
        pass
    
    def data_deduplication(self , workbook : Workbook,**kwargs):
        
        entireSheet = workbook.file_format() in [ FileFormatType.CSV() , FileFormatType.TSV() , FileFormatType.HTML() ,FileFormatType.M_HTML()]
         
        for sheet in workbook.worksheets:
            if entireSheet :    
                sheet.cells.remove_duplicates()
            else:
                for namerange in workbook.worksheets.get_named_ranges() :
                    worksheet = workbook.worksheets.get(namerange.worksheet().name() );
                    worksheet.cells.remove_duplicates(namerange.FirstRow, namerange.FirstColumn, namerange.FirstRow + namerange.RowCount - 1, namerange.FirstColumn + namerange.ColumnCount - 1);
            