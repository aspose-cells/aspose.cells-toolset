from aspose.cells import Workbook
from aspose.cells import FileFormat

class DataProcessing(object):
    
    def __init__(self):

        pass
     
    def data_cleansing(self , workbook : Workbook,  **kwargs):
        for  sheet in workbook.worksheets:
            sheet.cells.RemoveDuplicates()
            
        pass
