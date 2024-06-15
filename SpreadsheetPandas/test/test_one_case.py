from __future__ import absolute_import
import os
import sys
import unittest
import warnings
import re
import pandas as pd

ABSPATH = os.path.abspath(os.path.realpath(os.path.dirname(__file__)) + "/..")
sys.path.append(ABSPATH)

from spreadsheetpandas.spreadsheet_pandas import read_spreadsheet, write_spreadsheet
from spreadsheetpandas.data_conversion import *
from aspose.cells import Workbook 
from aspose.cells import Worksheet 
from aspose.cells import Cells
from aspose.cells import Cell 
from aspose.cells import CellsHelper



class TestOneCase( unittest.TestCase):
    
    def setUp(self):
        warnings.simplefilter('ignore', ResourceWarning)
    
    def tearDown(self):
        pass
    def test_read_spreadsheet_url(self):
        workbook = Workbook("../TestData/Input/BookTableData.xlsx")
        worksheet = workbook.worksheets.get("SaleTable")
        data =  list_object_to_list (worksheet.list_objects[0])
        print(data)
        pass        
    

if __name__ == '__main__':
    unittest.main()
