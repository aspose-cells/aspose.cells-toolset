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
from aspose.cells import Workbook 
from aspose.cells import Worksheet 
from aspose.cells import Cells
from aspose.cells import Cell 
from aspose.cells import CellsHelper



class TestBaseInfo( unittest.TestCase):
    
    def setUp(self):
        warnings.simplefilter('ignore', ResourceWarning)
    
    def tearDown(self):
        pass
    def test_base_info_table(self):
        workbook = Workbook("/home/cells/Projects/cells-toolset/TestData/Input/BookTableData.xlsx")
        print("===============")
        list_object = workbook.worksheets[0].list_objects[0]
        print (list_object.data_range.first_row)
        print (list_object.end_row)
        for row_index in range(list_object.data_range.first_row ,list_object.end_row ) :
            print(row_index)
        print("===============")
        pass        

if __name__ == '__main__':
    unittest.main()
