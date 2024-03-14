from __future__ import absolute_import

import os
import sys
import unittest
import warnings
import re



ABSPATH = os.path.abspath(os.path.realpath(os.path.dirname(__file__)) + "/..")
sys.path.append(ABSPATH)

from src.ExcelPandas import ExcelPandas
from aspose.cells import Workbook 
from aspose.cells import CellsHelper

class TestExcelPandas( unittest.TestCase):
    
    def setUp(self):
        warnings.simplefilter('ignore', ResourceWarning)
    
    def tearDown(self):
        pass
    
    
    def test_read_spreadsheet_listobject(self):
        excelPandas = ExcelPandas()
        df = excelPandas.read_spreadsheet("D:\cells-toolset\TestData\Input\BookChartData.xlsx", sheet_index = 0, list_object_index = 0)        
        pass
    
    def test_read_spreadsheet_sheet(self):
        excelPandas = ExcelPandas()
        df = excelPandas.read_spreadsheet("D:\cells-toolset\TestData\Input\BookChartData.xlsx", sheet_index = 2)
        pass   
    
    def test_read_spreadsheet_chart(self):
        excelPandas = ExcelPandas()
        df = excelPandas.read_spreadsheet("D:\cells-toolset\TestData\Input\BookChartData.xlsx", sheet_index = 1,chart_index = 0)
        print(df)
        pass 
        
if __name__ == '__main__':
    unittest.main()
