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
    
    def test_read_spreadsheet_csv(self):
        excelPandas = ExcelPandas()
        df = excelPandas.read_spreadsheet("D:\cells-toolset\TestData\Input\BookTableData.csv")    
        assert df.shape == (20,5)
        # print(df)
        pass
    def test_read_spreadsheet_xlsx(self):
        excelPandas = ExcelPandas()
        df = excelPandas.read_spreadsheet("D:\cells-toolset\TestData\Input\BookData.xlsx")    
        assert df.shape == (20,5)
        # print(df)
        pass
    def test_read_spreadsheet_name(self):
        excelPandas = ExcelPandas()
        df = excelPandas.read_spreadsheet("D:\cells-toolset\TestData\Input\BookChartData.xlsx" , name_text = "RangeData")    
        assert df.shape == (20,5)

        pass    
    def test_read_spreadsheet_listobject(self):
        excelPandas = ExcelPandas()
        df = excelPandas.read_spreadsheet("D:\cells-toolset\TestData\Input\BookChartData.xlsx", sheet_index = 0, list_object_index = 0)      
        assert df.shape == (20,5)
        # print(df)
        pass
    def test_read_spreadsheet_cellarea(self):
        excelPandas = ExcelPandas()
        df = excelPandas.read_spreadsheet("D:\cells-toolset\TestData\Input\BookChartData.xlsx", sheet_index = 2 , cell_area="D15:H35")
        assert df.shape == (20,5)
        # print(df)
        pass      
    def test_read_spreadsheet_sheet(self):
        excelPandas = ExcelPandas()
        df = excelPandas.read_spreadsheet("D:\cells-toolset\TestData\Input\BookChartData.xlsx", sheet_index = 2)
        assert df.shape == (20,5)
        pass   
    
    def test_read_spreadsheet_chart(self):
        excelPandas = ExcelPandas()
        df = excelPandas.read_spreadsheet("D:\cells-toolset\TestData\Input\BookChartData.xlsx", sheet_index = 1,chart_index = 0)
        assert df.shape == (20,4)
        pass 
        
if __name__ == '__main__':
    unittest.main()
