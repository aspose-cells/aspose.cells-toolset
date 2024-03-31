from __future__ import absolute_import

import os
import sys
import unittest
import warnings
import re
import pandas as pd


ABSPATH = os.path.abspath(os.path.realpath(os.path.dirname(__file__)) + "/..")
sys.path.append(ABSPATH)

from src.SpreadsheetPandas  import SpreadsheetPandas
from aspose.cells import Workbook 
from aspose.cells import CellsHelper

class TestSpreadsheetPandas( unittest.TestCase):
    
    def setUp(self):
        warnings.simplefilter('ignore', ResourceWarning)
    
    def tearDown(self):
        pass
    def test_write_spreadsheet_df(self):
        excelPandas = SpreadsheetPandas()
        data = pd.DataFrame( [['Google', 10], ['Runoob', 12], ['Wiki', 13]], columns=['Site', 'Age'])
        excelPandas.write_spreadsheet("D:\cells-toolset\TestData\Output\BookWriteTable.xlsx",data)    
        # print(df)
        pass
    def test_read_spreadsheet_csv(self):
        excelPandas = SpreadsheetPandas()
        df = excelPandas.read_spreadsheet("D:\cells-toolset\TestData\Input\BookTableData.csv")    
        assert df.shape == (20,5)
        # print(df)
        pass
    def test_read_spreadsheet_xlsx(self):
        excelPandas = SpreadsheetPandas()
        df = excelPandas.read_spreadsheet("D:\cells-toolset\TestData\Input\BookData.xlsx")    
        assert df.shape == (20,5)
        # print(df)
        pass
    def test_read_spreadsheet_name(self):
        excelPandas = SpreadsheetPandas()
        df = excelPandas.read_spreadsheet("D:\cells-toolset\TestData\Input\BookChartData.xlsx" , name_text = "RangeData")    
        assert df.shape == (20,5)

        pass    
    def test_read_spreadsheet_listobject(self):
        excelPandas = SpreadsheetPandas()
        df = excelPandas.read_spreadsheet("D:\cells-toolset\TestData\Input\BookChartData.xlsx", sheet_index = 0, list_object_index = 0)      
        assert df.shape == (20,5)
        # print(df)
        pass
    def test_read_spreadsheet_cellarea(self):
        excelPandas = SpreadsheetPandas()
        df = excelPandas.read_spreadsheet("D:\cells-toolset\TestData\Input\BookChartData.xlsx", sheet_index = 2 , cell_area="D15:H35")
        assert df.shape == (20,5)
        # print(df)
        pass      
    def test_read_spreadsheet_sheet(self):
        excelPandas = SpreadsheetPandas()
        df = excelPandas.read_spreadsheet("D:\cells-toolset\TestData\Input\BookChartData.xlsx", sheet_index = 2)
        assert df.shape == (20,5)
        pass   
    
    def test_read_spreadsheet_chart(self):
        excelPandas = SpreadsheetPandas()
        df = excelPandas.read_spreadsheet("D:\cells-toolset\TestData\Input\BookChartData.xlsx", sheet_index = 1,chart_index = 0)
        assert df.shape == (20,4)
        pass 
        
if __name__ == '__main__':
    unittest.main()
