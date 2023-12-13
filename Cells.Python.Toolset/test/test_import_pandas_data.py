from __future__ import absolute_import

import numpy as np
import pandas as pd
import os
import sys
import unittest
import warnings

from asposecellstoolset.CellsToolset import import_dataframe_into_file

ABSPATH = os.path.abspath(os.path.realpath(os.path.dirname(__file__)) + "/..")
sys.path.append(ABSPATH)

from asposecellstoolset import CellsImportUtility
from aspose.cells import Workbook 

class TestImportPandasData( unittest.TestCase):
    
    def setUp(self):
        warnings.simplefilter('ignore', ResourceWarning)
    
    def tearDown(self):
        pass
    
    def test_import_dataframe_data_vertical(self):       
        dates = pd.date_range("20130101", periods=6)
        df = pd.DataFrame(np.random.randn(6, 4), index=dates, columns=list("ABCD"))
        import_tool = CellsImportUtility()
        workbook = Workbook() 
        import_tool.import_data_into_workbook( workbook ,df, is_vertical=True)
        workbook.save("D:/cells-toolset/TestData/Output/import_dataframe_vertical.xlsx")
        pass
    
    def test_import_dataframe_data_vertical(self):       
        dates = pd.date_range("20130101", periods=6)
        df = pd.DataFrame(np.random.randn(6, 4), index=dates, columns=list("ABCD"))       
        import_dataframe_into_file( "D:/cells-toolset/TestData/Output/import_dataframe_vertical.xlsx" ,df, 0,0,0,True)
        pass
    
if __name__ == '__main__':
    unittest.main()