from __future__ import absolute_import
from ast import Import
import numpy as np
import pandas as pd
import os
import sys
import unittest
import warnings

ABSPATH = os.path.abspath(os.path.realpath(os.path.dirname(__file__)) + "/..")
sys.path.append(ABSPATH)


from asposecellstoolset.CellsExportUtility import CellsExportUtility
from asposecellstoolset.CellsToolset import *
from aspose.cells import Workbook 
from aspose.cells import CellsHelper
class TestExportData( unittest.TestCase):
    
    def setUp(self):
        warnings.simplefilter('ignore', ResourceWarning)
    
    def tearDown(self):
        pass
    
    def test_export_range(self):       
        print("test_export_range")
        exported_data =export_range_data( "D:\Cells.Toolset\TestData\Input\ExportData.xlsx" , 0 ,"C5:H18")
        print( np.array( exported_data))
        print( pd.DataFrame(exported_data))
        pass
    
    def test_export_list_object(self):       
        print("test_export_list_object")
        exported_data =export_list_object_data( "D:\Cells.Toolset\TestData\Input\ExportData.xlsx" , 1, 0 )
        print( np.array( exported_data))
        print( pd.DataFrame(exported_data))
        pass
    
    def test_export_pivot_table(self):       
        print("test_export_pivot_table")
        exported_data =export_pivot_table_data( "D:\Cells.Toolset\TestData\Input\ExportData.xlsx" , 3, 0 )
        print( np.array( exported_data))
        print( pd.DataFrame(exported_data))
        pass

    def test_export_data(self):   
        print("test_export_data")
        exported_data =export_worksheet_data( "D:\Cells.Toolset\TestData\Input\ExportData.xlsx" , 0 )
        print( np.array( exported_data))
        print( pd.DataFrame(exported_data))
        pass
        
if __name__ == '__main__':
    unittest.main()
