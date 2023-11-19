from __future__ import absolute_import

import numpy as np
import os
import sys
import unittest
import warnings

ABSPATH = os.path.abspath(os.path.realpath(os.path.dirname(__file__)) + "/..")
sys.path.append(ABSPATH)

from asposecellstoolset import CellsImportUtility
from aspose.cells import Workbook 

class TestImportNumArrayData( unittest.TestCase):
    
    def setUp(self):
        warnings.simplefilter('ignore', ResourceWarning)
    
    def tearDown(self):
        pass
    
    def test_import_row_data_vertical(self):
        
        import_tool = CellsImportUtility()
        workbook = Workbook()
        data = np.array([1,2,3,4,5,6,7,8,9])
       
        import_tool.import_data_into_workbook( workbook ,data, is_vertical=True)
        workbook.save("D:/Cells.Toolset/TestData/Output/import_numarray_int_vertical.xlsx")
        pass
    
    def test_import_row_data_horizontal(self):
        
        import_tool = CellsImportUtility()
        workbook = Workbook()
        data = np.array([1,2,3,4,5,6,7,8,9])
        import_tool.import_data_into_workbook( workbook ,data, is_vertical=False)
        workbook.save("D:/Cells.Toolset/TestData/Output/import_numarray_int_horizontal.xlsx")
        pass    

    def test_import_table_data_horizontal(self):        
        import_tool = CellsImportUtility()
        workbook = Workbook()
        data = np.array([[1, 2, 3], [4, 5, 6], [7, 8, 9]])
        import_tool.import_data_into_workbook( workbook ,data, is_vertical=False)
        workbook.save("D:/Cells.Toolset/TestData/Output/import_numarray_int_table_horizontal.xlsx")
        pass        
    
    def test_import_table_data_vertical(self):        
        import_tool = CellsImportUtility()
        workbook = Workbook()
        data = np.array([[1, 2, 3], [4, 5, 6], [7, 8, 9]])
        import_tool.import_data_into_workbook( workbook ,data, is_vertical=True)
        workbook.save("D:/Cells.Toolset/TestData/Output/import_numarray_int_table_vertical.xlsx")
        pass        
    
    def test_import_tables_data_horizontal(self):        
        import_tool = CellsImportUtility()
        workbook = Workbook()
        data = np.array([[[1, 2, 3], [1, 2, 3], [1, 2, 3]], [[4, 5, 6], [4, 5, 6], [4, 5, 6]], [[7, 8, 9], [7, 8, 9], [7, 8, 9]]])
        import_tool.import_data_into_workbook( workbook ,data, is_vertical=False)
        workbook.save("D:/Cells.Toolset/TestData/Output/import_numarray_tables_horizontal.xlsx")
        pass  

    def test_import_tables_data_vertical(self):        
        import_tool = CellsImportUtility()
        workbook = Workbook()
        data = np.array([[[1, 2, 3], [1, 2, 3], [1, 2, 3]], [[4, 5, 6], [4, 5, 6], [4, 5, 6]], [[7, 8, 9], [7, 8, 9], [7, 8, 9]]])
        import_tool.import_data_into_workbook( workbook ,data, is_vertical=True)
        workbook.save("D:/Cells.Toolset/TestData/Output/import_numarray_tables_vertical.xlsx")
        
        pass  
    
    def test_import_tables_data_horizontal_one_sheet(self):        
        import_tool = CellsImportUtility()
        workbook = Workbook()
        data = np.array([[[1, 2, 3], [1, 2, 3], [1, 2, 3]], [[4, 5, 6], [4, 5, 6], [4, 5, 6]], [[7, 8, 9], [7, 8, 9], [7, 8, 9]]])
        import_tool.import_data_into_workbook( workbook ,data, is_vertical=False,one_sheet=True)
        workbook.save("D:/Cells.Toolset/TestData/Output/import_numarray_tables_horizontal_one_sheet.xlsx")
        pass  

    def test_import_tables_data_vertical_one_sheet(self):        
        import_tool = CellsImportUtility()
        workbook = Workbook()
        data = np.array([[[1, 2, 3], [1, 2, 3], [1, 2, 3]], [[4, 5, 6], [4, 5, 6], [4, 5, 6]], [[7, 8, 9], [7, 8, 9], [7, 8, 9]]])
        import_tool.import_data_into_workbook( workbook ,data, is_vertical=True,one_sheet=True)
        workbook.save("D:/Cells.Toolset/TestData/Output/import_numarray_tables_vertical_one_sheet.xlsx")
        
        pass  

if __name__ == '__main__':
    unittest.main()