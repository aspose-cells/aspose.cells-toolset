from __future__ import absolute_import

import os
import sys
import unittest
import warnings

from asposecellstoolset.CellsToolset import import_aggregate_data_into_file

ABSPATH = os.path.abspath(os.path.realpath(os.path.dirname(__file__)) + "/..")
sys.path.append(ABSPATH)

from asposecellstoolset import CellsImportUtility
from aspose.cells import Workbook 

class TestImportData( unittest.TestCase):
    
    def setUp(self):
        warnings.simplefilter('ignore', ResourceWarning)
    
    def tearDown(self):
        pass
    
    def test_import_row_data_vertical(self):        
        data = (1,2,3,4,5,6,7,8,9)       
        import_aggregate_data_into_file("D:/cells-toolset/TestData/Output/import_int_vertical.xlsx"  ,data, 0,0,0,is_vertical=True)
        pass
    
    def test_import_row_data_vertical_function(self):        
        import_tool = CellsImportUtility()
        workbook = Workbook()
        data = (1,2,3,4,5,6,7,8,9)       
        import_tool.import_data_into_workbook( workbook ,data, is_vertical=True)
        workbook.save("D:/cells-toolset/TestData/Output/import_int_vertical.xlsx")
        pass
    
    def test_import_row_data_horizontal(self):        
        import_tool = CellsImportUtility()
        workbook = Workbook()
        data = (1,2,3,4,5,6,7,8,9)
        import_tool.import_data_into_workbook( workbook ,data, is_vertical=False)
        workbook.save("D:/cells-toolset/TestData/Output/import_int_horizontal.xlsx")
        pass    
    
    def test_import_table_data_vertical(self):        
        import_tool = CellsImportUtility()
        workbook = Workbook()
        data = ((1, 2, 3), (4, 5, 6), (7, 8, 9))
        import_tool.import_data_into_workbook( workbook ,data, is_vertical=True)
        workbook.save("D:/cells-toolset/TestData/Output/import_int_table_vertical.xlsx")
        pass   
    def test_import_table_data_vertical(self):        
        import_tool = CellsImportUtility()
        data = ((1, 2, 3), (4, 5, 6), (7, 8, 9))
        import_aggregate_data_into_file("D:/cells-toolset/TestData/Output/import_int_vertical.xlsx"  ,data, 0,0,0,is_vertical=True)
        pass
    def test_import_table_data_horizontal(self):        
        import_tool = CellsImportUtility()
        workbook = Workbook()
        data = ((1, 2, 3), (4, 5, 6), (7, 8, 9))
        import_tool.import_data_into_workbook( workbook ,data, is_vertical=False)
        workbook.save("D:/cells-toolset/TestData/Output/import_int_table_horizontal.xlsx")
        pass        
     
    
    def test_import_row_list_data_vertical(self):        
        import_tool = CellsImportUtility()
        workbook = Workbook()
        data = [1,2,3,4,5,6,7,8,9]       
        import_tool.import_data_into_workbook( workbook ,data, is_vertical=True)
        workbook.save("D:/cells-toolset/TestData/Output/import_int_list_vertical.xlsx")
        pass
    
    def test_import_row_list_data_horizontal(self):        
        import_tool = CellsImportUtility()
        workbook = Workbook()
        data = [1,2,3,4,5,6,7,8,9]
        import_tool.import_data_into_workbook( workbook ,data, is_vertical=False)
        workbook.save("D:/cells-toolset/TestData/Output/import_int_list_table_horizontal.xlsx")
        pass    

    def test_import_row_list_data_horizontal_function(self):        
        data = [1,2,3,4,5,6,7,8,9]
        import_aggregate_data_into_file("D:/cells-toolset/TestData/Output/import_int_vertical.xlsx"  ,data, 0,0,0,is_vertical=True)
        pass    
    
    def test_import_list_data_vertical(self):        
        import_tool = CellsImportUtility()
        workbook = Workbook()
        data = [[1, 2, 3], [4, 5, 6], [7, 8, 9]]
        import_tool.import_data_into_workbook( workbook ,data, is_vertical=True)
        workbook.save("D:/cells-toolset/TestData/Output/import_int_list_table_vertical.xlsx")
        pass   
    
    def test_import_list_data_horizontal(self):        
        import_tool = CellsImportUtility()
        workbook = Workbook()
        data = [[1, 2, 3], [4, 5, 6], [7, 8, 9]]
        import_tool.import_data_into_workbook( workbook ,data, is_vertical=False)
        workbook.save("D:/cells-toolset/TestData/Output/import_int_list_table_horizontal.xlsx")
        pass             
    
    def test_import_dict_data_horizontal(self):        
        import_tool = CellsImportUtility()
        workbook = Workbook()
        data = {'name':'roy' ,'age': 19, 'Education':'university'}
        import_tool.import_data_into_workbook( workbook ,data, is_vertical=False)
        workbook.save("D:/cells-toolset/TestData/Output/import_dict_horizontal.xlsx")
        pass  

    def test_import_dict_data_vertical(self):        
        import_tool = CellsImportUtility()
        workbook = Workbook()
        data = {'name':'roy' ,'age': 19, 'Education':'university'}
        import_tool.import_data_into_workbook( workbook ,data, is_vertical=True)
        workbook.save("D:/cells-toolset/TestData/Output/import_dict_vertical.xlsx")
        
    def test_import_dict_data_vertical_function(self):        
        data = {'name':'roy' ,'age': 19, 'Education':'university'}
        import_aggregate_data_into_file("D:/cells-toolset/TestData/Output/import_dict_vertical.xlsx"  ,data, 0,0,0,is_vertical=True)
 
if __name__ == '__main__':
    unittest.main()