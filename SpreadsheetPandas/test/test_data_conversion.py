
from __future__ import absolute_import

import os
import sys
import unittest
import warnings
import re
from numpy import ndarray
import pandas as pd

ABSPATH = os.path.abspath(os.path.realpath(os.path.dirname(__file__)) + "/..")
sys.path.append(ABSPATH)
from spreadsheetpandas.data_conversion import *

from aspose.cells import Workbook 
from aspose.cells import CellsHelper

class TestDataConversion( unittest.TestCase):
    
    def setUp(self):
        warnings.simplefilter('ignore', ResourceWarning)
    
    def tearDown(self):
        pass

    def test_list_to_worksheet(self ):
        list_data = [[1,2,3],[4,5,6],[7,8,9]]        
        workbook = Workbook("D:\cells-toolset\TestData\Input\BookTableData.xlsx")
        worksheet = workbook.worksheets[0];
        list_to_worksheet(list_data,worksheet)
        workbook.save("D:\\cells-toolset\\TestData\\Output\\test_list_to_worksheet.xlsx");
        pass
    
    def test_tuple_to_worksheet(self ):
        tuple_data = ((1,2,3),(4,5,6),(7,8,9))   
        workbook = Workbook("D:\cells-toolset\TestData\Input\BookTableData.xlsx")
        worksheet = workbook.worksheets[0];
        tuple_to_worksheet(tuple_data,worksheet)
        workbook.save("D:\\cells-toolset\\TestData\\Output\\test_tuple_to_worksheet.xlsx");
        pass    
    
    def test_ndarray_to_worksheet(self ):
        ndarray_data =np.array([[1,2,3],[4,5,6],[7,8,9]] )   
        workbook = Workbook("D:\cells-toolset\TestData\Input\BookTableData.xlsx")
        worksheet = workbook.worksheets[0];
        ndarray_to_worksheet(ndarray_data,worksheet)
        workbook.save("D:\\cells-toolset\\TestData\\Output\\test_ndarray_to_worksheet.xlsx");
        pass    
    
    def test_dataframe_to_worksheet(self ):
        dataframe_data =pd.DataFrame([[1,2,3],[4,5,6],[7,8,9]] , columns=['A', 'B','C'])   
        workbook = Workbook("D:\cells-toolset\TestData\Input\BookTableData.xlsx")
        worksheet = workbook.worksheets[0];
        dataframe_to_worksheet(dataframe_data,worksheet)
        workbook.save("D:\\cells-toolset\\TestData\\Output\\test_dataframe_to_worksheet.xlsx");
        pass        
    
    def test_list_to_list_object(self ):
        list_data = [[1,2,3],[4,5,6],[7,8,9]]        
        workbook = Workbook("D:\cells-toolset\TestData\Input\BookTableData.xlsx")
        index = workbook.worksheets.add();
        worksheet = workbook.worksheets[index]
        list_to_list_object(list_data,worksheet)
        workbook.save("D:\\cells-toolset\\TestData\\Output\\test_list_to_list_object.xlsx");
        pass
    
    def test_tuple_to_list_object(self ):
        tuple_data = ((1,2,3),(4,5,6),(7,8,9))
        workbook = Workbook("D:\cells-toolset\TestData\Input\BookTableData.xlsx")
        index = workbook.worksheets.add();
        worksheet = workbook.worksheets[index]
        list_to_list_object(tuple_data,worksheet)
        workbook.save("D:\\cells-toolset\\TestData\\Output\\test_tuple_to_list_object.xlsx");
        pass    
    
    def test_ndarray_to_list_object(self ):
        ndarray_data =np.array([[1,2,3],[4,5,6],[7,8,9]] )      
        workbook = Workbook("D:\cells-toolset\TestData\Input\BookTableData.xlsx")
        index = workbook.worksheets.add();
        worksheet = workbook.worksheets[index]
        ndarray_to_list_object(ndarray_data,worksheet)
        workbook.save("D:\\cells-toolset\\TestData\\Output\\test_ndarray_to_list_object.xlsx");
        pass
    
    def test_dataframe_to_list_object(self ):
        dataframe_data = pd.DataFrame([[1,2,3],[4,5,6],[7,8,9]] )      
        workbook = Workbook("D:\cells-toolset\TestData\Input\BookTableData.xlsx")
        index = workbook.worksheets.add();
        worksheet = workbook.worksheets[index]
        dataframe_to_list_object(dataframe_data,worksheet)
        workbook.save("D:\\cells-toolset\\TestData\\Output\\test_list_to_list_object.xlsx");
        pass
    
    def test_list_to_range(self ):
        list_data = [[1,2,3],[4,5,6],[7,8,9]]        
        workbook = Workbook("D:\cells-toolset\TestData\Input\BookTableData.xlsx")
        index = workbook.worksheets.add();
        worksheet = workbook.worksheets[index]
        list_to_range(list_data,worksheet)
        workbook.save("D:\\cells-toolset\\TestData\\Output\\test_list_to_range.xlsx");
        pass
    
    def test_tuple_to_range(self ):
        tuple_data = ((1,2,3),(4,5,6),(7,8,9))
        workbook = Workbook("D:\cells-toolset\TestData\Input\BookTableData.xlsx")
        index = workbook.worksheets.add();
        worksheet = workbook.worksheets[index]
        list_to_range(tuple_data,worksheet)
        workbook.save("D:\\cells-toolset\\TestData\\Output\\test_tuple_to_range.xlsx");
        pass    
    
    def test_ndarray_to_range(self ):
        ndarray_data =np.array([[1,2,3],[4,5,6],[7,8,9]] )      
        workbook = Workbook("D:\cells-toolset\TestData\Input\BookTableData.xlsx")
        index = workbook.worksheets.add();
        worksheet = workbook.worksheets[index]
        ndarray_to_range(ndarray_data,worksheet)
        workbook.save("D:\\cells-toolset\\TestData\\Output\\test_ndarray_to_range.xlsx");
        pass
    
    def test_dataframe_to_range(self ):
        dataframe_data = pd.DataFrame([[1,2,3],[4,5,6],[7,8,9]] )      
        workbook = Workbook("D:\cells-toolset\TestData\Input\BookTableData.xlsx")
        index = workbook.worksheets.add();
        worksheet = workbook.worksheets[index]
        dataframe_to_range(dataframe_data,worksheet)
        workbook.save("D:\\cells-toolset\\TestData\\Output\\test_dataframe_to_range.xlsx");
        pass    
    
if __name__ == '__main__':
    unittest.main()
