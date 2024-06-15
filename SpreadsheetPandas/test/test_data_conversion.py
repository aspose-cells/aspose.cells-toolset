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
        workbook = Workbook("../TestData/Input/BookTableData.xlsx")
        worksheet = workbook.worksheets[0];
        list_to_worksheet(list_data,worksheet)
        workbook.save("../TestData/Output/test_list_to_worksheet.xlsx");
        pass
    
    def test_tuple_to_worksheet(self ):
        tuple_data = ((1,2,3),(4,5,6),(7,8,9))   
        workbook = Workbook("../TestData/Input/BookTableData.xlsx")
        worksheet = workbook.worksheets[0];
        tuple_to_worksheet(tuple_data,worksheet)
        workbook.save("../TestData/Output/test_tuple_to_worksheet.xlsx");
        pass    
    
    def test_ndarray_to_worksheet(self ):
        ndarray_data =np.array([[1,2,3],[4,5,6],[7,8,9]] )   
        workbook = Workbook("../TestData/Input/BookTableData.xlsx")
        worksheet = workbook.worksheets[0];
        ndarray_to_worksheet(ndarray_data,worksheet)
        workbook.save("../TestData/Output/test_ndarray_to_worksheet.xlsx");
        pass    
    
    def test_dataframe_to_worksheet(self ):
        dataframe_data =pd.DataFrame([[1,2,3],[4,5,6],[7,8,9]] , columns=['A', 'B','C'])   
        workbook = Workbook("../TestData/Input/BookTableData.xlsx")
        worksheet = workbook.worksheets[0];
        dataframe_to_worksheet(dataframe_data,worksheet)
        workbook.save("../TestData/Output/test_dataframe_to_worksheet.xlsx");
        pass        
    
    def test_list_to_list_object(self ):
        list_data = [[1,2,3],[4,5,6],[7,8,9]]        
        workbook = Workbook("../TestData/Input/BookTableData.xlsx")
        index = workbook.worksheets.add();
        worksheet = workbook.worksheets[index]
        list_to_list_object(list_data,worksheet)
        workbook.save("../TestData/Output/test_list_to_list_object.xlsx");
        pass
    
    def test_tuple_to_list_object(self ):
        tuple_data = ((1,2,3),(4,5,6),(7,8,9))
        workbook = Workbook("../TestData/Input/BookTableData.xlsx")
        index = workbook.worksheets.add();
        worksheet = workbook.worksheets[index]
        list_to_list_object(tuple_data,worksheet)
        workbook.save("../TestData/Output/test_tuple_to_list_object.xlsx");
        pass    
    
    def test_ndarray_to_list_object(self ):
        ndarray_data =np.array([[1,2,3],[4,5,6],[7,8,9]] )      
        workbook = Workbook("../TestData/Input/BookTableData.xlsx")
        index = workbook.worksheets.add();
        worksheet = workbook.worksheets[index]
        ndarray_to_list_object(ndarray_data,worksheet)
        workbook.save("../TestData/Output/test_ndarray_to_list_object.xlsx");
        pass
    
    def test_dataframe_to_list_object(self ):
        dataframe_data = pd.DataFrame([[1,2,3],[4,5,6],[7,8,9]] )      
        workbook = Workbook("../TestData/Input/BookTableData.xlsx")
        index = workbook.worksheets.add();
        worksheet = workbook.worksheets[index]
        dataframe_to_list_object(dataframe_data,worksheet)
        workbook.save("../TestData/Output/test_list_to_list_object.xlsx");
        pass
    
    def test_list_to_range(self ):
        list_data = [[1,2,3],[4,5,6],[7,8,9]]        
        workbook = Workbook("../TestData/Input/BookTableData.xlsx")
        index = workbook.worksheets.add();
        worksheet = workbook.worksheets[index]
        list_to_range(list_data,worksheet)
        workbook.save("../TestData/Output/test_list_to_range.xlsx");
        pass
    
    def test_tuple_to_range(self ):
        tuple_data = ((1,2,3),(4,5,6),(7,8,9))
        workbook = Workbook("../TestData/Input/BookTableData.xlsx")
        index = workbook.worksheets.add();
        worksheet = workbook.worksheets[index]
        list_to_range(tuple_data,worksheet)
        workbook.save("../TestData/Output/test_tuple_to_range.xlsx");
        pass    
    
    def test_ndarray_to_range(self ):
        ndarray_data =np.array([[1,2,3],[4,5,6],[7,8,9]] )      
        workbook = Workbook("../TestData/Input/BookTableData.xlsx")
        index = workbook.worksheets.add();
        worksheet = workbook.worksheets[index]
        ndarray_to_range(ndarray_data,worksheet)
        workbook.save("../TestData/Output/test_ndarray_to_range.xlsx");
        pass
    
    def test_dataframe_to_range(self ):
        dataframe_data = pd.DataFrame([[1,2,3],[4,5,6],[7,8,9]] )      
        workbook = Workbook("../TestData/Input/BookTableData.xlsx")
        index = workbook.worksheets.add();
        worksheet = workbook.worksheets[index]
        dataframe_to_range(dataframe_data,worksheet)
        workbook.save("../TestData/Output/test_dataframe_to_range.xlsx");
        pass   

    def test_list_to_name(self ):
        list_data = [[1,2,3],[4,5,6],[7,8,9]]        
        workbook = Workbook("../TestData/Input/BookTableData.xlsx")
        index = workbook.worksheets.add();
        worksheet = workbook.worksheets[index]
        list_to_name(list_data,worksheet)
        workbook.save("../TestData/Output/test_list_to_name.xlsx");
        pass
    
    def test_tuple_to_name(self ):
        tuple_data = ((1,2,3),(4,5,6),(7,8,9))
        workbook = Workbook("../TestData/Input/BookTableData.xlsx")
        index = workbook.worksheets.add();
        worksheet = workbook.worksheets[index]
        list_to_name(tuple_data,worksheet)
        workbook.save("../TestData/Output/test_tuple_to_name.xlsx");
        pass    
    
    def test_ndarray_to_name(self ):
        ndarray_data =np.array([[1,2,3],[4,5,6],[7,8,9]] )      
        workbook = Workbook("../TestData/Input/BookTableData.xlsx")
        index = workbook.worksheets.add();
        worksheet = workbook.worksheets[index]
        ndarray_to_name(ndarray_data,worksheet)
        workbook.save("../TestData/Output/test_ndarray_to_name.xlsx");
        pass
    
    def test_dataframe_to_name(self ):
        dataframe_data = pd.DataFrame([[1,2,3],[4,5,6],[7,8,9]] )      
        workbook = Workbook("../TestData/Input/BookTableData.xlsx")
        index = workbook.worksheets.add();
        worksheet = workbook.worksheets[index]
        dataframe_to_name(dataframe_data,worksheet)
        workbook.save("../TestData/Output/test_dataframe_to_name.xlsx");
        pass    

    def test_worksheet_to_list(self ):
        workbook = Workbook("../TestData/Input/BookTableData.xlsx")
        worksheet = workbook.worksheets.get("SaleSheet")
        data =  worksheet_to_list (worksheet)
        assert len(data) ==25
        pass    

    def test_worksheet_to_tuple(self ):
        workbook = Workbook("../TestData/Input/BookTableData.xlsx")
        worksheet = workbook.worksheets.get("SaleSheet")
        data =  worksheet_to_tuple (worksheet)
        assert len(data) ==25
        pass    
    def test_worksheet_to_ndarry(self ):
        workbook = Workbook("../TestData/Input/BookTableData.xlsx")
        worksheet = workbook.worksheets.get("SaleSheet")
        data =  worksheet_to_ndarray (worksheet)
        assert data.shape == (25,6)
        pass    
    
    def test_worksheet_to_dataframe(self ):
        workbook = Workbook("../TestData/Input/BookTableData.xlsx")
        worksheet = workbook.worksheets.get("SaleSheet")
        data =  worksheet_to_dataframe (worksheet)        
        assert data.shape == (26,7)
        pass    
    
    def test_list_object_to_list(self ):
        workbook = Workbook("../TestData/Input/BookTableData.xlsx")
        worksheet = workbook.worksheets.get("SaleTable")
        data =  list_object_to_list (worksheet.list_objects[0])
        assert len(data) ==21
        pass    

    def test_list_object_to_tuple(self ):
        workbook = Workbook("../TestData/Input/BookTableData.xlsx")
        worksheet = workbook.worksheets.get("SaleTable")
        data =  list_object_to_tuple (worksheet.list_objects[0])
        assert len(data) ==21
        pass    
    def test_list_object_to_ndarry(self ):
        workbook = Workbook("../TestData/Input/BookTableData.xlsx")
        worksheet = workbook.worksheets.get("SaleTable")
        data =  list_object_to_ndarray (worksheet.list_objects[0])
        assert data.shape == (20,5)
        pass    
    def test_list_object_to_dataframe(self ):
        workbook = Workbook("../TestData/Input/BookTableData.xlsx")
        worksheet = workbook.worksheets.get("SaleTable")
        data =  list_object_to_dataframe (worksheet.list_objects[0])
        assert data.shape == (20,5)
        pass    
    
    def test_range_to_list(self ):
        workbook = Workbook("../TestData/Input/BookTableData.xlsx")
        range_name = workbook.worksheets[2].cells.create_range("D15", "H35");
        data =  range_to_list( range_name )
        assert len(data) == 21
        pass    

    def test_range_to_tuple(self ):
        workbook = Workbook("../TestData/Input/BookTableData.xlsx")
        range_name = workbook.worksheets.get("SaleRange").cells.create_range("D15", "H35");
        data =  range_to_tuple (range_name)
        assert len(data) ==21
        pass    
    def test_range_to_ndarry(self ):
        workbook = Workbook("../TestData/Input/BookTableData.xlsx")
        range_name = workbook.worksheets.get("SaleRange").cells.create_range("D16", "H35");
        data =  range_to_ndarray (range_name)
        assert data.shape == (20,5)
        pass    
    def test_range_to_dataframe(self ):
        workbook = Workbook("../TestData/Input/BookTableData.xlsx")
        range_name = workbook.worksheets.get("SaleRange").cells.create_range("D15", "H35");
        data =  range_to_dataframe (range_name)
        assert data.shape == (20,5)
        pass    

    def test_name_to_list(self ):
        workbook = Workbook("../TestData/Input/BookTableData.xlsx")
        name = workbook.worksheets.names.get("RangeData");
        data =  name_to_list( name )
        assert len(data) == 21
        pass    

    def test_name_to_tuple(self ):
        workbook = Workbook("../TestData/Input/BookTableData.xlsx")
        name = workbook.worksheets.names.get("RangeData");
        data =  name_to_tuple (name)
        assert len(data) ==21
        pass    
    def test_name_to_ndarry(self ):
        workbook = Workbook("../TestData/Input/BookTableData.xlsx")
        name = workbook.worksheets.names.get("RangeData");
        data =  name_to_ndarray (name)
        assert data.shape == (21,5)
        pass    
    def test_name_to_dataframe(self ):
        workbook = Workbook("../TestData/Input/BookTableData.xlsx")
        name = workbook.worksheets.names.get("RangeData");
        data =  name_to_dataframe (name)
        assert data.shape == (20,5)
        pass  

if __name__ == '__main__':
    unittest.main()
