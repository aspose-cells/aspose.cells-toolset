from __future__ import absolute_import

import os
import sys
import unittest
import warnings
import re
import numpy as np
import pandas as pd

ABSPATH = os.path.abspath(os.path.realpath(os.path.dirname(__file__)) + "/..")
sys.path.append(ABSPATH)

from asposecellstoolset.DataConversion import DataConversion
from aspose.cells import Workbook 
from aspose.cells import License 
from aspose.cells import CellsHelper

class TestDataConversion( unittest.TestCase):
    
    def setUp(self):
        warnings.simplefilter('ignore', ResourceWarning)
        license = License()
        license.set_license("D:\cells.cloud\src\Aspose.Cells.Cloud.MicroService\Aspose.Total.lic")
    def tearDown(self):
        pass

    def test_data_conversion_table2dataframe(self):
        workbook = Workbook("d:\cells-toolset\TestData\Input\BookTableData.xlsx")
        worksheet = workbook.worksheets[0]       
        table = worksheet.list_objects[0]
        data_conversion = DataConversion()  
        df = data_conversion.listobject_to_dataframe(table)
        assert df.shape == (20,5)
        # for column_name in df.columns:            
        #     print( column_name)
        #     for column_value in df[column_name]:
        #         print(column_value)
        pass
    def test_data_conversion_sheet2dataframe(self):
        workbook = Workbook("d:\cells-toolset\TestData\Input\BookTableData.xlsx")
        worksheet = workbook.worksheets[1]
        data_conversion = DataConversion()  
        df = data_conversion.worksheet_to_dataframe(worksheet)
        assert df.shape == (20,5)
        pass    
    def test_data_conversion_range2dataframe(self):
        workbook = Workbook("d:\cells-toolset\TestData\Input\BookTableData.xlsx")
        range_name= workbook.worksheets.get_range_by_name("RangeData")
        data_conversion = DataConversion()  
        df = data_conversion.range_to_dataframe(range_name)
        assert df.shape == (20,5)            
        pass        
    def test_data_conversion_dataframe2range(self):
        df = pd.DataFrame(
            {
                "A": 1.0,
                "B": pd.Timestamp("20130102"),
                "C": pd.Series(1, index=list(range(4)), dtype="float32"),
                "D": np.array([3] * 4, dtype="int32"),
                "E": pd.Categorical(["test", "train", "test", "train"]),
                "F": "foo",
            })
        workbook = Workbook()
        worksheet = workbook.worksheets[0]
        cells = worksheet.cells
        data_conversion = DataConversion()  
        range_data = data_conversion.dataframe_to_range(df ,cells,3,2)
        assert range_data.first_row == 3
        assert range_data.first_column == 2
        assert range_data.row_count == 5
        assert range_data.column_count == 6
        workbook.save("d:\cells-toolset\TestData\Output\dataframe2range.xlsx")
        pass    
    def test_data_conversion_dataframe2table(self):
        df = pd.DataFrame(
            {
                "A": 1.0,
                "B": pd.Timestamp("20130102"),
                "C": pd.Series(1, index=list(range(4)), dtype="float32"),
                "D": np.array([3] * 4, dtype="int32"),
                "E": pd.Categorical(["test", "train", "test", "train"]),
                "F": "foo",
            })
        workbook = Workbook()
        worksheet = workbook.worksheets[0]
        cells = worksheet.cells
        data_conversion = DataConversion()  
        table_data = data_conversion.dataframe_to_listobject(df ,cells,3,2)

        workbook.save("d:\cells-toolset\TestData\Output\dataframe2table.xlsx")
    
if __name__ == '__main__':
    unittest.main()
