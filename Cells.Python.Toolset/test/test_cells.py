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


from aspose.cells import Workbook 
from aspose.cells import CellsHelper

class TestCells( unittest.TestCase):
    
    def setUp(self):
        warnings.simplefilter('ignore', ResourceWarning)
    
    def tearDown(self):
        pass
    
    def test_print_datafre(self):
        data = [['Google', 10], ['Runoob', 12], ['Wiki', 13]]
        print(data)
        data_frame = pd.DataFrame(data, columns=['Site', 'Age'])
        print(data_frame["Site"].dtype )
        print(data_frame["Age"].dtype )
        print(data_frame)
        for column_name in data_frame.columns:
           print (column_name)
           for df_value in data_frame[column_name]:
               print(df_value)

        pass
    # def test_cells_helper(self):       
    #     row_index = []
    #     column_index = []
    #     CellsHelper.cell_name_to_index("C10",row_index,column_index)
    #     print(row_index , column_index)
    #     range_name ="C10:D100"
    #     pos = range_name.find(":")
    #     print(pos,range_name[0:pos], range_name[pos+1:] )
    #     range_name ="C10"
    #     pos = range_name.find(":")
    #     print(pos)
    #     pass
    # def test_cells_pivottable(self):       
    #     workbook = Workbook("d:\cells-toolset\TestData\Input\ExportData.xlsx")
    #     print( workbook.worksheets[3].pivot_tables[0].table_range1)
    #     print( workbook.worksheets[3].pivot_tables[0].table_range2)
    #     pass
    # def test_cells_chart(self):       
    #     print( "test_cells_chart")
    #     workbook = Workbook("d:\cells-toolset\TestData\Input\ExportData.xlsx")
    #     datarange =  workbook.worksheets[5].charts[0].n_series.category_data
    #     # "=Chart2!$C$5:$F$150"
    #     re.findall("=(\w)$()$()")
    #     # print( workbook.worksheets[5].charts[0].n_series.category_data)
    #     # print( workbook.worksheets[5].charts[0].n_series[0].area)
    #     # print( workbook.worksheets[2].charts[0].n_series.category_data)
    #     # print(len( workbook.worksheets[2].charts[0].n_series))
    #     pass
        
        
if __name__ == '__main__':
    unittest.main()
