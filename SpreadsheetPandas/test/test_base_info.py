from __future__ import absolute_import
import os
import sys
import unittest
import warnings
import re
import pandas as pd

ABSPATH = os.path.abspath(os.path.realpath(os.path.dirname(__file__)) + "/..")
sys.path.append(ABSPATH)

from spreadsheetpandas.spreadsheet_pandas import read_spreadsheet, write_spreadsheet
from aspose.cells import Workbook 
from aspose.cells import Worksheet 
from aspose.cells import Cells
from aspose.cells import Cell 
from aspose.cells import CellsHelper



class TestBaseInfo( unittest.TestCase):
    
    def setUp(self):
        warnings.simplefilter('ignore', ResourceWarning)
    
    def tearDown(self):
        pass
    def test_dict_out(self):
        table = {'up': {'iphone': {'2012': 122, '2013': 122}, 'ipad': {'2012': 122, '2013': 120}}, 'down': {'iphone': {'2012': 122, '2013': 122}, 'ipad': {'2012': 122, '2013': 120}}}
        row =[]
        new_table = []
        value_map_column = {"2011":0,"2012":1, "2013":2}
        value_map_column_list = ["2011","2012", "2013"]
        self._sl_(table, row,new_table,0,2,value_map_column)
        print(new_table)
        # for firstcell in table:
        #     for  second in table[firstcell]:
        #         for third in table[firstcell][second]:
        #             row = [firstcell,second,]

    def _sl_(self , dict_data :dict, row :list, result : list, cur_level: int , deep_level :int , value_map_column_list :list ) ->list:
        if cur_level == deep_level:
            new_row = row.copy()
            for column in value_map_column_list:
                if column in dict_data:
                    new_row.append(dict_data[column])
                else:
                    new_row.append(0)
            # for  key in  dict_data :
            #     new_row.append(dict_data[key])
            result.append(new_row)
            pass
        else:
            for  key in  dict_data :                
                new_row = row.copy()
                new_row.append (key)
                print(cur_level , new_row)
                self._sl_(dict_data[key],new_row ,result, cur_level +1 ,deep_level,value_map_column_list)

            

    # def test_list_unique(self):
    #     header =["A","Product","Year", "Sales" ]
    #     source = [["up","iphone","2012",122],["up","ipad","2012",122],["up","iphone","2013",122] ,["up","ipad","2013",120],["down","iphone","2012",122],["down","ipad","2012",122],["down","iphone","2013",122] ,["down","ipad","2013",120]]
    #     source.insert(0,header)
    #     print( source)
    #     pass
    # def test_base_info_table(self):
    #     workbook = Workbook("/home/cells/Projects/cells-toolset/TestData/Input/BookTableData.xlsx")
    #     print("===============")
    #     list_object = workbook.worksheets[0].list_objects[0]
    #     print (list_object.data_range.first_row)
    #     print (list_object.end_row)
    #     for row_index in range(list_object.data_range.first_row ,list_object.end_row ) :
    #         print(row_index)
    #     print("===============")
    #     pass        

    # def test_build_row_dict(self):
    #     # rows ={"up":  { "ipad" : { "2015": 100 , "2016": 3000 } } }
    #     source = [["up","iphone","2012",122],["up","ipad","2012",122],["up","iphone","2013",122] ,["up","ipad","2013",120],["down","iphone","2012",122],["down","ipad","2012",122],["down","iphone","2013",122] ,["down","ipad","2013",120]]
        
    #     # rows = {}
    #     # for old_row in source:
    #     #     cur_row = old_row
    #     #     IsFirstCell = True
    #     #     for cell in old_row :
    #     #         if IsFirstCell : 
    #     #             if cell in rows :
    #     #                 cur_row = self.__build_row_(rows ,cell )
    #     #             else :
    #     #                 rows[cell] ={}
    #     #                 cur_row = rows[cell]
    #     #             IsFirstCell = False
    #     #         else:
    #     #             if cell in cur_row :
    #     #                 cur_row = self.__build_row_(cur_row ,cell )
    #     #             else :
    #     #                 cur_row[cell] ={}
    #     #                 cur_row = cur_row[cell]
    #     # print(source)
    #     # print(rows)

    #     print("=======================================================================")       
    #     rows = {} 
    #     pivot_index = 2
    #     value_index = 3 
    #     column_index = 0
    #     pivot_value = None
    #     value_value = None
    #     for old_row in source:
    #         cur_row = old_row
    #         IsFirstCell = True
    #         column_index = 0
    #         for cell in old_row :
    #             if column_index == pivot_index :
    #                 pivot_value = cell
    #             elif column_index == value_index:
    #                 value_value = cell
    #             else:
    #                 if IsFirstCell : 
    #                     if cell in rows :
    #                         cur_row = self.__build_row_(rows ,cell )
    #                     else :
    #                         rows[cell] ={}
    #                         cur_row = rows[cell]
    #                     IsFirstCell = False
    #                 else:
    #                     if cell in cur_row :
    #                         cur_row = self.__build_row_(cur_row ,cell )
    #                     else :
    #                         cur_row[cell] ={}
    #                         cur_row = cur_row[cell]              
    #             column_index = column_index +1      
    #         cur_row[pivot_value] = value_value
    #     print(rows)            
    #     pass
        
    def __build_row_(self, rows , key ) -> dict: 
        return rows[key]
    

if __name__ == '__main__':
    unittest.main()
