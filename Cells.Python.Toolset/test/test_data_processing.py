from __future__ import absolute_import

import os
import sys
import unittest
import warnings
import re

ABSPATH = os.path.abspath(os.path.realpath(os.path.dirname(__file__)) + "/..")
sys.path.append(ABSPATH)

from asposecellstoolset.DataProcessing import DataProcessing
from aspose.cells import Workbook 
from aspose.cells import CellsHelper

class TestDataProcessing( unittest.TestCase):
    
    def setUp(self):
        warnings.simplefilter('ignore', ResourceWarning)
    
    def tearDown(self):
        pass

    def test_data_deduplication(self):
        workbook = Workbook("d:\cells-toolset\TestData\Input\BookCsvDuplicateData.csv")
        data_processing = DataProcessing()  
        data_processing.data_deduplication(workbook)
        workbook.save("d:\cells-toolset\TestData\Output\BookCsvDuplicateData.xlsx")
        pass
    
if __name__ == '__main__':
    unittest.main()
