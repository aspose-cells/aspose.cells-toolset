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



class TestOneCase( unittest.TestCase):
    
    def setUp(self):
        warnings.simplefilter('ignore', ResourceWarning)
    
    def tearDown(self):
        pass
    def test_read_spreadsheet_url(self):
        df = read_spreadsheet('https://docs.aspose.cloud/cells/supported-file-formats/')    
        assert df.shape      
        pass        
    

if __name__ == '__main__':
    unittest.main()
