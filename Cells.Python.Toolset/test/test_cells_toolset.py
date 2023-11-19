from __future__ import absolute_import

import os
from sqlite3 import Row
import sys
import unittest
import warnings

ABSPATH = os.path.abspath(os.path.realpath(os.path.dirname(__file__)) + "/..")
sys.path.append(ABSPATH)


from aspose.cells import Workbook 
from asposecellstoolset.CellsToolset import *
import numpy as np

class TestCellsToolset( unittest.TestCase):
    
    def setUp(self):
        warnings.simplefilter('ignore', ResourceWarning)
    
    def tearDown(self):
        pass
    
    def test_import_data(self):       
        data = np.array(["row","column","table","range","shape","workbook","worksheet","cells","picture"])
        import_data_into_file( "ImportNDArray.xlsx",data )
        pass
        
if __name__ == '__main__':
    unittest.main()
