
import matplotlib.pyplot as plt
import matplotlib.axes 
import numpy as np
import os
import sys
import unittest
import warnings
import re


class TestReMatch( unittest.TestCase):
    
    def setUp(self):
       
        warnings.simplefilter('ignore', ResourceWarning)
    
    def tearDown(self):
        pass
    
    def test_simply(self):
        value = "=SaleChart!$D$7A:$D$26"
        partten = "r'^=(w+)!\$(w+)\$(d+):\$(w+)\$(d+)'"
        # matchObj = re.match( partten, value, re.M|re.I)
        matchObj = re.match( r'^=(.*)!\$(.*)\$(\d+):\$(.*)\$(\d+)', value, re.M|re.I)
        print(matchObj )
        # print(matchObj.group(1) )
        # print(matchObj.group(2) )
        # print(matchObj.group(3) )
        # print(matchObj.group(4) )
        # print(matchObj.group(5) )
        pass
        
   
    
if __name__ == '__main__':
    unittest.main()
        