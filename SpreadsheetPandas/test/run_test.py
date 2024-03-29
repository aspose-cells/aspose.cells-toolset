import io
import os
import unittest
import xmlrunner

if __name__ == '__main__':
    case_path = os.getcwd()
    discover = unittest.defaultTestLoader.discover(case_path,pattern="test_*.py")    
    testRunner=xmlrunner.XMLTestRunner(output='test-reports')
    testRunner.run(discover)
       
