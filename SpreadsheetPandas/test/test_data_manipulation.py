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
from spreadsheetpandas.data_manipulation import pivot_column, unpivot_column
from aspose.cells import Workbook
from aspose.cells import Worksheet
from aspose.cells import Cells
from aspose.cells import Cell
from aspose.cells import CellsHelper


class TestDataManipulation(unittest.TestCase):

    def setUp(self):
        warnings.simplefilter("ignore", ResourceWarning)

    def tearDown(self):
        pass

    def test_pivot_column(self):
        workbook = Workbook("../TestData/Input/BookTableW2L.xlsx")
        table = workbook.worksheets[0].list_objects[0]
        pivot_column(table, "Date", "Value", "")
        pass

    def test_unpivot_culmn(self):
        workbook = Workbook("../TestData/Input/BLogMonthStatData.xlsx")
        table = workbook.worksheets[0].list_objects[0]
        column_namne_list = [
            "3D",
            "BarCode",
            "CAD",
            "Cells",
            "Customer",
            "Diagram",
            "Email",
            "HTML",
            "Imaging",
            "OCR",
            "OMR",
            "PDF",
            "Slides",
            "Tasks",
            "Total",
            "Video",
            "Words",
        ]
        unpivot_list = unpivot_column(
            table,
            column_namne_list,
            "product",
            "blog",
        )
        assert len(unpivot_list) > 2
        assert len(table.list_columns) - len(column_namne_list) + 2 == len(
            unpivot_list[0]
        )
        pass


if __name__ == "__main__":
    unittest.main()
