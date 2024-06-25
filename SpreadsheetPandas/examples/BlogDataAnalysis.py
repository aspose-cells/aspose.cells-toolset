from __future__ import absolute_import
from calendar import c

import os
import sys
import unittest
import warnings
import re
from numpy import ndarray
import pandas as pd
from pandas.plotting import table


ABSPATH = os.path.abspath(os.path.realpath(os.path.dirname(__file__)) + "/..")
sys.path.append(ABSPATH)
from spreadsheetpandas.data_conversion import *
from spreadsheetpandas.spreadsheet_pandas import read_spreadsheet
from spreadsheetpandas.data_manipulation import pivot_column
from aspose.cells import Workbook
from aspose.cells import Worksheet
from aspose.cells.charts import ChartType
from aspose.cells import CellsHelper

def data_cleansing(table: ListObject):
    data_dict = {
        "Aspose.3D Cloud Product Family": "3D",
        "Aspose.BarCode Cloud Product Family": "BarCode",
        "Aspose.CAD Cloud Product Family": "CAD",
        "Aspose.Cells Cloud Product Family": "Cells",
        "Aspose.Diagram Cloud Product Family": "Diagram",
        "Aspose.Email Cloud Product Family": "Email",
        "Aspose.HTML Cloud Product Family": "HTML",
        "Aspose.Imaging Cloud Product Family": "Imaging",
        "Aspose.OCR Cloud Product Family": "OCR",
        "Aspose.OMR Cloud Product Family": "OMR",
        "Aspose.PDF Cloud Product Family": "PDF",
        "Aspose.Pdf Cloud Product Family": "PDF",
        "Aspose.Slides Cloud Product Family": "Slides",
        "Aspose.Tasks Cloud Product Family": "Tasks",
        "Aspose.Total Cloud Product Family": "Total",
        "Aspose.Video Cloud Product Family": "Video",
        "Aspose.Words Cloud Product Family": "Words",
        "Customer Newsletters": "Customer",
        "Aspose.3D Cloud Product Family": "3D",
        "Aspose.BarCode Cloud Product Family": "BarCode",
        "Aspose.CAD Cloud Product Family": "CAD",
        "Aspose.Cells Cloud Product Family": "Cells",
        "Aspose.Diagram Cloud Product Family": "Diagram",
        "Aspose.Email Cloud Product Family": "Email",
        "Aspose.HTML Cloud Product Family": "HTML",
        "Aspose.Imaging Cloud Product Family": "Imaging",
        "Aspose.OCR Cloud Product Family": "OCR",
        "Aspose.OMR Cloud Product Family": "OMR",
        "Aspose.PDF Cloud Product Family": "PDF",
        "Aspose.Pdf Cloud Product Family": "PDF",
        "Aspose.Slides Cloud Product Family": "Slides",
        "Aspose.Tasks Cloud Product Family": "Tasks",
        "Aspose.Total Cloud Product Family": "Total",
        "Aspose.Video Cloud Product Family": "Video",
        "Aspose.Words Cloud Product Family": "Words",
        "Customer Newsletters": "Customer",
    }
    cells = table.data_range.worksheet.cells
    table_data_begin_row_index =  table.data_range.first_row 
    table_data_end_row_index =  table.data_range.first_row  +  table.data_range.row_count
    table_data_begin_column_index =  table.data_range.first_column  
    table_data_end_column_index =  table.data_range.first_column  +  table.data_range.column_count
    for row_index in range( table_data_begin_row_index, table_data_end_row_index):        
        cur_pivot_column_value = cells.get(row_index,1).value    
        if cur_pivot_column_value in data_dict:
            cells.get(row_index,1).put_value(data_dict[cur_pivot_column_value])
    pass

def dynamic_chart(table : ListObject , sheet :Worksheet):
    source_cells = table.data_range.worksheet.cells
    table_data_begin_row_index =  table.data_range.first_row 
    table_data_end_row_index =  table.data_range.first_row  +  table.data_range.row_count
    table_data_begin_column_index =  table.data_range.first_column  
    table_data_end_column_index =  table.data_range.first_column  +  table.data_range.column_count    
    target_cells = sheet.cells
    target_cells.get(0,0).put_value("Date")
    target_cells.get(0,1).formula = "=OFFSET(ProductsMonthStat!B1:B1,0,ProductMonthStat!I1)"
    for row_index in range(table_data_begin_row_index,table_data_end_row_index):
        target_cells.get(row_index,0).put_value(source_cells.get(row_index,0).value)
        target_cells.get(row_index,1).formula = "=OFFSET(ProductsMonthStat!B{0}:B{0},0,ProductMonthStat!I1)".format(row_index+1)
    pass

workbook = Workbook("../TestData/Input/BLogData.xlsx")

new_sheet_index = workbook.worksheets.add_copy(0)
new_sheet =  workbook.worksheets[new_sheet_index]
new_sheet.name = 'base_data_2'
table1 = new_sheet.list_objects[0]
data_cleansing(table1)
list1 = pivot_column(
    table1,
    "Categories",
    "",
    "",
    out_fields=["Date"],
    date_to_string_fields={"Date": "%Y-%m"},
)

data_stat_product =  workbook.worksheets.add("ProductsMonthStat")
products_month_stat_table = list_to_list_object(list1 , data_stat_product)

product_month_sheet =  workbook.worksheets.add("ProductMonthStat")
dynamic_chart(products_month_stat_table , product_month_sheet)
dashboard_sheet =  workbook.worksheets.add("Dashboard")
dashboard_sheet.is_gridlines_visible = False
workbook.worksheets.active_sheet_index = dashboard_sheet.index
chartIndex = dashboard_sheet.charts.add( ChartType.LINE, "ProductMonthStat!$A$1:$B$147", True, 1, 0, 27, 14);
column_index = 0
row_index = 3
dashboard_sheet.cells.get(row_index,20).put_value("Product") 
dashboard_sheet.cells.get(row_index,21).put_value("SUM") 
dashboard_sheet.cells.get(row_index,22).put_value("AVERAGE") 
dashboard_sheet.cells.get(row_index,23).put_value("STDEV.P") 
dashboard_sheet.cells.get(row_index,24).put_value("STDEV.S") 
dashboard_sheet.cells.get(row_index,25).put_value("MIN") 
dashboard_sheet.cells.get(row_index,26).put_value("MEDIAN") 
dashboard_sheet.cells.get(row_index,27).put_value("MAX") 
row_index = row_index + 1

for column in products_month_stat_table.list_columns:
    if  column_index == 0: 
        column_index = column_index + 1
        continue
    dashboard_sheet.cells.get(28,column_index).put_value(column.name)
    dashboard_sheet.cells.set_column_width(column_index, 6.71);
    column_index = column_index + 1
    dashboard_sheet.cells.get(row_index,20).put_value(column.name) 
    dashboard_sheet.cells.get(row_index,21).formula = "=SUM({0}[{1}])".format(products_month_stat_table.display_name , column.name)
    dashboard_sheet.cells.get(row_index,22).formula = "=AVERAGE({0}[{1}])".format(products_month_stat_table.display_name , column.name)
    dashboard_sheet.cells.get(row_index,23).formula = "=STDEV.P({0}[{1}])".format(products_month_stat_table.display_name , column.name)
    dashboard_sheet.cells.get(row_index,24).formula = "=STDEV.S({0}[{1}])".format(products_month_stat_table.display_name , column.name)
    dashboard_sheet.cells.get(row_index,25).formula = "=MIN({0}[{1}])".format(products_month_stat_table.display_name , column.name)
    dashboard_sheet.cells.get(row_index,26).formula = "=MEDIAN({0}[{1}])".format(products_month_stat_table.display_name , column.name)
    dashboard_sheet.cells.get(row_index,27).formula = "=MAX({0}[{1}])".format(products_month_stat_table.display_name , column.name)
    row_index = row_index + 1

scrollBar =  dashboard_sheet.shapes.add_scroll_bar(29, 0, 0,50, 20,950);
scrollBar.current_value = 0;
scrollBar.min = 0;
scrollBar.max = len(list1[0]) - 2;
scrollBar.incremental_change = 1;
scrollBar.page_change = 1;
scrollBar.linked_cell = "ProductMonthStat!$I$1";
scrollBar.is_horizontal = True;

workbook.save("../TestData/Output/BLogData.xlsx")

