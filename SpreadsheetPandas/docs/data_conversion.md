**Data Conversion**

Spreadsheet Pandas provides many functions for converting data objects between Aspose.Cells data objects and other popular data objects. For example, converting Excel sheet data to Pandas DataFrame objects, converting Python list data to Excel ListObject data, and more. Spreadsheet Pandas offers a broad range of data conversion features that significantly reduce developers' workload.


# Aspose.Cells basic operations

## How to obtain a worksheet object from an Excel file

Get a worksheet object is simple:

``` Python
from aspose.cells import Workbook 

workbook = Workbook("BookTableData.xlsx")
worksheet = workbook.worksheets.get("SaleSheet")
                                                                                                                                                                                                                                                                                                                                                                                                                                                         
```

## How to obtain a list object from an Excel file

Get a list object is also simple:

``` Python
from aspose.cells import Workbook 

workbook = Workbook("BookTableData.xlsx")
worksheet = workbook.worksheets.get("SaleSheet")
listobject = worksheet.list_objects[0]

```

## How to obtain a name from an Excel file

Get a list object is also simple:

``` Python
from aspose.cells import Workbook 

workbook = Workbook("BookTableData.xlsx")
name = workbook.worksheets.name[0]

```
## How to obtain a cells object from an Excel file

Get a list object is also simple:

``` Python
from aspose.cells import Workbook 

workbook = Workbook("BookTableData.xlsx")
worksheet = workbook.worksheets[0]
cells = worksheet.cells

```

## How to create a range from an Excel file

Get a list object is also simple:

``` Python
from aspose.cells import Workbook 

workbook = Workbook("BookTableData.xlsx")
cells = workbook.worksheets[0].cells
range_ = cells.create_range("A1","D10")
```


# Data Conversion

## How to obtain list data form an Excel worksheet

Suppose you want to process the following excel data :
<div>
  <table>
  <tr>	<td>Product</td><td>Year</td><td>Month</td><td>Sale Number</td><td>Sale Amount </td>  </tr>
  <tr>	<td>iPad</td><td>2023</td><td>10</td><td>120</td><td>420000</td>   </tr>
  <tr>	<td>iPhone</td><td>2023</td><td>10</td><td>120</td><td>780000</td>  </tr>
  </table>
</div>

Get a list object is also simple:

``` Python
from aspose.cells import Workbook 
from spreadsheetpandas.data_conversion import *

workbook = Workbook("BookTableData.xlsx")
worksheet = workbook.worksheets.get("SaleSheet")
data =  worksheet_to_list(worksheet)

```

## How to obtain numpy ndarray data form an Excel worksheet

Suppose you want to process the following excel data :
<div>
  <table>
  <tr>	<td>10</td><td>14</td><td>12</td><td>121</td><td>108</td>  </tr>
  <tr>	<td>50</td><td>23</td><td>10</td><td>120</td><td>420</td>   </tr>
  <tr>	<td>20</td><td>20</td><td>8</td><td>120</td><td>780</td>  </tr>
  </table>
</div>

Get a numpy ndarray object is also simple:

``` Python
from aspose.cells import Workbook 
from spreadsheetpandas.data_conversion import *

workbook = Workbook("BookTableData.xlsx")
worksheet = workbook.worksheets.get("SaleSheet")
data =  worksheet_to_ndarray(worksheet)

```

## How to obtain Pandas DataFrame form an Excel worksheet

Suppose you want to process the following excel data :
<div>
  <table>
  <tr>	<td>Product</td><td>Year</td><td>Month</td><td>Sale Number</td><td>Sale Amount </td>  </tr>
  <tr>	<td>iPad</td><td>2023</td><td>10</td><td>120</td><td>420000</td>   </tr>
  <tr>	<td>iPhone</td><td>2023</td><td>10</td><td>120</td><td>780000</td>  </tr>
  </table>
</div>

Get a Pandas DataFrame is also simple:

``` Python
from aspose.cells import Workbook 
from spreadsheetpandas.data_conversion import *

workbook = Workbook("BookTableData.xlsx")
worksheet = workbook.worksheets.get("SaleSheet")
data =  worksheet_to_dataframe(worksheet)

```

## How to save Pandas DataFrame array an Excel worksheet

Suppose you want to process the following data frame:

``` Python
dataframe_data =pd.DataFrame([['iPad',2023,300],'iPad',2022,600],['iPhone',2022,600],['iPhone',2023,700]] , columns=['Product', 'Year','Sale'])   
```

Save a Pandas DataFrame as Worksheet is simple:

``` Python
from aspose.cells import Workbook 
from spreadsheetpandas.data_conversion import *

dataframe_data =pd.DataFrame([['iPad',2023,300],'iPad',2022,600],['iPhone',2022,600],['iPhone',2023,700]] , columns=)   
workbook = Workbook("BookTableData.xlsx")
worksheet = workbook.worksheets.get("SaleSheet")
dataframe_to_worksheet(dataframe_data,worksheet)

```

## How to save an python list as an Excel table

Suppose you want to process the following data frame:

``` Python
list_data =[['Product', 'Year','Sale'],['iPad',2023,300],'iPad',2022,600],['iPhone',2022,600],['iPhone',2023,700]]
```

Save a Python list as list object is simple:

``` Python
from aspose.cells import Workbook 
from spreadsheetpandas.data_conversion import *
import pandas as pd

list_data =[['Product', 'Year','Sale'],['iPad',2023,300],'iPad',2022,600],['iPhone',2022,600],['iPhone',2023,700]] 
workbook = Workbook("BookTableData.xlsx")
worksheet = workbook.worksheets.get("SaleSheet")
list_to_list_object(dataframe_data,worksheet,begin_row_index=10,begin_column_index=3)

```


## How to save an python list as an Excel Name

Suppose you want to process the following data frame:

``` Python
list_data =[['Product', 'Year','Sale'],['iPad',2023,300],'iPad',2022,600],['iPhone',2022,600],['iPhone',2023,700]]
```

Save a Python list as name is simple:

``` Python
from aspose.cells import Workbook 
from spreadsheetpandas.data_conversion import *
import pandas as pd

list_data =[['Product', 'Year','Sale'],['iPad',2023,300],'iPad',2022,600],['iPhone',2022,600],['iPhone',2023,700]] 
workbook = Workbook("BookTableData.xlsx")
worksheet = workbook.worksheets.get("SaleSheet")
list_to_name(list_data,worksheet)

```




