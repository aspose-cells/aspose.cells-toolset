
# Data Access 


## Work with Excel files

### **Get data from Excel file or other format spreadsheet files.**

Read a csv file is simple:

``` Python
import spreadsheetpandas
data_frame = read_spreadsheet("example.csv")

```

Read an Excel file is also simple:

``` Python
import spreadsheetpandas
data = read_spreadsheet("example.xlsx")

```

The same applies to a xls file:

``` Python
import spreadsheetpandas
data = read_spreadsheet("example.xls")

```

### **How to get specific data from Excel files?** 

It's a little more complicated to point out the specific location, such as getting the data for a ListObject

``` Python
import spreadsheetpandas
data = read_spreadsheet("example.xls",sheet_index=>0, list_object_index=>0)

```

or 

``` Python
import spreadsheetpandas
data = read_spreadsheet("example.xls",list_object_name=>"statistics")

```

### **Save data as Excel files**


Saving an Excel file is still simple:

``` Python

import spreadsheetpandas
import pandas as pd 
array = [[1, 2, 3], [4, 5, 6], [7, 8, 9]]
write_spreadsheet("example.xlsx"，pd.DataFrame(array) )

```

### **How to save data to a specific position of an Excel file?** 

It's a little more complicated to point out the specific location, such as begin row and column.

``` Python
import spreadsheetpandas
import pandas as pd 
array = [[1, 2, 3], [4, 5, 6], [7, 8, 9]]
write_spreadsheet("example.xlsx"，pd.DataFrame(array) , begin_row_index=>3, begin_column_index=>4)

```


## Work with Html files

### **Get data from Html file or Uri.**

Read a Html file is simple:

``` Python
import spreadsheetpandas
data_frame = read_spreadsheet("example.html")

```
Read a Html content form uri is simple: 
``` Python
import spreadsheetpandas
data_frame = read_spreadsheet("https://docs.aspose.cloud/cells/supported-file-formats/")

```


## **Only get table data from Html files or uri?**

# Support Data File Format

|**Format**|**Description**|**Load**|**Save**|
| :- | :- | :- | :- |
|[XLS](https://docs.fileformat.com/spreadsheet/xls/)|Excel 95/5.0 - 2003 Workbook.|&radic;|&radic;|
|[XLSX](https://docs.fileformat.com/spreadsheet/xlsx/)|Office Open XML SpreadsheetML Workbook or template file, with or without macros.|&radic;|&radic;|
|[XLSB](https://docs.fileformat.com/spreadsheet/xlsb/)|Excel Binary Workbook.|&radic;|&radic;|
|[XLSM](https://docs.fileformat.com/spreadsheet/xlsm/)|Excel Macro-Enabled Workbook.|&radic;|&radic;|
|[XLT](https://docs.fileformat.com/spreadsheet/xlt/)|Excel 97 - Excel 2003 Template.|&radic;|&radic;|
|[XLTX](https://docs.fileformat.com/spreadsheet/xltx/)|Excel Template.|&radic;|&radic;|
|[XLTM](https://docs.fileformat.com/spreadsheet/xltm/)|Excel Macro-Enabled Template.|&radic;|&radic;|
|[XLAM](https://docs.fileformat.com/spreadsheet/xlam/)|An Excel Macro-Enabled Add-In file that's used to add new functions to Excel.| |&radic;|
|[CSV](https://docs.fileformat.com/spreadsheet/csv/)|CSV (Comma Separated Value) file.|&radic;|&radic;|
|[TSV](https://docs.fileformat.com/spreadsheet/tsv/)|TSV (Tab-separated values) file.|&radic;|&radic;|
|TabDelimited|Tab-delimited text file, same with TSV file.|&radic;|&radic;|
|[TXT](https://docs.fileformat.com/word-processing/txt/)|Delimited plain text file.|&radic;|&radic;|
|[HTML](https://docs.fileformat.com/web/html/)|HTML format.|&radic;|&radic;|
|[MHTML](https://docs.fileformat.com/web/mhtml/)|MHTML file.|&radic;|&radic;|
|[ODS](https://docs.fileformat.com/spreadsheet/ods/)|ODS (OpenDocument Spreadsheet).|&radic;|&radic;|
|SpreadsheetML|Excel 2003 XML file.|&radic;|&radic;|
|[Numbers](https://docs.fileformat.com/spreadsheet/numbers/)|The document is created by Apple's "Numbers" application which forms part of Apple's iWork office suite, a set of applications which run on the Mac OS X and iOS operating systems.|&radic;||
|[JSON](https://docs.fileformat.com/web/json/)|JavaScript Object Notation|&radic;|&radic;|
|[DIF](https://docs.fileformat.com/spreadsheet/dif/)|Data Interchange Format.| |&radic;|
|[PDF](https://docs.fileformat.com/pdf/)|Adobe Portable Document Format.| |&radic;|
|[XPS](https://docs.fileformat.com/page-description-language/xps/)|XML Paper Specification Format.| |&radic;|
|[SVG](https://docs.fileformat.com/page-description-language/svg/)|Scalable Vector Graphics Format.| |&radic;|
|[TIFF](https://docs.fileformat.com/image/tiff/)|Tagged Image File Format| |&radic;|
|[PNG](https://docs.fileformat.com/image/png/)|Portable Network Graphics Format| |&radic;|
|[BMP](https://docs.fileformat.com/image/bmp/)|Bitmap Image Format| |&radic;|
|[EMF](https://docs.fileformat.com/image/emf/)|Enhanced metafile Format| |&radic;|
|[JPEG](https://docs.fileformat.com/image/jpeg/)|JPEG is a type of image format that is saved using the method of lossy compression.| |&radic;|
|[GIF](https://docs.fileformat.com/image/gif/)|Graphical Interchange Format| |&radic;|
|[MARKDOWN](https://docs.fileformat.com/word-processing/md/)|Represents a markdown document.| |&radic;|
|[SXC](https://docs.fileformat.com/spreadsheet/sxc/)|An XML based format used by OpenOffice and StarOffice|&radic;|&radic;|
|[FODS](https://docs.fileformat.com/spreadsheet/fods/)|This is an Open Document format stored as flat XML.|&radic;|&radic;|
|[DOCX](https://docs.fileformat.com/word-processing/docx/)|A well-known format for Microsoft Word documents that is a combination of XML and binary files.||&radic;|
|[PPTX](https://docs.fileformat.com/presentation/pptx/)|The PPTX format is based on the Microsoft PowerPoint open XML presentation file format.||&radic;|
