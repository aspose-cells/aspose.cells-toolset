# Spreadsheet Processing Toolset Python High Code API

[Product Page](https://products.aspose.com/cells/python-net/) | [Docs](https://docs.aspose.com/cells/python-net/) | [Demos](https://products.aspose.app/cells/family/) | [API Reference](https://reference.aspose.com/cells/python-net/) | [Examples](https://github.com/aspose-cells/aspose.cells-toolset) | [Blog](https://blog.aspose.com/category/cells/) | [Free Support](https://forum.aspose.com/c/cells) | [Temporary License](https://purchase.aspose.com/temporary-license)

[Aspose.Cells for Python via .NET](https://products.aspose.com/cells/python-net/) is a scalable and feature-rich API to process Excel&reg; spreadsheets using Python. API offers Excel&reg; file creation, manipulation, conversion and rendering. Developers can format worksheets, rows, columns or cells to the most granular level, create manipulate chart and pivot tables, render worksheets, charts and specific data ranges to PDF or images, add calculate Excel&reg;'s built-in and custom formulas and much more - all without any dependency on Microsoft Office or Excel&reg; application.

## Spreadsheet Python via .NET On-premise API Features

- Spreadsheet generation & manipulation via API.
- High-quality file format conversion & rendering.
- Print Microsoft Excel&reg; files to physical or virtual printers.
- Combine, modify, protect, or parse Excel&reg; sheets.
- Apply worksheet formatting.
- Configure and apply page setup for the worksheets.
- Create & customize Excel&reg; charts, Pivot Tables, conditional
  formatting rules, slicers, tables & spark-lines.
- Convert Excel&reg; charts to images & PDF.
- Convert Excel&reg; files to various other formats.
- Formula calculation engine that supports all basic and advanced Excel&reg; functions.

Please visit the [official documentation](https://docs.aspose.com/cells/python-net/) for a more detailed list of features.

## Read & Write Sreadsheet File Formats

**Microsoft Excel&reg;:** XLS, XLSX, XLSB, XLSM, XLT, XLTX, XLTM, CSV, TSV, TabDelimited, SpreadsheetML\
**OpenOffice:** ODS, SXC, FODS\
**Text:** TXT\
**Web:** HTML, MHTML\
**iWork&reg;:** Numbers\
**Other:** SXC, FODS

## Save Spreadsheet Files AS

**Microsoft Word&reg;:** DOCX\
**Microsoft PowerPoint&reg;:** PPTX\
**Microsoft Excel&reg;:** XLAM\
**Fixed Layout:** PDF, XPS\
**Data Interchange:** DIF\
**Vector Graphics:** SVG\
**Image:** TIFF,PNG, BMP, JPEG, GIF\
**Meta File:** EMF\
**Markdown:** MD

Please visit [Supported File Formats](https://docs.aspose.com/cells/python-net/supported-file-formats/) for further details.

## System Requirements

Your machine does not need to have Microsoft Excel&reg; or OpenOffice&reg; software installed.

### Supported Operating Systems

**Microsoft Windows&reg;:** Windows Desktop & Server (`x64`, `x86`)\
**Linux:** Ubuntu, OpenSUSE, CentOS, and others\
**Other:** Any operating system (OS) that can install Mono(.NET 4.0 Framework support) or use .NET core.

## Get Started

### Installation via `pip`

The Aspose.Cells for Python via .NET is [available at pypi.org](https://pypi.org/project/aspose-cells-python/). To install it, please run the following command:

`pip install aspose-cells-python`

The pandas is [available at pypi.org](https://pypi.org/project/pandas/). To install it, please run the following command:

`pip install pandas`

The numpy is [available at pypi.org](https://pypi.org/project/numpy/). To install it, please run the following command:

`pip install numpy`

### Import Numpy array to excel using aspose.cells-toolset

```python

#import the python package
import numpy as np
import aspose.cells
from aspose.cells import License,Workbook,FileFormatType

import_tool = CellsImportUtility()
workbook = Workbook()
data = np.array([1,2,3,4,5,6,7,8,9])
       
import_tool.import_data_into_workbook( workbook ,data, is_vertical=True)
workbook.save("import_numarray_int_vertical.xlsx")

```

## Import Pandas DataFrame to excel using aspose.cells-toolset

```python
#import the python package
import numpy as np
import pandas as pd
import aspose.cells
from aspose.cells import License,Workbook,FileFormatType

dates = pd.date_range("20130101", periods=6)
df = pd.DataFrame(np.random.randn(6, 4), index=dates, columns=list("ABCD"))
import_tool = CellsImportUtility()
workbook = Workbook() 
import_tool.import_data_into_workbook( workbook ,df, is_vertical=True)
workbook.save("import_dataframe_vertical.xlsx")

```

[Product Page](https://products.aspose.com/cells/python-net) | [Docs](https://docs.aspose.com/cells/python-net/) | [Demos](https://products.aspose.app/cells/family/) | [API Reference](https://reference.aspose.com/cells/python-net/) | [Examples](https://github.com/aspose-cells/aspose.cells-toolset) | [Blog](https://blog.aspose.com/category/cells/) | [Free Support](https://forum.aspose.com/c/cells) | [Temporary License](https://purchase.aspose.com/temporary-license)